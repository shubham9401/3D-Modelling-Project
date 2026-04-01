"""
Validator: Two-tier validation system for SolidWorks models.

TIER 1 (SCORED): Dimensions, Features, Structure — always reliable.
TIER 2 (CONDITIONAL): Volume/SA — scored only for simple shapes (Python math).
INFO SECTION: Volume/SA for complex shapes + design checks (shown, not scored).

Score is based ONLY on what can be reliably verified.
"""

import json
import math


# ============================================================
# SPEC EXTRACTION (via LLM)
# ============================================================

def create_spec_from_prompt(user_prompt):
    """
    Sends the user prompt to the LLM to extract expected design specs.
    Returns a dict with expected dimensions, features, etc.
    """
    from validation_prompt import VALIDATION_SYSTEM_PROMPT
    from llm_client import _call_llm

    print("📐 Extracting expected specs from prompt...")

    try:
        raw = _call_llm(VALIDATION_SYSTEM_PROMPT, user_prompt)

        # Clean and parse JSON
        clean = raw.replace("```json", "").replace("```", "").strip()
        start = clean.find('{')
        end = clean.rfind('}')
        if start == -1 or end == -1:
            raise ValueError("No JSON object found in LLM response")
        spec = json.loads(clean[start:end+1])
        return spec

    except Exception as e:
        print(f"❌ Failed to extract specs: {e}")
        return None


# ============================================================
# DIMENSION CHECK
# ============================================================

def _check_dimension(name, expected, actual, tolerance_pct):
    """Compare a single dimension. Returns (score, status, detail)."""
    if expected is None:
        return 100, "SKIP", f"{name}: not specified in prompt"

    if actual is None:
        return 0, "FAIL", f"{name}: could not measure"

    diff = abs(actual - expected)
    tolerance = expected * (tolerance_pct / 100)

    if diff <= tolerance * 0.5:
        return 100, "PASS", f"{name}: {actual}mm (expected {expected}mm)"
    elif diff <= tolerance:
        pct_off = round(diff / expected * 100, 1)
        return 70, "WARN", f"{name}: {actual}mm (expected {expected}mm, off by {pct_off}%)"
    else:
        pct_off = round(diff / expected * 100, 1)
        return 0, "FAIL", f"{name}: {actual}mm (expected {expected}mm, off by {pct_off}%)"


# ============================================================
# FEATURE CHECK
# ============================================================

def _check_features(expected_features, actual_feature_types):
    """Check if expected features are present. Returns (score, checks)."""
    if not expected_features:
        return 100, []

    checks = []
    found = 0

    TYPE_ALIASES = {
        "Extrude": ["ICE", "Extrusion", "Boss-Extrude", "Extrude"],
        "Cut": ["ICE", "Cut-Extrude", "Cut", "CutExtrude"],
        "Shell": ["Shell", "ShellFeature"],
        "Fillet": ["Fillet", "ConstRadiusFillet"],
        "Chamfer": ["Chamfer", "ChamferFeature"],
        "Revolve": ["Revolution", "Revolve", "BossRevolve"],
        "Loft": ["Loft", "LoftFeature"],
        "Sweep": ["Sweep", "SweepFeature"],
        "CircularPattern": ["CirPattern", "CircularPattern"],
        "LinearPattern": ["LPattern", "LinearPattern"],
        "Hole": ["HoleWzd", "Hole", "HoleWizard"],
        "Sketch": ["ProfileFeature", "3DProfileFeature"],
        "Thread": ["Thread", "CosmeticThread"],
        "Mirror": ["MirrorPattern", "Mirror"],
        "Rib": ["Rib", "RibFeature"],
    }

    actual_types_lower = {k.lower(): v for k, v in actual_feature_types.items()}

    for expected_feat in expected_features:
        aliases = TYPE_ALIASES.get(expected_feat, [expected_feat])
        matched = False

        for alias in aliases:
            if alias in actual_feature_types or alias.lower() in actual_types_lower:
                matched = True
                break

        if matched:
            found += 1
            checks.append(("PASS", f"Feature '{expected_feat}': found"))
        else:
            checks.append(("FAIL", f"Feature '{expected_feat}': MISSING"))

    score = round((found / len(expected_features)) * 100) if expected_features else 100
    return score, checks


# ============================================================
# VOLUME / SURFACE AREA (Python formulas for simple shapes)
# ============================================================

def _compute_expected_volume(shape, w, h, d):
    """Compute expected volume using Python math. Returns (volume, formula_name) or (None, None)."""
    if shape == "box":
        return w * h * d, "W×H×D"
    elif shape == "cylinder":
        r = w / 2
        return math.pi * r * r * h, "π×r²×H"
    elif shape == "sphere":
        r = w / 2
        return (4/3) * math.pi * r**3, "4/3×π×r³"
    elif shape == "cone":
        r = w / 2
        return (1/3) * math.pi * r * r * h, "1/3×π×r²×H"
    return None, None


def _compute_expected_sa(shape, w, h, d):
    """Compute expected surface area using Python math. Returns (sa, formula_name) or (None, None)."""
    if shape == "box":
        return 2 * (w*h + w*d + h*d), "2(WH+WD+HD)"
    elif shape == "cylinder":
        r = w / 2
        return 2 * math.pi * r * (r + h), "2πr(r+H)"
    elif shape == "sphere":
        r = w / 2
        return 4 * math.pi * r * r, "4πr²"
    elif shape == "cone":
        r = w / 2
        slant = math.sqrt(r*r + h*h)
        return math.pi * r * (r + slant), "πr(r+s)"
    return None, None


# ============================================================
# MAIN VALIDATION LOGIC — TWO-TIER SYSTEM
# ============================================================

def validate_model(spec, actual_properties):
    """
    Two-tier validation:
      TIER 1 (always scored): Dimensions (30%), Features (30%), Structure (15%)
      TIER 2 (scored if computable): Volume (15%), Surface Area (10%)
      INFO (never scored): Volume/SA for custom shapes, design checks

    Only scores what can be reliably verified.
    Reports everything for human review.
    """
    tolerance = spec.get("tolerance_percent", 10)
    scored_checks = []     # Checks that affect the score
    info_items = []        # Reported but NOT scored
    deviations = []

    # ════════════════════════════════════════════════════════
    # TIER 1: ALWAYS SCORED
    # ════════════════════════════════════════════════════════

    # ── 1A. Bounding Box Dimensions (30%) ──
    dim_scores = []
    expected_dims = spec.get("expected_dimensions", {})
    actual_dims = actual_properties.get("dimensions", {})
    shape = spec.get("expected_shape", "custom")

    # For custom shapes, use wider tolerance because expected dims
    # are inferred defaults (not user-specified numbers)
    dim_tolerance = tolerance
    if shape == "custom":
        dim_tolerance = max(tolerance, 30)  # At least 30% for inferred dims

    for dim_name, dim_key in [("Width", "width"), ("Height", "height"), ("Depth", "depth")]:
        expected_val = expected_dims.get(dim_key)
        actual_val = actual_dims.get(dim_key)
        score, status, detail = _check_dimension(dim_name, expected_val, actual_val, dim_tolerance)

        if status != "SKIP":
            dim_scores.append(score)
            scored_checks.append({"check": dim_name, "expected": expected_val, "actual": actual_val, "status": status})
            if status == "FAIL":
                deviations.append(detail)

    dim_avg = sum(dim_scores) / len(dim_scores) if dim_scores else 100

    # ── 1B. Feature Completeness (30%) ──
    expected_features = spec.get("expected_features", [])
    feature_tree_available = actual_properties.get("feature_tree_available", True)

    if not feature_tree_available and actual_properties.get("feature_count", 0) > 0:
        feat_score = 100
        scored_checks.append({
            "check": "Features",
            "expected": ", ".join(expected_features) if expected_features else "N/A",
            "actual": f"{actual_properties['feature_count']} features",
            "status": "PASS",
        })
    else:
        feat_score, feat_checks = _check_features(expected_features, actual_properties.get("feature_types", {}))

        for status, detail in feat_checks:
            feat_name = detail.split("'")[1] if "'" in detail else detail
            scored_checks.append({
                "check": f"Feature: {feat_name}",
                "expected": "Yes",
                "actual": "Yes" if status == "PASS" else "No",
                "status": status,
            })
            if status == "FAIL":
                deviations.append(detail)

    # ── 1C. Structure: Body & Face Count (15%) ──
    expected_bodies = spec.get("expected_body_count", 1)
    actual_bodies = actual_properties.get("body_count", 0)
    actual_faces = actual_properties.get("face_count", 0)

    body_match = actual_bodies == expected_bodies
    struct_score = 100 if body_match else 50

    scored_checks.append({
        "check": "Body Count",
        "expected": expected_bodies,
        "actual": actual_bodies,
        "status": "PASS" if body_match else "FAIL",
    })
    if not body_match:
        deviations.append(f"Body count: expected {expected_bodies}, got {actual_bodies}")

    # Face count: always show as info (LLM estimates are unreliable)
    info_items.append({
        "check": "Face Count",
        "actual": actual_faces,
        "note": f"{actual_faces} faces detected",
    })

    # ════════════════════════════════════════════════════════
    # TIER 2: SCORED ONLY FOR SIMPLE SHAPES
    # ════════════════════════════════════════════════════════

    shape = spec.get("expected_shape", "custom")
    has_all_dims = expected_dims.get("width") and expected_dims.get("height") and expected_dims.get("depth")
    actual_vol = actual_properties.get("volume_mm3")
    actual_sa = actual_properties.get("surface_area_mm2")

    volume_score = None   # None = not scored, weight redistributed
    sa_score = None

    if shape in ("box", "cylinder", "sphere", "cone") and has_all_dims:
        # SIMPLE SHAPE → Python computes exact expected values → score them
        w = expected_dims["width"]
        h = expected_dims["height"]
        d = expected_dims["depth"]

        # Volume
        expected_vol, vol_formula = _compute_expected_volume(shape, w, h, d)
        if actual_vol and expected_vol and expected_vol > 0:
            vol_error = abs(1 - actual_vol / expected_vol) * 100
            if vol_error <= 5:
                volume_score = 100
                vol_status = "PASS"
            elif vol_error <= 15:
                volume_score = 80
                vol_status = "PASS"
            elif vol_error <= 30:
                volume_score = 50
                vol_status = "WARN"
            else:
                volume_score = 0
                vol_status = "FAIL"

            scored_checks.append({
                "check": "Volume",
                "expected": f"{round(expected_vol, 1)} ({vol_formula})",
                "actual": actual_vol,
                "status": vol_status,
            })
            if vol_status in ("WARN", "FAIL"):
                deviations.append(f"Volume: expected {round(expected_vol,1)}mm³, got {actual_vol}mm³ (±{vol_error:.1f}%)")

        # Surface Area
        expected_sa, sa_formula = _compute_expected_sa(shape, w, h, d)
        if actual_sa and expected_sa and expected_sa > 0:
            sa_error = abs(1 - actual_sa / expected_sa) * 100
            if sa_error <= 10:
                sa_score = 100
                sa_status = "PASS"
            elif sa_error <= 25:
                sa_score = 70
                sa_status = "PASS"
            else:
                sa_score = 0
                sa_status = "FAIL"

            scored_checks.append({
                "check": "Surface Area",
                "expected": f"{round(expected_sa, 1)} ({sa_formula})",
                "actual": actual_sa,
                "status": sa_status,
            })
            if sa_status in ("WARN", "FAIL"):
                deviations.append(f"Surface Area: expected {round(expected_sa,1)}mm², got {actual_sa}mm²")
    else:
        # CUSTOM/COMPLEX → show volume/SA as INFO, don't score
        if actual_vol:
            info_items.append({
                "check": "Volume",
                "actual": f"{actual_vol:,.1f} mm³",
                "note": "Not scored (complex shape — no formula to verify against)",
            })
        if actual_sa:
            info_items.append({
                "check": "Surface Area",
                "actual": f"{actual_sa:,.1f} mm²",
                "note": "Not scored (complex shape — no formula to verify against)",
            })

    # ════════════════════════════════════════════════════════
    # INFO SECTION: Design Checks & Specific Measurements
    # ════════════════════════════════════════════════════════

    # Add LLM-extracted specific measurements (for human review)
    specific_measurements = spec.get("specific_measurements", [])
    for measurement in specific_measurements:
        info_items.append({
            "check": "Spec",
            "actual": measurement,
            "note": "From prompt — verify visually",
        })

    # Add design checks (for human review)
    design_checks = spec.get("design_checks", [])
    for check in design_checks:
        info_items.append({
            "check": "Design",
            "actual": check,
            "note": "Verify visually in SolidWorks",
        })

    # ════════════════════════════════════════════════════════
    # FINAL SCORE — only from scored checks
    # ════════════════════════════════════════════════════════

    weights = {
        "dim": 0.30,
        "feat": 0.30,
        "struct": 0.15,
        "vol": 0.15,
        "sa": 0.10,
    }

    scores = {
        "dim": dim_avg,
        "feat": feat_score,
        "struct": struct_score,
        "vol": volume_score,      # None if not scored
        "sa": sa_score,           # None if not scored
    }

    # Redistribute weight from unskored categories
    active = {k: v for k, v in scores.items() if v is not None}
    if active:
        total_weight = sum(weights[k] for k in active)
        final_score = sum(scores[k] * (weights[k] / total_weight) for k in active)
    else:
        final_score = 0

    return {
        "score": round(final_score),
        "scored_checks": scored_checks,
        "info_items": info_items,
        "deviations": deviations,
        "scores_breakdown": {k: round(v) if v is not None else "N/A" for k, v in scores.items()},
    }


def run_validation(user_prompt):
    """Full pipeline: inspect model → extract specs → compare → report."""
    try:
        # Step 1: Inspect the current model
        print("\n🔍 Inspecting current model...")
        from tools.model_inspector import get_model_properties
        actual = get_model_properties()

        if actual is None:
            print("❌ Could not inspect model")
            return None

        dims = actual["dimensions"]
        print(f"   ✅ Model: {dims['width']}W × {dims['height']}H × {dims['depth']}D mm, "
              f"{actual['body_count']} bodies, {actual['face_count']} faces, "
              f"{actual['feature_count']} features")

        # Step 2: Extract specs from prompt
        spec = create_spec_from_prompt(user_prompt)
        if spec is None:
            print("   ❌ Could not extract specs from prompt")
            return None

        print(f"   📐 Expected: {spec.get('description', 'N/A')}")

        # Step 3: Compare
        print("\n📊 Comparing actual model vs expected specs...")
        result = validate_model(spec, actual)

        # Step 4: Print summary to terminal
        _print_report(user_prompt, result)

        # Step 5: Save detailed HTML report
        _save_html_report(user_prompt, result, spec, actual)

        # Step 6: Save JSON (machine readable)
        _save_json_report(user_prompt, result, spec, actual)

        return result

    except Exception as e:
        print(f"\n❌ Validation error: {e}")
        import traceback
        traceback.print_exc()
        return None


# ============================================================
# REPORT OUTPUT
# ============================================================

def _print_report(prompt, result):
    """Print a concise summary to terminal (full details in HTML report)."""
    score = result["score"]

    if score >= 85:
        grade = "✅ EXCELLENT"
    elif score >= 70:
        grade = "✅ GOOD"
    elif score >= 50:
        grade = "⚠️  NEEDS WORK"
    else:
        grade = "❌ POOR"

    print("\n" + "═" * 60)
    print(f"  SCORE: {score}/100  ({grade})")
    print("═" * 60)

    for c in result["scored_checks"]:
        icon = {"PASS": "✅", "WARN": "⚠️ ", "FAIL": "❌"}.get(c["status"], "  ")
        print(f"  {icon} {c['check']}: {c.get('actual', 'N/A')}")

    if result["deviations"]:
        for dev in result["deviations"]:
            print(f"  ❌ {dev}")

    print("═" * 60)


def _save_json_report(prompt, result, spec, actual):
    """Save machine-readable JSON report."""
    report = {
        "prompt": prompt,
        "score": result["score"],
        "status": "PASS" if result["score"] >= 70 else "FAIL",
        "expected_spec": spec,
        "actual_properties": actual,
        "scored_checks": result["scored_checks"],
        "info_items": result["info_items"],
        "deviations": result["deviations"],
        "scores_breakdown": result["scores_breakdown"],
    }
    with open("validation_report.json", "w") as f:
        json.dump(report, f, indent=2)


def _save_html_report(prompt, result, spec, actual):
    """Generate a comprehensive, detailed HTML validation report."""
    from datetime import datetime

    score = result["score"]
    if score >= 85:
        grade, grade_color = "EXCELLENT", "#22c55e"
    elif score >= 70:
        grade, grade_color = "GOOD", "#84cc16"
    elif score >= 50:
        grade, grade_color = "NEEDS WORK", "#f59e0b"
    else:
        grade, grade_color = "POOR", "#ef4444"

    # Build the full feature tree HTML
    features = actual.get("features", [])
    feature_rows = ""
    for i, f in enumerate(features, 1):
        feature_rows += f"""
            <tr>
                <td>{i}</td>
                <td><code>{f.get('name', 'Unknown')}</code></td>
                <td><span class="badge">{f.get('type', 'Unknown')}</span></td>
            </tr>"""
    if not features:
        feature_rows = '<tr><td colspan="3" class="muted">No features traversable</td></tr>'

    # Bounding box
    bbox = actual.get("bounding_box", {})
    bbox_html = ""
    if bbox.get("min_x") is not None:
        bbox_html = f"""
        <table class="data-table">
            <tr><th>Axis</th><th>Min (mm)</th><th>Max (mm)</th><th>Size (mm)</th></tr>
            <tr><td>X (Width)</td><td>{bbox.get('min_x', 0)}</td><td>{bbox.get('max_x', 0)}</td><td><strong>{bbox.get('max_x', 0) - bbox.get('min_x', 0):.2f}</strong></td></tr>
            <tr><td>Y (Height)</td><td>{bbox.get('min_y', 0)}</td><td>{bbox.get('max_y', 0)}</td><td><strong>{bbox.get('max_y', 0) - bbox.get('min_y', 0):.2f}</strong></td></tr>
            <tr><td>Z (Depth)</td><td>{bbox.get('min_z', 0)}</td><td>{bbox.get('max_z', 0)}</td><td><strong>{bbox.get('max_z', 0) - bbox.get('min_z', 0):.2f}</strong></td></tr>
        </table>"""

    # Scored checks rows
    scored_rows = ""
    for c in result["scored_checks"]:
        status = c["status"]
        if status == "PASS":
            status_html = '<span class="status-pass">✅ PASS</span>'
        elif status == "WARN":
            status_html = '<span class="status-warn">⚠️ WARN</span>'
        elif status == "FAIL":
            status_html = '<span class="status-fail">❌ FAIL</span>'
        else:
            status_html = '<span class="status-skip">⏭️ SKIP</span>'
        scored_rows += f"""
            <tr>
                <td>{c['check']}</td>
                <td>{c.get('expected', 'N/A')}</td>
                <td>{c.get('actual', 'N/A')}</td>
                <td>{status_html}</td>
            </tr>"""

    # Score breakdown
    breakdown = result.get("scores_breakdown", {})
    labels = {"dim": "Dimensions", "feat": "Features", "struct": "Structure", "vol": "Volume", "sa": "Surface Area"}
    breakdown_rows = ""
    for key, label in labels.items():
        val = breakdown.get(key, "N/A")
        if val == "N/A":
            breakdown_rows += f'<tr><td>{label}</td><td class="muted">— not scored</td><td></td></tr>'
        else:
            bar_width = val
            bar_color = "#22c55e" if val >= 80 else ("#f59e0b" if val >= 50 else "#ef4444")
            breakdown_rows += f"""
                <tr>
                    <td>{label}</td>
                    <td>{val}/100</td>
                    <td><div class="bar-bg"><div class="bar-fill" style="width:{bar_width}%; background:{bar_color}"></div></div></td>
                </tr>"""

    # Info items
    info_html = ""
    specs_items = [i for i in result["info_items"] if i["check"] == "Spec"]
    design_items = [i for i in result["info_items"] if i["check"] == "Design"]
    measure_items = [i for i in result["info_items"] if i["check"] not in ("Spec", "Design")]

    if measure_items:
        info_html += '<h3>📊 Measured Properties</h3><table class="data-table"><tr><th>Property</th><th>Value</th><th>Note</th></tr>'
        for item in measure_items:
            info_html += f'<tr><td>{item["check"]}</td><td><strong>{item["actual"]}</strong></td><td class="muted">{item.get("note", "")}</td></tr>'
        info_html += '</table>'

    if specs_items:
        info_html += '<h3>📋 Specific Measurements (from prompt)</h3><ul class="check-list">'
        for item in specs_items:
            info_html += f'<li class="check-item spec-item">{item["actual"]}</li>'
        info_html += '</ul>'

    if design_items:
        info_html += '<h3>👁️ Design Verification Checklist</h3><p class="muted">Verify these visually in SolidWorks:</p><ul class="check-list">'
        for item in design_items:
            info_html += f'<li class="check-item design-item">{item["actual"]}</li>'
        info_html += '</ul>'

    # Deviations
    dev_html = ""
    if result["deviations"]:
        dev_html = '<h3>❌ Deviations Found</h3><ul class="deviation-list">'
        for dev in result["deviations"]:
            dev_html += f'<li>{dev}</li>'
        dev_html += '</ul>'

    # Volume / Surface Area detail
    vol_sa_html = ""
    actual_vol = actual.get("volume_mm3")
    actual_sa = actual.get("surface_area_mm2")
    shape = spec.get("expected_shape", "custom")

    if actual_vol or actual_sa:
        vol_sa_html = '<table class="data-table"><tr><th>Metric</th><th>Measured Value</th><th>Scoring</th></tr>'
        if actual_vol:
            vol_scored = "Scored (Python formula)" if shape in ("box", "cylinder", "sphere", "cone") else "Not scored (complex shape)"
            vol_sa_html += f'<tr><td>Volume</td><td><strong>{actual_vol:,.2f} mm³</strong></td><td class="muted">{vol_scored}</td></tr>'
        if actual_sa:
            sa_scored = "Scored (Python formula)" if shape in ("box", "cylinder", "sphere", "cone") else "Not scored (complex shape)"
            vol_sa_html += f'<tr><td>Surface Area</td><td><strong>{actual_sa:,.2f} mm²</strong></td><td class="muted">{sa_scored}</td></tr>'
        vol_sa_html += '</table>'

    # Feature types summary
    feature_types = actual.get("feature_types", {})
    ft_html = ""
    if feature_types:
        ft_html = '<div class="feature-tags">'
        for ft, count in feature_types.items():
            ft_html += f'<span class="badge">{ft} ×{count}</span> '
        ft_html += '</div>'

    # Expected spec summary
    expected_dims = spec.get("expected_dimensions", {})
    exp_html = f"""
        <table class="data-table">
            <tr><th>Field</th><th>Value</th></tr>
            <tr><td>Description</td><td>{spec.get('description', 'N/A')}</td></tr>
            <tr><td>Expected Shape</td><td><span class="badge">{spec.get('expected_shape', 'N/A')}</span></td></tr>
            <tr><td>Expected Dimensions</td><td>{expected_dims.get('width', '?')} W × {expected_dims.get('height', '?')} H × {expected_dims.get('depth', '?')} D mm</td></tr>
            <tr><td>Expected Features</td><td>{', '.join(spec.get('expected_features', []))}</td></tr>
            <tr><td>Expected Bodies</td><td>{spec.get('expected_body_count', '?')}</td></tr>
            <tr><td>Tolerance</td><td>{spec.get('tolerance_percent', 10)}%</td></tr>
        </table>"""

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Validation Report — {spec.get('description', 'Model')}</title>
<style>
    * {{ margin: 0; padding: 0; box-sizing: border-box; }}
    body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #0f172a; color: #e2e8f0; padding: 24px; line-height: 1.6; }}
    .container {{ max-width: 900px; margin: 0 auto; }}

    /* Header */
    .header {{ background: linear-gradient(135deg, #1e293b, #334155); border-radius: 12px; padding: 32px; margin-bottom: 24px; border: 1px solid #475569; }}
    .header h1 {{ font-size: 24px; margin-bottom: 8px; }}
    .header .prompt {{ color: #94a3b8; font-style: italic; margin-bottom: 16px; word-break: break-word; }}
    .header .meta {{ color: #64748b; font-size: 13px; }}

    /* Score card */
    .score-card {{ display: flex; align-items: center; gap: 24px; background: #1e293b; border-radius: 12px; padding: 24px; margin-bottom: 24px; border: 1px solid #475569; }}
    .score-circle {{ width: 100px; height: 100px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 32px; font-weight: 800; border: 4px solid {grade_color}; color: {grade_color}; flex-shrink: 0; }}
    .score-details {{ flex: 1; }}
    .score-grade {{ font-size: 20px; font-weight: 700; color: {grade_color}; }}
    .score-note {{ color: #94a3b8; margin-top: 4px; font-size: 14px; }}

    /* Section */
    .section {{ background: #1e293b; border-radius: 12px; padding: 24px; margin-bottom: 20px; border: 1px solid #475569; }}
    .section h2 {{ font-size: 18px; margin-bottom: 16px; color: #f1f5f9; border-bottom: 1px solid #334155; padding-bottom: 8px; }}
    .section h3 {{ font-size: 15px; margin: 16px 0 10px; color: #cbd5e1; }}

    /* Tables */
    .data-table {{ width: 100%; border-collapse: collapse; margin: 8px 0 16px; }}
    .data-table th {{ background: #334155; padding: 10px 14px; text-align: left; font-size: 13px; color: #94a3b8; text-transform: uppercase; letter-spacing: 0.5px; }}
    .data-table td {{ padding: 10px 14px; border-bottom: 1px solid #334155; font-size: 14px; }}
    .data-table tr:hover {{ background: #293548; }}
    .data-table code {{ background: #334155; padding: 2px 6px; border-radius: 4px; font-size: 13px; }}

    /* Badges */
    .badge {{ background: #334155; color: #93c5fd; padding: 3px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; display: inline-block; margin: 2px; }}

    /* Status */
    .status-pass {{ color: #22c55e; font-weight: 700; }}
    .status-warn {{ color: #f59e0b; font-weight: 700; }}
    .status-fail {{ color: #ef4444; font-weight: 700; }}
    .status-skip {{ color: #64748b; }}

    /* Progress bars */
    .bar-bg {{ background: #334155; border-radius: 4px; height: 8px; width: 120px; display: inline-block; vertical-align: middle; }}
    .bar-fill {{ height: 100%; border-radius: 4px; transition: width 0.3s; }}

    /* Lists */
    .check-list {{ list-style: none; padding: 0; }}
    .check-item {{ padding: 8px 12px; margin: 4px 0; border-radius: 6px; font-size: 14px; }}
    .spec-item {{ background: #1a2744; border-left: 3px solid #3b82f6; }}
    .design-item {{ background: #1a2744; border-left: 3px solid #a855f7; }}
    .deviation-list {{ list-style: none; padding: 0; }}
    .deviation-list li {{ padding: 8px 12px; margin: 4px 0; background: #2a1a1a; border-left: 3px solid #ef4444; border-radius: 6px; font-size: 14px; }}

    .muted {{ color: #64748b; font-size: 13px; }}
    .feature-tags {{ margin: 8px 0 16px; }}

    /* Two column layout */
    .two-col {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }}
    @media (max-width: 700px) {{ .two-col {{ grid-template-columns: 1fr; }} }}
</style>
</head>
<body>
<div class="container">

    <!-- HEADER -->
    <div class="header">
        <h1>🔧 Design Validation Report</h1>
        <div class="prompt">"{prompt}"</div>
        <div class="meta">Generated: {now} &nbsp;|&nbsp; Shape: {spec.get('expected_shape', 'custom')} &nbsp;|&nbsp; Tolerance: {spec.get('tolerance_percent', 10)}%</div>
    </div>

    <!-- SCORE -->
    <div class="score-card">
        <div class="score-circle">{score}</div>
        <div class="score-details">
            <div class="score-grade">{grade}</div>
            <div class="score-note">Score based only on reliably verifiable checks. Volume/SA scored only for simple shapes with exact Python formulas.</div>
        </div>
    </div>

    <!-- SCORED CHECKS -->
    <div class="section">
        <h2>✅ Scored Checks</h2>
        <table class="data-table">
            <tr><th>Check</th><th>Expected</th><th>Actual</th><th>Status</th></tr>
            {scored_rows}
        </table>

        <h3>Score Breakdown</h3>
        <table class="data-table">
            <tr><th style="width:150px">Category</th><th style="width:80px">Score</th><th>Bar</th></tr>
            {breakdown_rows}
        </table>
    </div>

    <!-- DETAILED MEASUREMENTS -->
    <div class="section">
        <h2>📐 Model Measurements</h2>

        <div class="two-col">
            <div>
                <h3>Bounding Box</h3>
                {bbox_html if bbox_html else '<p class="muted">Not available</p>'}
            </div>
            <div>
                <h3>Volume & Surface Area</h3>
                {vol_sa_html if vol_sa_html else '<p class="muted">Not available</p>'}
            </div>
        </div>

        <h3>Body Summary</h3>
        <table class="data-table">
            <tr><th>Bodies</th><th>Faces</th><th>Edges</th></tr>
            <tr><td><strong>{actual.get('body_count', 0)}</strong></td><td><strong>{actual.get('face_count', 0)}</strong></td><td><strong>{actual.get('edge_count', 0)}</strong></td></tr>
        </table>
    </div>

    <!-- FEATURE TREE -->
    <div class="section">
        <h2>🌳 Feature Tree ({actual.get('feature_count', 0)} features)</h2>
        {ft_html}
        <table class="data-table">
            <tr><th>#</th><th>Feature Name</th><th>Type</th></tr>
            {feature_rows}
        </table>
    </div>

    <!-- INFO & VERIFICATION -->
    <div class="section">
        <h2>🔍 Verification Details</h2>
        {info_html if info_html else '<p class="muted">No additional verification items</p>'}
        {dev_html}
    </div>

    <!-- EXPECTED SPEC -->
    <div class="section">
        <h2>📋 Expected Specification (from LLM)</h2>
        {exp_html}
        {f'<p class="muted" style="margin-top:12px"><strong>Notes:</strong> {spec.get("notes", "")}</p>' if spec.get("notes") else ''}
    </div>

</div>
</body>
</html>"""

    report_path = "validation_report.html"
    try:
        with open(report_path, "w", encoding="utf-8") as f:
            f.write(html)
        print(f"📄 Detailed report saved to '{report_path}'")
        print(f"📁 JSON data saved to 'validation_report.json'")

        # Auto-open in browser
        import webbrowser
        import os
        abs_path = os.path.abspath(report_path)
        webbrowser.open(f"file:///{abs_path}")
        print(f"🌐 Opened report in browser\n")
    except Exception as e:
        print(f"⚠️  Could not save/open HTML report: {e}\n")