"""
Validator: Compares a generated SolidWorks model against expected specs.

Produces:
    - A similarity score (0-100%)
    - An error analysis report with per-check PASS/WARN/FAIL status
"""

import json
import os


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
        # Find the JSON object
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
# COMPARISON LOGIC
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


def _check_features(expected_features, actual_feature_types):
    """Check if expected features are present. Returns (score, checks)."""
    if not expected_features:
        return 100, []

    checks = []
    found = 0

    # Map common feature type names to SolidWorks internal names
    TYPE_ALIASES = {
        "Extrude": ["ICE", "Extrusion", "Boss-Extrude", "Extrude"],
        "Cut": ["ICE", "Cut-Extrude", "Cut"],
        "Shell": ["Shell"],
        "Fillet": ["Fillet"],
        "Chamfer": ["Chamfer"],
        "Revolve": ["Revolution", "Revolve"],
        "Loft": ["Loft"],
        "Sweep": ["Sweep"],
        "CircularPattern": ["CirPattern"],
        "LinearPattern": ["LPattern"],
        "Hole": ["HoleWzd", "Hole"],
        "Sketch": ["ProfileFeature"],
    }

    actual_types_lower = {k.lower(): v for k, v in actual_feature_types.items()}

    for expected_feat in expected_features:
        # Check direct match or alias match
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


def validate_model(spec, actual_properties):
    """
    Compare expected spec against actual model properties.

    Weights:
        Volume:               25%  (single most reliable geometric truth)
        Bounding Box Dims:    25%  (height, width, depth directly measurable)
        Feature Completeness: 20%  (did the right operations get applied?)
        Surface Area:         15%  (catches wall thickness errors)
        Face/Body Count:      15%  (catches structural mistakes)

    Returns:
        dict with 'score' (0-100), 'checks' (list), 'deviations' (list)
    """
    tolerance = spec.get("tolerance_percent", 10)
    checks = []
    deviations = []

    # ── 1. Bounding Box Dimensions (25%) ──
    dim_scores = []
    expected_dims = spec.get("expected_dimensions", {})
    actual_dims = actual_properties.get("dimensions", {})

    for dim_name, dim_key in [("Width", "width"), ("Height", "height"), ("Depth", "depth")]:
        expected_val = expected_dims.get(dim_key)
        actual_val = actual_dims.get(dim_key)
        score, status, detail = _check_dimension(dim_name, expected_val, actual_val, tolerance)

        if status != "SKIP":
            dim_scores.append(score)
            checks.append({"check": dim_name, "expected": expected_val, "actual": actual_val, "status": status})
            if status == "FAIL":
                deviations.append(detail)

    dim_avg = sum(dim_scores) / len(dim_scores) if dim_scores else 100

    # ── 2. Feature Completeness (20%) ──
    expected_features = spec.get("expected_features", [])
    feature_tree_available = actual_properties.get("feature_tree_available", True)
    
    if not feature_tree_available and actual_properties.get("feature_count", 0) > 0:
        feat_score = 100
        checks.append({
            "check": "Features",
            "expected": ", ".join(expected_features) if expected_features else "N/A",
            "actual": f"{actual_properties['feature_count']} features",
            "status": "PASS",
        })
    else:
        feat_score, feat_checks = _check_features(expected_features, actual_properties.get("feature_types", {}))

        for status, detail in feat_checks:
            feat_name = detail.split("'")[1] if "'" in detail else detail
            checks.append({
                "check": f"Feature: {feat_name}",
                "expected": "Yes",
                "actual": "Yes" if status == "PASS" else "No",
                "status": status,
            })
            if status == "FAIL":
                deviations.append(detail)

    # ── 3. Volume (25%) ──
    volume_score = 100  # Default pass if we can't measure
    if actual_properties.get("volume_mm3") and expected_dims.get("width") and expected_dims.get("height") and expected_dims.get("depth"):
        w = expected_dims["width"]
        h = expected_dims["height"]
        d = expected_dims["depth"]
        shape = spec.get("expected_shape", "box")

        if shape == "box":
            expected_vol = w * h * d
        elif shape == "cylinder":
            r = w / 2
            expected_vol = 3.14159 * r * r * h
        elif shape == "sphere":
            r = w / 2
            expected_vol = (4/3) * 3.14159 * r * r * r
        else:
            expected_vol = None

        if expected_vol and expected_vol > 0:
            actual_vol = actual_properties["volume_mm3"]
            vol_ratio = actual_vol / expected_vol
            if 0.8 <= vol_ratio <= 1.2:
                volume_score = 100
                checks.append({"check": "Volume", "expected": round(expected_vol, 1), "actual": actual_vol, "status": "PASS"})
            elif 0.5 <= vol_ratio <= 1.5:
                volume_score = 50
                checks.append({"check": "Volume", "expected": round(expected_vol, 1), "actual": actual_vol, "status": "WARN"})
                deviations.append(f"Volume deviation: expected ~{round(expected_vol,1)}mm³, got {actual_vol}mm³")
            else:
                volume_score = 0
                checks.append({"check": "Volume", "expected": round(expected_vol, 1), "actual": actual_vol, "status": "FAIL"})
                deviations.append(f"Volume mismatch: expected ~{round(expected_vol,1)}mm³, got {actual_vol}mm³")
    else:
        # Volume not measurable — skip with neutral score
        checks.append({"check": "Volume", "expected": "N/A", "actual": "N/A", "status": "SKIP"})

    # ── 4. Surface Area (15%) ──
    surface_score = 100
    if actual_properties.get("surface_area_mm2") and expected_dims.get("width") and expected_dims.get("height") and expected_dims.get("depth"):
        w = expected_dims["width"]
        h = expected_dims["height"]
        d = expected_dims["depth"]
        shape = spec.get("expected_shape", "box")

        if shape == "box":
            expected_sa = 2 * (w*h + w*d + h*d)
        elif shape == "cylinder":
            r = w / 2
            expected_sa = 2 * 3.14159 * r * (r + h)
        elif shape == "sphere":
            r = w / 2
            expected_sa = 4 * 3.14159 * r * r
        else:
            expected_sa = None

        if expected_sa and expected_sa > 0:
            actual_sa = actual_properties["surface_area_mm2"]
            sa_ratio = actual_sa / expected_sa
            if 0.8 <= sa_ratio <= 1.2:
                surface_score = 100
                checks.append({"check": "Surface Area", "expected": round(expected_sa, 1), "actual": actual_sa, "status": "PASS"})
            elif 0.5 <= sa_ratio <= 1.5:
                surface_score = 50
                checks.append({"check": "Surface Area", "expected": round(expected_sa, 1), "actual": actual_sa, "status": "WARN"})
                deviations.append(f"Surface area deviation: expected ~{round(expected_sa,1)}mm², got {actual_sa}mm²")
            else:
                surface_score = 0
                checks.append({"check": "Surface Area", "expected": round(expected_sa, 1), "actual": actual_sa, "status": "FAIL"})
                deviations.append(f"Surface area mismatch: expected ~{round(expected_sa,1)}mm², got {actual_sa}mm²")
    else:
        checks.append({"check": "Surface Area", "expected": "N/A", "actual": "N/A", "status": "SKIP"})

    # ── 5. Face/Body Count (15%) ──
    expected_bodies = spec.get("expected_body_count", 1)
    actual_bodies = actual_properties.get("body_count", 0)
    actual_faces = actual_properties.get("face_count", 0)
    
    body_match = actual_bodies == expected_bodies
    
    # Estimate expected face count based on shape
    shape = spec.get("expected_shape", "box")
    expected_faces_estimate = None
    if shape == "box":
        expected_faces_estimate = 6
    elif shape == "cylinder":
        expected_faces_estimate = 3  # top, bottom, curved
    elif shape == "sphere":
        expected_faces_estimate = 1  # single curved surface
    
    face_ok = True
    if expected_faces_estimate and actual_faces > 0:
        face_ok = actual_faces >= expected_faces_estimate
    
    if body_match and face_ok:
        fb_score = 100
        status_str = "PASS"
    elif body_match or face_ok:
        fb_score = 50
        status_str = "WARN"
    else:
        fb_score = 0
        status_str = "FAIL"
    
    checks.append({
        "check": "Face/Body Count",
        "expected": f"{expected_bodies}B / ~{expected_faces_estimate or '?'}F",
        "actual": f"{actual_bodies}B / {actual_faces}F",
        "status": status_str,
    })
    if not body_match:
        deviations.append(f"Body count: expected {expected_bodies}, got {actual_bodies}")
    if expected_faces_estimate and actual_faces > 0 and not face_ok:
        deviations.append(f"Face count: expected ≥{expected_faces_estimate}, got {actual_faces}")

    # ── Weighted Final Score ──
    overall_score = round(
        volume_score   * 0.25 +
        dim_avg        * 0.25 +
        feat_score     * 0.20 +
        surface_score  * 0.15 +
        fb_score       * 0.15
    )

    return {
        "score": overall_score,
        "checks": checks,
        "deviations": deviations,
    }


# ============================================================
# REPORT FORMATTING
# ============================================================

def format_report(user_prompt, result):
    """
    Formats the validation result as a human-readable report.
    """
    lines = []
    lines.append("")
    lines.append("═" * 55)
    lines.append("  DESIGN VALIDATION REPORT")
    lines.append("═" * 55)
    lines.append(f'  Prompt: "{user_prompt}"')
    lines.append("")

    score = result["score"]
    if score >= 80:
        grade = "✅ EXCELLENT"
    elif score >= 60:
        grade = "⚠️  ACCEPTABLE"
    else:
        grade = "❌ POOR"

    lines.append(f"  OVERALL SCORE: {score}/100  ({grade})")
    lines.append("")

    # Table header
    lines.append(f"  {'Check':<20} {'Expected':<12} {'Actual':<12} {'Status':<8}")
    lines.append(f"  {'─'*20} {'─'*12} {'─'*12} {'─'*8}")

    for check in result["checks"]:
        status_icon = {"PASS": "✅ PASS", "WARN": "⚠️  WARN", "FAIL": "❌ FAIL"}.get(check["status"], check["status"])
        exp_str = str(check.get("expected", ""))[:11]
        act_str = str(check.get("actual", ""))[:11]
        lines.append(f"  {check['check']:<20} {exp_str:<12} {act_str:<12} {status_icon}")

    if result["deviations"]:
        lines.append("")
        lines.append("  Deviations:")
        for dev in result["deviations"]:
            lines.append(f"    ❌ {dev}")

    lines.append("")
    lines.append("═" * 55)

    return "\n".join(lines)


# ============================================================
# MAIN VALIDATION PIPELINE
# ============================================================

def run_validation(user_prompt):
    """
    Full validation pipeline:
    1. Inspect the current SolidWorks model
    2. Extract expected specs from the user prompt via LLM
    3. Compare actual vs expected
    4. Print the validation report

    Returns the result dict with score and checks.
    """
    from tools.model_inspector import get_model_properties

    print("\n🔍 Inspecting current model...")
    try:
        actual = get_model_properties()
        print(f"   ✅ Model inspected: {actual['dimensions']['width']}W x "
              f"{actual['dimensions']['height']}H x {actual['dimensions']['depth']}D mm, "
              f"{actual['feature_count']} features, {actual['body_count']} bodies")
    except Exception as e:
        print(f"   ❌ Failed to inspect model: {e}")
        return None

    # Extract specs from prompt
    spec = create_spec_from_prompt(user_prompt)
    if spec is None:
        print("   ❌ Could not extract specs from prompt")
        return None

    print(f"   📐 Expected: {spec.get('description', 'N/A')}")

    # Compare
    print("\n📊 Comparing actual model vs expected specs...")
    result = validate_model(spec, actual)

    # Print report
    report = format_report(user_prompt, result)
    print(report)

    # Save report to file
    report_data = {
        "prompt": user_prompt,
        "spec": spec,
        "actual": actual,
        "result": result,
    }
    with open("validation_report.json", "w") as f:
        json.dump(report_data, f, indent=4)
    print(f"📁 Full report saved to 'validation_report.json'")

    return result
