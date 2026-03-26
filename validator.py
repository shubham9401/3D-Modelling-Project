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
        "Extrude": ["ICE", "Extrusion", "Boss-Extrude", "Extrude", "Extrusion"],
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

    ✅ FIXED: Now uses expected_volume_mm3 from spec (calculated by validation_prompt.py)
    The validation prompt already accounts for shell, cuts, holes, etc.
    We just compare actual vs expected directly.

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

    # ── 3. Volume (25%) - ✅ FIXED: USE SPEC VALUE ──
    volume_score = None  # None means SKIP (will redistribute weight)
    actual_vol = actual_properties.get("volume_mm3")
    
    # ✅ USE EXPECTED VOLUME FROM SPEC (calculated by validation_prompt.py with shell awareness)
    expected_vol = spec.get("expected_volume_mm3")
    
    if actual_vol and expected_vol and expected_vol > 0:
        # Spec already accounts for shell, cuts, holes - just compare directly
        vol_ratio = actual_vol / expected_vol
        error_pct = abs(1 - vol_ratio) * 100
        
        if error_pct <= 5:
            # Within 5% - excellent
            volume_score = 100
            checks.append({"check": "Volume", "expected": round(expected_vol, 1), "actual": actual_vol, "status": "PASS"})
        elif error_pct <= 10:
            # Within 10% - good
            volume_score = 90
            checks.append({"check": "Volume", "expected": round(expected_vol, 1), "actual": actual_vol, "status": "PASS"})
        elif error_pct <= 20:
            # Within 20% - acceptable
            volume_score = 70
            checks.append({"check": "Volume", "expected": round(expected_vol, 1), "actual": actual_vol, "status": "WARN"})
            deviations.append(f"Volume deviation: expected {round(expected_vol,1)}mm³, got {actual_vol}mm³ (±{error_pct:.1f}%)")
        else:
            # Outside 20% - failure
            volume_score = 0
            checks.append({"check": "Volume", "expected": round(expected_vol, 1), "actual": actual_vol, "status": "FAIL"})
            deviations.append(f"Volume mismatch: expected {round(expected_vol,1)}mm³, got {actual_vol}mm³ (±{error_pct:.1f}%)")
    elif actual_vol and not expected_vol:
        # Volume measured but no expected value - just show it exists
        volume_score = None  # Don't score, redistribute weight
        checks.append({"check": "Volume", "expected": "N/A (not calculable)", "actual": actual_vol, "status": "SKIP"})
    else:
        # No volume measured
        checks.append({"check": "Volume", "expected": "N/A", "actual": str(actual_vol) if actual_vol else "N/A", "status": "SKIP"})

    # ── 4. Surface Area (15%) - ✅ FIXED: USE SPEC VALUE ──
    surface_score = None  # None means SKIP
    actual_sa = actual_properties.get("surface_area_mm2")
    
    # ✅ USE EXPECTED SURFACE AREA FROM SPEC
    expected_sa = spec.get("expected_surface_area_mm2")
    
    if actual_sa and expected_sa and expected_sa > 0:
        sa_ratio = actual_sa / expected_sa
        error_pct = abs(1 - sa_ratio) * 100
        
        if error_pct <= 10:
            # Within 10% - good (SA is harder to predict exactly)
            surface_score = 100
            checks.append({"check": "Surface Area", "expected": round(expected_sa, 1), "actual": actual_sa, "status": "PASS"})
        elif error_pct <= 20:
            # Within 20% - acceptable
            surface_score = 80
            checks.append({"check": "Surface Area", "expected": round(expected_sa, 1), "actual": actual_sa, "status": "PASS"})
        elif error_pct <= 30:
            # Within 30% - warning
            surface_score = 50
            checks.append({"check": "Surface Area", "expected": round(expected_sa, 1), "actual": actual_sa, "status": "WARN"})
            deviations.append(f"Surface area deviation: expected {round(expected_sa,1)}mm², got {actual_sa}mm²")
        else:
            # Outside 30% - failure
            surface_score = 0
            checks.append({"check": "Surface Area", "expected": round(expected_sa, 1), "actual": actual_sa, "status": "FAIL"})
            deviations.append(f"Surface area mismatch: expected {round(expected_sa,1)}mm², got {actual_sa}mm²")
    elif actual_sa and not expected_sa:
        # SA measured but no expected value
        surface_score = None
        checks.append({"check": "Surface Area", "expected": "N/A (not calculable)", "actual": actual_sa, "status": "SKIP"})
    else:
        # No SA measured
        checks.append({"check": "Surface Area", "expected": "N/A", "actual": str(actual_sa) if actual_sa else "N/A", "status": "SKIP"})

    # ── 5. Face/Body Count (15%) ──
    expected_bodies = spec.get("expected_body_count", 1)
    actual_bodies = actual_properties.get("body_count", 0)
    actual_faces = actual_properties.get("face_count", 0)
    expected_faces = spec.get("expected_face_count")
    
    body_match = actual_bodies == expected_bodies
    
    # Face count validation
    face_score = 100
    if expected_faces and isinstance(expected_faces, int):
        # Exact expected face count provided
        face_diff = abs(actual_faces - expected_faces)
        if face_diff == 0:
            face_score = 100
        elif face_diff <= 2:
            face_score = 90
        elif face_diff <= 5:
            face_score = 70
        else:
            face_score = 50
    # else: no expected face count, give full score

    body_score = 100 if body_match else 50
    struct_score = (face_score + body_score) / 2

    expected_faces_str = str(expected_faces) if expected_faces else "?"
    checks.append({
        "check": "Face/Body Count",
        "expected": f"{expected_bodies}B / ~{expected_faces_str}F",
        "actual": f"{actual_bodies}B / {actual_faces}F",
        "status": "PASS" if (body_match and face_score >= 70) else "WARN"
    })

    # ── FINAL SCORE CALCULATION ──
    # Weights: dim=25%, feat=20%, vol=25%, sa=15%, struct=15%
    weights = {
        "dim": 0.25,
        "feat": 0.20,
        "vol": 0.25,
        "sa": 0.15,
        "struct": 0.15
    }

    scores = {
        "dim": dim_avg,
        "feat": feat_score,
        "vol": volume_score,
        "sa": surface_score,
        "struct": struct_score
    }

    # Redistribute weight from skipped categories
    active_categories = {k: v for k, v in scores.items() if v is not None}
    if not active_categories:
        final_score = 0
    else:
        total_weight = sum(weights[k] for k in active_categories.keys())
        final_score = sum(scores[k] * (weights[k] / total_weight) for k in active_categories.keys())

    return {
        "score": round(final_score, 0),
        "checks": checks,
        "deviations": deviations,
        "scores_breakdown": {k: round(v) if v is not None else None for k, v in scores.items()}
    }


# ============================================================
# MAIN VALIDATION FUNCTION
# ============================================================

def run_validation(user_prompt):
    """
    Main entry point: validates the current SolidWorks model against the prompt.
    
    Returns validation result dict or None on error.
    """
    try:
        # Step 1: Get actual model properties
        print("\n🔍 Inspecting current model...")
        from tools.model_inspector import get_model_properties
        actual_props = get_model_properties()
        
        if actual_props is None:
            print("❌ Could not inspect model")
            return None
        
        print(f"   ✅ Model inspected: {actual_props['dimensions']['width']}W x {actual_props['dimensions']['height']}H x {actual_props['dimensions']['depth']}D mm, {actual_props['feature_count']} features, {actual_props['body_count']} bodies")
        
        # Step 2: Extract expected specs from prompt
        spec = create_spec_from_prompt(user_prompt)
        if spec is None:
            print("❌ Could not extract expected specs from prompt")
            return None
        
        print(f"   📐 Expected: {spec.get('description', 'N/A')}")
        
        # Step 3: Compare actual vs expected
        print("\n📊 Comparing actual model vs expected specs...")
        result = validate_model(spec, actual_props)
        
        # Step 4: Print report
        print_validation_report(user_prompt, result, spec, actual_props)
        
        # Step 5: Save detailed report
        save_validation_report(user_prompt, result, spec, actual_props)
        
        return result
        
    except Exception as e:
        print(f"\n❌ Validation error: {e}")
        import traceback
        traceback.print_exc()
        return None


def print_validation_report(prompt, result, spec, actual_props):
    """Print a formatted validation report to console."""
    score = result["score"]
    
    # Status emoji
    if score >= 80:
        status_emoji = "✅ EXCELLENT"
    elif score >= 60:
        status_emoji = "⚠️  ACCEPTABLE"
    else:
        status_emoji = "❌ NEEDS WORK"
    
    print("\n" + "═" * 55)
    print("  DESIGN VALIDATION REPORT")
    print("═" * 55)
    print(f'  Prompt: "{prompt}"\n')
    print(f"  OVERALL SCORE: {score}/100  ({status_emoji})\n")
    
    # Print checks
    print(f"  {'Check':<20} {'Expected':<12} {'Actual':<12} {'Status':<8}")
    print(f"  {'─'*20} {'─'*12} {'─'*12} {'─'*8}")
    
    for check in result["checks"]:
        name = check["check"]
        expected = str(check["expected"])[:10]
        actual = str(check["actual"])[:10]
        status = check["status"]
        
        # Status formatting
        if status == "PASS":
            status_str = "✅ PASS"
        elif status == "WARN":
            status_str = "⚠️  WARN"
        elif status == "FAIL":
            status_str = "❌ FAIL"
        else:
            status_str = "SKIP"
        
        print(f"  {name:<20} {expected:<12} {actual:<12} {status_str:<8}")
    
    print("\n" + "═" * 55)
    
    # Print breakdown
    breakdown = result.get("scores_breakdown", {})
    if any(v is not None for v in breakdown.values()):
        print("  Score Breakdown:")
        for category, score_val in breakdown.items():
            if score_val is not None:
                print(f"    {category}: {score_val}/100")
        print()
    
    # Print deviations
    if result["deviations"]:
        print("  Deviations found:")
        for dev in result["deviations"]:
            print(f"    • {dev}")
        print()


def save_validation_report(prompt, result, spec, actual_props):
    """Save detailed validation report to JSON file."""
    report = {
        "prompt": prompt,
        "score": result["score"],
        "status": "PASS" if result["score"] >= 70 else "FAIL",
        "expected_spec": spec,
        "actual_properties": {
            "dimensions": actual_props.get("dimensions"),
            "volume_mm3": actual_props.get("volume_mm3"),
            "surface_area_mm2": actual_props.get("surface_area_mm2"),
            "body_count": actual_props.get("body_count"),
            "face_count": actual_props.get("face_count"),
            "feature_count": actual_props.get("feature_count"),
        },
        "checks": result["checks"],
        "deviations": result["deviations"],
        "scores_breakdown": result.get("scores_breakdown", {}),
    }
    
    filename = "validation_report.json"
    with open(filename, "w") as f:
        json.dump(report, f, indent=2)
    
    print(f"📁 Full report saved to '{filename}'\n")