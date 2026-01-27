"""
Feature module for SolidWorks CAD engine.
"""

import math

try:
    from .solidworks_app import get_active_model, get_nothing
except ImportError:
    from solidworks_app import get_active_model, get_nothing

def _model():
    model = get_active_model()
    if model is None:
        raise Exception("No active SolidWorks document")
    return model

def _fm():
    return _model().FeatureManager

def _require_part():
    model = _model()
    val = model.GetType
    if callable(val):
        val = val()
    if val != 1:
        raise Exception("Active document is not a PART")

# ============================================================
# EXTRUDE
# ============================================================

def extrude(depth):
    """Boss-Extrude. Depth in mm. Negative = downward/backward."""
    _require_part()
    depth_m = depth / 1000.0
    
    result = _fm().FeatureExtrusion2(
        True, False, False,
        0, 0,
        depth_m, 0,
        False, False, False, False,
        0, 0,
        False, False, False, False,
        True, True, True,
        0, 0, False
    )
    
    if result is None:
        raise Exception("Extrusion failed")
    
    return f"Extruded {depth}mm"

def extrude_midplane(depth):
    """Mid-plane extrude."""
    _require_part()
    half_depth = depth / 2000.0
    
    _fm().FeatureExtrusion2(
        True, False, False,
        6, 0,
        half_depth, half_depth,
        False, False, False, False,
        0, 0,
        False, False, False, False,
        True, True, True,
        0, 0, False
    )
    
    return f"Mid-plane extruded {depth}mm"

# ============================================================
# CUT
# ============================================================

def cut_extrude(depth):
    """Cut-Extrude. Depth in mm."""
    _require_part()
    depth_m = depth / 1000.0
    
    _fm().FeatureCut4(
        True, False, False, 
        0, 0,
        depth_m, 0,
        False, False, False, False, 0, 0,
        False, False, False, False, False, 
        True, True, False, False, False,
        0, 0, False, False
    )
    
    return f"Cut extruded {depth}mm"

def cut_through_all():
    """Through-all cut."""
    _require_part()
    
    _fm().FeatureCut4(
        True, False, False, 
        1, 0,
        0, 0, 
        False, False, False, False, 0, 0,
        False, False, False, False, False, 
        True, True, False, False, False,
        0, 0, False, False
    )
    
    return "Cut through all"

# ============================================================
# REVOLVE
# ============================================================

def revolve(angle=360):
    """
    Revolve around vertical axis. Requires centerline in sketch.
    Angle in degrees.
    """
    _require_part()
    
    result = _fm().FeatureRevolve2(
        True, True,
        False, False, False, False,
        0, 0,
        math.radians(angle), 0,
        False, False,
        0, 0, 0,
        False, False, False
    )
    
    if result is None:
        raise Exception("Revolve failed - ensure sketch has centerline")
    
    return f"Revolved {angle} degrees"

# ============================================================
# REFINEMENTS
# ============================================================

def fillet(radius):
    """Fillet selected edges. Radius in mm."""
    _require_part()
    r = radius / 1000.0
    _fm().InsertFeatureFillet(195, r, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    return f"Fillet: {radius}mm"

def chamfer(distance, angle=45):
    """Chamfer selected edges."""
    _require_part()
    d = distance / 1000.0
    _fm().InsertFeatureChamfer(4, d, math.radians(angle), 0, 0, 0, 0, 0)
    return f"Chamfer: {distance}mm @ {angle}°"

def linear_pattern(count, spacing):
    """Linear pattern."""
    _require_part()
    s = spacing / 1000.0
    _fm().FeatureLinearPattern3(count, 1, s, 0, False, False, "", "", False, False, True)
    return f"Linear pattern: {count} x {spacing}mm"

def circular_pattern(count, angle=360):
    """Circular pattern."""
    _require_part()
    _fm().FeatureCircularPattern3(count, math.radians(angle), False, "", False, True)
    return f"Circular pattern: {count} over {angle}°"

def mirror_feature():
    """Mirror feature."""
    _require_part()
    _fm().InsertMirrorFeature2(False, True, False, False)
    return "Mirrored"

def get_feature_count():
    """Returns feature count."""
    _require_part()
    return _model().GetFeatureCount(False)