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
    
    # LOGIC FIX: Handle negative depth by flipping direction
    depth_val = float(depth)
    is_negative = depth_val < 0
    depth_m = abs(depth_val) / 1000.0
    
    # Arg 3 is 'FlipDir'. If depth is negative, set True.
    flip_dir = is_negative
    
    # 0 = Blind
    _fm().FeatureExtrusion2(
        True, False, flip_dir, # Sd, FlipSide, Dir
        0, 0,
        depth_m, 0,
        False, False, False, False,
        0, 0,
        False, False, False, False,
        True, True, True,
        0, 0, False
    )
    return f"Extruded {depth}mm"

def extrude_midplane(depth):
    """Mid-plane extrude."""
    _require_part()
    half_depth = depth / 2000.0
    
    # 6 = MidPlane
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
    
    # LOGIC FIX: Handle negative depth for cuts too
    depth_val = float(depth)
    is_negative = depth_val < 0
    depth_m = abs(depth_val) / 1000.0
    
    flip_dir = is_negative

    # T1=0 (Blind)
    _fm().FeatureCut4(
        True, False, flip_dir, # Sd, FlipSide, FlipDir
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
    
    # T1=1 (Through All)
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
    Revolve boss feature.
    Requires axis/centerline to be in sketch.
    Angle in degrees.
    """
    _require_part()
    
    # FIX: Updated to 20 arguments to match your VBA/Version
    _fm().FeatureRevolve2(
        True,                   # SingleDir
        True,                   # IsSolid
        False,                  # IsThin
        False,                  # ReverseDir
        False,                  # ReverseDir2
        False,                  # MergeFaces
        0,                      # Dir1Type (0=Blind)
        0,                      # Dir2Type
        math.radians(angle),    # Dir1Angle
        0,                      # Dir2Angle
        False,                  # ReverseOffset
        False,                  # UseOffset2
        0.01,                   # Offset1
        0.01,                   # Offset2
        0,                      # ThinType
        0,                      # ThinThickness1
        0,                      # ThinThickness2
        True,                   # UseFeatScope
        True,                   # UseAutoSelect
        True                    # PropagateFeatureToParts
    )
    return f"Revolved {angle} degrees"

# ============================================================
# REFINEMENTS
# ============================================================

def fillet(radius):
    """Fillet selected edges."""
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