"""
Feature module for MCP-style SolidWorks CAD engine.

Responsibilities:
- Create solid features from sketches
- Apply feature-level operations (patterns, mirror)
- Enforce safety guards
- Expose semantic, LLM-safe feature operations

Does NOT:
- Manage documents
- Create sketches
- Decide geometry
"""

import math

try:
    from .solidworks_app import get_active_model, get_nothing
except ImportError:
    from solidworks_app import get_active_model, get_nothing

# ============================================================
# INTERNAL HELPERS
# ============================================================

def _model():
    model = get_active_model()
    if model is None:
        raise Exception("No active SolidWorks document")
    return model

def _fm():
    return _model().FeatureManager

def _require_part():
    # swDocPART = 1
    if _model().GetType != 1:
        raise Exception("Active document is not a PART")

# ============================================================
# EXTRUDE FEATURES
# ============================================================

def extrude(depth):
    """
    Boss-Extrude the last sketch.
    Depth in mm.
    """
    _require_part()
    
    depth_m = depth / 1000.0  # mm to meters
    
    # FeatureExtrusion2 parameters based on working code:
    # (Sd, Flip, Dir, T1, T2, D1, D2, Dchk1, Dchk2, Ddir1, Ddir2, 
    #  Dang1, Dang2, OffsetReverse1, OffsetReverse2, TranslateSurface1,
    #  TranslateSurface2, Merge, UseFeatScope, UseAutoSelect, T0, StartOffset, FlipStartOffset)
    _fm().FeatureExtrusion2(
        True,           # Sd (single direction)
        False,          # Flip
        False,          # Dir
        0,              # T1 (end condition: Blind = 0)
        0,              # T2
        depth_m,        # D1 (depth)
        0,              # D2
        False,          # Dchk1
        False,          # Dchk2
        False,          # Ddir1
        False,          # Ddir2
        0,              # Dang1
        0,              # Dang2
        False,          # OffsetReverse1
        False,          # OffsetReverse2
        False,          # TranslateSurface1
        False,          # TranslateSurface2
        True,           # Merge
        True,           # UseFeatScope
        True,           # UseAutoSelect
        0,              # T0
        0,              # StartOffset
        False           # FlipStartOffset
    )

    return f"Extruded {depth}mm"


def extrude_midplane(depth):
    """
    Mid-plane boss extrude.
    Total depth in mm (extends depth/2 in both directions).
    """
    _require_part()
    
    half_depth = depth / 2000.0  # mm to meters, then half
    
    # T1 = 6 for mid-plane
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
# CUT FEATURES
# ============================================================

def cut_extrude(depth):
    """
    Cut-Extrude into the solid.
    Depth in mm.
    """
    _require_part()
    
    depth_m = depth / 1000.0
    
    _fm().FeatureCut3(
        True, False, False,
        0, 0,
        depth_m, 0,
        False, False, False, False,
        0, 0,
        False, False, False, False,
        False, False, False,
        False, False, False
    )

    return f"Cut extruded {depth}mm"


def cut_through_all():
    """
    Through-all cut.
    """
    _require_part()
    
    # T1 = 1 for Through All
    _fm().FeatureCut3(
        True, False, False,
        1, 0,
        0, 0,
        False, False, False, False,
        0, 0,
        False, False, False, False,
        False, False, False,
        False, False, False
    )

    return "Cut through all"

# ============================================================
# REVOLVE FEATURE
# ============================================================

def revolve(angle=360):
    """
    Revolve boss feature.
    Requires axis/centerline to be in sketch.
    Angle in degrees.
    """
    _require_part()
    
    _fm().FeatureRevolve2(
        True, True,
        False, False, False, False,
        0, 0,
        math.radians(angle), 0,
        False, False,
        0, 0, 0,
        False, False, False
    )

    return f"Revolved {angle} degrees"

# ============================================================
# FILLET & CHAMFER
# ============================================================

def fillet(radius):
    """
    Fillet selected edges.
    Radius in mm.
    """
    _require_part()
    
    r = radius / 1000.0
    
    _fm().InsertFeatureFillet(
        195,    # Options
        r,      # Radius
        0,      # Radius2
        0, 0, 0, 0,
        0, 0, 0, 0
    )

    return f"Fillet applied: radius {radius}mm"


def chamfer(distance, angle=45):
    """
    Distance-angle chamfer on selected edges.
    Distance in mm, angle in degrees.
    """
    _require_part()
    
    d = distance / 1000.0
    
    _fm().InsertFeatureChamfer(
        4,      # Type
        d,      # Distance
        math.radians(angle),
        0, 0, 0, 0, 0
    )

    return f"Chamfer applied: {distance}mm @ {angle} deg"

# ============================================================
# PATTERN FEATURES
# ============================================================

def linear_pattern(count, spacing):
    """
    Linear pattern of selected feature.
    Direction reference must be selected.
    Spacing in mm.
    """
    _require_part()
    
    s = spacing / 1000.0
    
    _fm().FeatureLinearPattern3(
        count,      # instances dir 1
        1,          # instances dir 2 (unused)
        s,          # spacing dir 1
        0,
        False, False,
        "", "",
        False, False,
        True
    )

    return f"Linear pattern: {count} instances @ {spacing}mm"


def circular_pattern(count, angle=360):
    """
    Circular pattern of selected feature.
    Axis must be selected.
    Angle in degrees.
    """
    _require_part()
    
    _fm().FeatureCircularPattern3(
        count,
        math.radians(angle),
        False,
        "",
        False,
        True
    )

    return f"Circular pattern: {count} instances over {angle} deg"


def mirror_feature():
    """
    Mirror selected feature about selected plane.
    """
    _require_part()
    
    _fm().InsertMirrorFeature2(
        False,   # mirror bodies
        True,    # mirror features
        False,   # geometry pattern
        False
    )

    return "Feature mirrored"

# ============================================================
# FEATURE INFO
# ============================================================

def get_feature_count():
    """
    Returns total feature count.
    """
    _require_part()
    return _model().GetFeatureCount(False)
