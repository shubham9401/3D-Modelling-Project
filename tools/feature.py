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
    from .solidworks_app import get_active_model
except ImportError:
    from solidworks_app import get_active_model

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
    if _model().GetType() != 1:
        raise Exception("Active document is not a PART")

def _require_last_sketch_closed():
    """
    Finds the last sketch feature and ensures it is closed.
    Works even after exiting sketch mode.
    """
    model = _model()
    feat = model.FirstFeature()
    last_sketch = None

    while feat:
        if feat.GetTypeName2() == "ProfileFeature":
            last_sketch = feat.GetSpecificFeature2()
        feat = feat.GetNextFeature()

    if last_sketch is None:
        raise Exception("No sketch found for feature creation")

    if not last_sketch.IsClosed():
        raise Exception("Last sketch is not a closed profile")

# ============================================================
# EXTRUDE FEATURES
# ============================================================

def extrude(depth):
    """
    Blind boss extrude
    """
    _require_part()
    _require_last_sketch_closed()

    _fm().FeatureExtrusion2(
        True, False, False,
        0, 0,
        depth / 1000, 0,
        False, False, False, False,
        0, 0,
        True, True, True,
        False, False, False,
        0, 0, False
    )

    return f"Extruded {depth} mm"


def extrude_midplane(depth):
    """
    Mid-plane boss extrude
    """
    _require_part()
    _require_last_sketch_closed()

    _fm().FeatureExtrusion2(
        True, False, False,
        6, 0,
        depth / 2000, 0,
        False, False, False, False,
        0, 0,
        True, True, True,
        False, False, False,
        0, 0, False
    )

    return f"Mid-plane extruded {depth} mm"

# ============================================================
# CUT FEATURES
# ============================================================

def cut_extrude(depth):
    """
    Blind cut extrude
    """
    _require_part()
    _require_last_sketch_closed()

    _fm().FeatureCut3(
        True, False, False,
        0, 0,
        depth / 1000, 0,
        False, False, False, False,
        0, 0,
        True, True, True,
        False, False, False,
        False, False, False
    )

    return f"Cut extruded {depth} mm"


def cut_through_all():
    """
    Through-all cut
    """
    _require_part()
    _require_last_sketch_closed()

    _fm().FeatureCut3(
        True, False, False,
        1, 0,
        0, 0,
        False, False, False, False,
        0, 0,
        True, True, True,
        False, False, False,
        False, False, False
    )

    return "Cut through all"

# ============================================================
# REVOLVE FEATURE
# ============================================================

def revolve(angle=360):
    """
    Revolve boss feature
    Requires axis/centerline to be selected
    """
    _require_part()
    _require_last_sketch_closed()

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
    Fillet selected edges
    """
    _require_part()

    _fm().InsertFeatureFillet(
        195,
        radius / 1000,
        0,
        0, 0, 0, 0,
        0, 0, 0, 0
    )

    return f"Fillet applied: radius {radius} mm"


def chamfer(distance, angle=45):
    """
    Distance-angle chamfer
    """
    _require_part()

    _fm().InsertFeatureChamfer(
        4,
        distance / 1000,
        math.radians(angle),
        0, 0, 0, 0, 0
    )

    return f"Chamfer applied: {distance} mm @ {angle}°"

# ============================================================
# PATTERN FEATURES (IMPORTANT)
# ============================================================

def linear_pattern(count, spacing):
    """
    Linear pattern of selected feature
    Direction reference must be selected
    """
    _require_part()

    _fm().FeatureLinearPattern3(
        count,              # instances dir 1
        1,                  # instances dir 2 (unused)
        spacing / 1000,     # spacing dir 1
        0,
        False, False,
        "", "",
        False, False,
        True
    )

    return f"Linear pattern: {count} instances @ {spacing} mm"


def circular_pattern(count, angle=360):
    """
    Circular pattern of selected feature
    Axis must be selected
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

    return f"Circular pattern: {count} instances over {angle}°"


def mirror_feature():
    """
    Mirror selected feature about selected plane
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
    Returns total feature count
    """
    _require_part()
    return _model().GetFeatureCount(False)
