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
    model = _model()
    # Safe property access for GetType
    val = model.GetType
    if callable(val):
        val = val()
    if val != 1:
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
    model = _model()
    nothing = get_nothing()
    
    # Make sure a sketch is selected
    model.ClearSelection2(True)
    last_sketch = ""
    for i in range(1, 40): # Scan up to 40 sketches
        sketch_name = f"Sketch{i}"
        if model.Extension.SelectByID2(sketch_name, "SKETCH", 0, 0, 0, False, 0, nothing, 0):
            last_sketch = sketch_name
    
    if last_sketch:
        model.ClearSelection2(True)
        model.Extension.SelectByID2(last_sketch, "SKETCH", 0, 0, 0, False, 0, nothing, 0)
    
    # FeatureExtrusion2 parameters
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
    """Mid-plane boss extrude."""
    _require_part()
    half_depth = depth / 2000.0
    model = _model()
    nothing = get_nothing()
    
    model.ClearSelection2(True)
    last_sketch = ""
    for i in range(1, 40):
        sketch_name = f"Sketch{i}"
        if model.Extension.SelectByID2(sketch_name, "SKETCH", 0, 0, 0, False, 0, nothing, 0):
            last_sketch = sketch_name
            
    if last_sketch:
        model.ClearSelection2(True)
        model.Extension.SelectByID2(last_sketch, "SKETCH", 0, 0, 0, False, 0, nothing, 0)

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
    model = _model()
    nothing = get_nothing()
    
    try: model.SelectionManager.EnableContourSelection = False
    except: pass
    
    model.ClearSelection2(True)
    last_sketch = ""
    for i in range(1, 40):
        sketch_name = f"Sketch{i}"
        if model.Extension.SelectByID2(sketch_name, "SKETCH", 0, 0, 0, False, 0, nothing, 0):
            last_sketch = sketch_name
    
    if last_sketch:
        model.ClearSelection2(True)
        model.Extension.SelectByID2(last_sketch, "SKETCH", 0, 0, 0, False, 0, nothing, 0)
    
    # FIX: T1=0 (Blind) so it uses depth_m
    _fm().FeatureCut4(
        True, False, False, 
        0, 0,               # T1 = 0 (Blind Cut)
        depth_m, 0,         # D1 = Depth
        False, False, False, False, 0, 0, False, False, False, False, False, 
        True, True, False, False, False, 0, 0, False, False
    )
    return f"Cut extruded {depth}mm"


def cut_through_all():
    """Through-all cut."""
    _require_part()
    _fm().FeatureCut4(
        True, False, False, 
        1, 0,               # T1 = 1 (Through All)
        0, 0, 
        False, False, False, False, 0, 0, False, False, False, False, False, 
        True, True, False, False, False, 0, 0, False, False
    )
    return "Cut through all"

# ============================================================
# REVOLVE, FILLET, PATTERNS
# ============================================================

def revolve(angle=360):
    _require_part()
    _fm().FeatureRevolve2(True, True, False, False, False, False, 0, 0, math.radians(angle), 0, False, False, 0, 0, 0, False, False, False)
    return f"Revolved {angle} degrees"


def fillet(radius):
    _require_part()
    r = radius / 1000.0
    _fm().InsertFeatureFillet(195, r, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    return f"Fillet applied: radius {radius}mm"


def chamfer(distance, angle=45):
    _require_part()
    d = distance / 1000.0
    _fm().InsertFeatureChamfer(4, d, math.radians(angle), 0, 0, 0, 0, 0)
    return f"Chamfer applied: {distance}mm @ {angle} deg"


def linear_pattern(count, spacing):
    _require_part()
    s = spacing / 1000.0
    _fm().FeatureLinearPattern3(count, 1, s, 0, False, False, "", "", False, False, True)
    return f"Linear pattern: {count} x {spacing}mm"


def circular_pattern(count, angle=360):
    _require_part()
    _fm().FeatureCircularPattern3(count, math.radians(angle), False, "", False, True)
    return f"Circular pattern: {count} over {angle} deg"


def mirror_feature():
    _require_part()
    _fm().InsertMirrorFeature2(False, True, False, False)
    return "Feature mirrored"


def get_feature_count():
    _require_part()
    return _model().GetFeatureCount(False)