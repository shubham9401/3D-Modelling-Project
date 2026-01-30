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

def revolve(angle=360, profile_name=None, axis_name="Line1"):
    """
    Revolve boss feature with proper profile and axis selection.
    
    IMPORTANT: This function handles the complete revolve workflow:
    1. Exits the current sketch (if active)
    2. Selects the profile (semicircle/arc) 
    3. Selects the axis line with correct mark value (16)
    4. Executes the revolve
    
    Args:
        angle: Rotation angle in degrees (default 360 for full revolution)
        profile_name: Name of the sketch segment to revolve (auto-detects if None)
        axis_name: Name of the axis line (default "Line1")
    """
    _require_part()
    
    model = _model()
    nothing = get_nothing()
    
    # Step 1: Ensure we are in an active sketch (or select it)
    # Since validate_closed_profile now keeps sketch active, we proceed.
    
    # Step 2: Clear any existing selection
    model.ClearSelection2(True)
    
    # Step 3: Select the PROFILE (Mark = 0)
    # IMPORTANT: For a solid revolve, we need a CLOSED profile.
    # So we must select BOTH the Arc AND the Diameter Line (which is also the axis).
    
    # 3a. Select the curved part (Arc)
    if profile_name:
        profile_candidates = [profile_name]
    else:
        profile_candidates = ["Arc1", "Circle1", "Arc2", "Line2"] # Common names
    
    profile_found = False
    used_profile = None
    
    for candidate in profile_candidates:
        if model.Extension.SelectByID2(candidate, "SKETCHSEGMENT", 0, 0, 0, False, 0, nothing, 0):
            profile_found = True
            used_profile = candidate
            break
            
    if not profile_found:
         # Try selecting just the sketch itself if possible
         pass 

    # 3b. Select the Axis Line AS PART OF THE PROFILE (Mark 0) to close the loop
    # We try multiple line names because if draw_centerline was used, Line1 might be construction.
    # The diameter line might be Line2 (or Line3). 
    # Selecting extra lines usually doesn't hurt if they are part of the chain or don't exist.
    line_candidates = [axis_name]
    if axis_name == "Line1":
        line_candidates.extend(["Line2", "Line3"])
        
    for line_name in line_candidates:
         model.Extension.SelectByID2(line_name, "SKETCHSEGMENT", 0, 0, 0, True, 0, nothing, 0)
    
    # Step 4: Select the AXIS (centerline) with Mark = 16, Append = True
    # We try multiple candidates for the axis too!
    # If Line1 (Centerline) fails, we use coordinate fallback.
    # REMOVED Line2/Line3 candidates because for Cone, Line2 is the Base, which causes Bicone result.
    axis_candidates = [axis_name]
    # if axis_name == "Line1":
    #    axis_candidates.extend(["Line2", "Line3"]) # BAD IDEA for Cones
        
    axis_selected = False
    for ax in axis_candidates:
        if model.Extension.SelectByID2(ax, "SKETCHSEGMENT", 0, 0, 0, True, 16, nothing, 0):
            axis_selected = True
            axis_name = ax # Update used axis name
            break
    
    if not axis_selected:
        # One last desperate try: Select by coordinate slightly OFF origin along Y axis.
        # This targets the VERTICAL line (Line1) and avoids the horizontal one (Line2).
        # 0.001m = 1mm. Most sketches are larger than 1mm.
        if model.Extension.SelectByID2("", "SKETCHSEGMENT", 0, 0.001, 0, True, 16, nothing, 0):
             axis_selected = True
             axis_name = "VerticalLine_Pos"
        elif model.Extension.SelectByID2("", "SKETCHSEGMENT", 0, -0.001, 0, True, 16, nothing, 0):
             axis_selected = True
             axis_name = "VerticalLine_Neg"

    if not axis_selected:
        raise Exception(f"Failed to select axis. Tried names: {axis_candidates} and Vertical coordinates.")
    
    # Step 5: Execute the revolve
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
    return f"Revolved {angle} degrees (profile={used_profile}, axis={axis_name})"


def revolve_simple(angle=360):
    """
    Simple revolve - assumes profile and axis are already selected.
    Use this if you've manually selected the profile (mark=0) and axis (mark=16).
    """
    _require_part()
    
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
# SHELL & LOFT
# ============================================================

def shell(thickness):
    """
    Shell feature - hollows out a solid body.
    
    IMPORTANT: Pre-select the face(s) to REMOVE before calling this!
    Use select_face_at_coordinate() to select the top face first.
    
    Based on VBA: Part.InsertFeatureShell(thickness, outward)
    
    Args:
        thickness: Wall thickness in mm
    """
    _require_part()
    model = _model()
    
    t = thickness / 1000.0  # Convert mm to meters
    
    # InsertFeatureShell is on ModelDoc2, NOT FeatureManager!
    # Parameters: 
    #   Thickness (double in meters)
    #   Outward (bool): False = shell inward (normal for cups), True = shell outward
    try:
        result = model.InsertFeatureShell(t, False)
        # InsertFeatureShell may return None on success - don't check return value
        # VBA macro also doesn't check return value
    except Exception as e:
        raise Exception(f"Shell failed with error: {e}")
    
    return f"Shell: {thickness}mm walls"

def loft():
    """
    Loft feature - creates smooth transition between selected profiles.
    Requires: Two or more sketches selected before calling.
    """
    _require_part()
    
    _fm().InsertProtrusionBlend(
        False,  # Closed
        True,   # KeepTangency  
        False,  # ForceNonRational
        1.0,    # TightnessFactor
        0,      # StartTangentType
        0,      # EndTangentType
        False,  # IsThinBody
        0, 0,   # Thickness1, 2
        0,      # ThicknessType
        True,   # UseFeatScope
        False   # PropagateFeatureToParts
    )
    return "Loft created between profiles"

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