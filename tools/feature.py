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
    Use select_sketch() to select the sketches first.
    """
    _require_part()
    fm = _fm()
    
    # InsertProtrusionBlend parameters from VBA:
    # InsertProtrusionBlend(bClosed, bKeepTangency, bForceNonRational, dTightnessFactor, 
    #                       nStartTangentType, nEndTangentType, dStartTangentLength, dEndTangentLength,
    #                       bStartMatchingFaces, bEndMatchingFaces, bIsThinBody, 
    #                       dThickness1, dThickness2, nThicknessType, bUseFeatScope, 
    #                       bUseAutoSelect, bPropagateFeatureToParts)
    result = fm.InsertProtrusionBlend(
        False,  # bClosed
        True,   # bKeepTangency  
        False,  # bForceNonRational
        1,      # dTightnessFactor
        0,      # nStartTangentType
        0,      # nEndTangentType
        1,      # dStartTangentLength
        1,      # dEndTangentLength
        True,   # bStartMatchingFaces
        True,   # bEndMatchingFaces
        False,  # bIsThinBody
        0,      # dThickness1
        0,      # dThickness2
        0,      # nThicknessType
        True,   # bUseFeatScope
        True,   # bUseAutoSelect
        True    # bPropagateFeatureToParts
    )
    
    if result:
        return "Loft created between profiles"
    else:
        raise Exception("Loft creation failed. Make sure you have selected at least 2 sketches.")

def sweep():
    """
    Sweep feature - sweeps a profile sketch along a path sketch.
    
    Requires:
    - Profile sketch selected with mark=1
    - Path sketch selected with mark=4
    
    Use select_sketch() to select the sketches first:
    1. select_sketch("ProfileSketch", mark=1, append=False)
    2. select_sketch("PathSketch", mark=4, append=True)
    3. sweep()
    """
    _require_part()
    fm = _fm()
    model = _model()
    
    feat_count_before = fm.GetFeatureCount(True)
    
    # STRATEGY 1: InsertProtrusionSweep2 (most reliable for modern SW)
    try:
        print("    DEBUG: Attempting Strategy 1 (InsertProtrusionSweep2)...")
        # Parameters: (Propagate, Alignment, TwistCtrl, MergeBodies, 
        #              AlignWithEndFaces, AdvancedSmoothing, StartMatchingType,
        #              EndMatchingType, IsThinBody, Thickness1, Thickness2)
        result = fm.InsertProtrusionSweep2(
            False,  # Propagate
            0,      # Alignment (0 = None/FollowPath)
            0,      # TwistCtrl (0 = Follow Path)
            True,   # MergeBodies
            False,  # AlignWithEndFaces
            False,  # AdvancedSmoothing
            0,      # StartMatchingType
            0,      # EndMatchingType
            False,  # IsThinBody
            0,      # Thickness1
            0       # Thickness2
        )
        feat_count_after = fm.GetFeatureCount(True)
        if feat_count_after > feat_count_before:
            model.ForceRebuild3(True)
            return "Sweep created (InsertProtrusionSweep2)"
        print("    DEBUG: Strategy 1 returned but no new feature created")
    except Exception as e:
        print(f"    DEBUG: Strategy 1 failed: {e}")
    
    # STRATEGY 2: InsertProtrusionSweep with varying param counts
    for param_set_name, params in [
        ("7 params", (False, False, 0, False, False, 0, 0)),
        ("4 params", (False, 0, 0, False)),
        ("5 params", (False, 0, 0, False, True)),
        ("3 params", (False, 0, 0)),
    ]:
        try:
            print(f"    DEBUG: Attempting Strategy 2 ({param_set_name})...")
            fm.InsertProtrusionSweep(*params)
            feat_count_after = fm.GetFeatureCount(True)
            if feat_count_after > feat_count_before:
                model.ForceRebuild3(True)
                return f"Sweep created ({param_set_name})"
        except Exception as e:
            print(f"    DEBUG: Strategy 2 ({param_set_name}) failed: {e}")
    
    # STRATEGY 3: CreateFeature with SweepFeatureData
    try:
        print("    DEBUG: Attempting Strategy 3 (CreateFeature scan)...")
        sweep_type_id = None
        for type_id in range(0, 250):
            try:
                swFeatData = fm.CreateDefinition(type_id)
                if swFeatData is not None:
                    try:
                        # Check if this is a sweep definition
                        _ = swFeatData.PathAlignmentType
                        sweep_type_id = type_id
                        print(f"    DEBUG: Found sweep type ID: {type_id}")
                        break
                    except:
                        pass
            except:
                pass
        
        if sweep_type_id is not None:
            swFeatData = fm.CreateDefinition(sweep_type_id)
            try:
                swFeatData.Merge = True
            except:
                pass
            try:
                swFeatData.AutoSelect = True
            except:
                pass
            
            result = fm.CreateFeature(swFeatData)
            if result is not None:
                model.ForceRebuild3(True)
                return "Sweep created (CreateFeature)"
            
            feat_count_after = fm.GetFeatureCount(True)
            if feat_count_after > feat_count_before:
                model.ForceRebuild3(True)
                return "Sweep created (CreateFeature)"
    except Exception as e:
        print(f"    DEBUG: Strategy 3 failed: {e}")
    
    # STRATEGY 4: Try InsertProtrusionSweep3 (some SW versions)
    try:
        print("    DEBUG: Attempting Strategy 4 (InsertProtrusionSweep3)...")
        result = fm.InsertProtrusionSweep3(
            False,  # Propagate  
            False,  # IsThinBody
            0,      # ThinType
            0,      # Thickness1
            0,      # Thickness2
            0,      # TwistCtrl
            0,      # PathAlignmentType
            True,   # MergeSmooth
            0,      # TwistAngle
            False,  # AdvancedSmoothing
            0,      # StartMatchingType
            0,      # EndMatchingType
            False,  # AlignWithEndFaces
            True,   # UseFeatScope
            True    # UseAutoSelect
        )
        feat_count_after = fm.GetFeatureCount(True)
        if feat_count_after > feat_count_before:
            model.ForceRebuild3(True)
            return "Sweep created (InsertProtrusionSweep3)"
    except Exception as e:
        print(f"    DEBUG: Strategy 4 failed: {e}")
    
    # Final Check
    feat_count_after = fm.GetFeatureCount(True)
    if feat_count_after > feat_count_before:
        model.ForceRebuild3(True)
        return "Sweep created"
        
    raise Exception("Sweep creation failed. Make sure profile (mark=1) and path (mark=4) sketches are selected and intersect.")

def create_reference_plane(offset, plane="Front"):
    """
    Creates a reference plane at an offset distance from an existing plane.
    
    Args:
        offset: Distance in mm from the reference plane
        plane: Base plane name - "Front", "Top", or "Right"
    
    Returns:
        Success message with the new plane name
    """
    _require_part()
    model = _model()
    fm = _fm()
    
    # Map plane names to SolidWorks plane names
    plane_map = {
        "Front": "Front Plane",
        "Top": "Top Plane",
        "Right": "Right Plane"
    }
    
    plane_name = plane_map.get(plane, f"{plane} Plane")
    
    # Select the base plane
    nothing = get_nothing()
    model.Extension.SelectByID2(plane_name, "PLANE", 0, 0, 0, False, 0, nothing, 0)
    
    # Convert offset to meters
    offset_m = offset / 1000.0
    
    # InsertRefPlane(FirstConstraint, FirstValue, SecondConstraint, SecondValue, ThirdConstraint, ThirdValue)
    # Constraint 8 = swRefPlaneReferenceConstraint_Parallel with offset
    # Value is the offset distance in meters
    ref_plane = fm.InsertRefPlane(8, offset_m, 0, 0, 0, 0)
    
    model.ClearSelection2(True)
    
    if ref_plane:
        return f"Reference plane created at {offset}mm offset from {plane}"
    else:
        raise Exception(f"Failed to create reference plane at {offset}mm from {plane}")

def select_sketch(sketch_name, mark=0, append=False):
    """
    Selects a sketch by name for use in loft or other operations.
    
    Args:
        sketch_name: Name of the sketch (e.g., "Sketch1", "Sketch2")
        mark: Selection mark (1 for loft profiles, 4 for guide curves)
        append: Whether to append to existing selection (True) or replace (False)
    
    Returns:
        Success message
    """
    _require_part()
    model = _model()
    nothing = get_nothing()
    
    result = model.Extension.SelectByID2(sketch_name, "SKETCH", 0, 0, 0, append, mark, nothing, 0)
    
    if result:
        return f"Selected {sketch_name}"
    else:
        raise Exception(f"Failed to select {sketch_name}")

# ============================================================
# REFINEMENTS
# ============================================================

def fillet(radius):
    """
    Fillet selected edges.
    
    Args:
        radius: Fillet radius in mm
    """
    import win32com.client
    import pythoncom
    
    _require_part()
    model = _model()
    fm = _fm()
    
    r = radius / 1000.0  # Convert mm to meters
    
    # Check edges are selected
    selMgr = model.SelectionManager
    sel_count = selMgr.GetSelectedObjectCount2(-1)
    if sel_count == 0:
        raise Exception("No edges selected!")
    
    print(f"    DEBUG: {sel_count} edge(s) selected for fillet, radius={r}m")
    
    # Get feature count before to verify fillet creation
    feat_count_before = fm.GetFeatureCount(True)
    
    # Create empty VARIANT for None/Nothing values
    nothing = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
    
    # Try different option values for FeatureFillet
    # swFeatureFilletOptions_e:
    # 1 = swFeatureFilletUniformRadius
    # 2 = swFeatureFilletKeepEdge  
    # 4 = swFeatureFilletKeepSurface
    # 64 = swFeatureFilletPropagate
    # 128 = swFeatureFilletFullPreview
    # Commonly used: 1 (uniform radius), 65 (uniform + propagate)
    
    options_to_try = [1, 65, 193, 195, 0, 64, 128, 3]
    
    for opts in options_to_try:
        try:
            print(f"    DEBUG: Trying FeatureFillet with Options={opts}...")
            result = fm.FeatureFillet(opts, r, 0, 0, nothing, nothing, nothing)
            
            # Check if feature count increased
            feat_count_after = fm.GetFeatureCount(True)
            if feat_count_after > feat_count_before:
                print(f"    DEBUG: Feature count increased from {feat_count_before} to {feat_count_after}")
                return f"Fillet: {radius}mm applied (Options={opts})"
            
            if result is not None:
                return f"Fillet: {radius}mm applied"
                
        except Exception as e:
            print(f"    DEBUG: FeatureFillet Options={opts} failed: {e}")
    
    # Try FeatureFillet3 with different options
    for opts in [1, 65, 193, 0]:
        try:
            print(f"    DEBUG: Trying FeatureFillet3 with Options={opts}...")
            result = fm.FeatureFillet3(
                opts, r, 0, 0,
                nothing, nothing, nothing, nothing, nothing,
                False, False
            )
            
            feat_count_after = fm.GetFeatureCount(True)
            if feat_count_after > feat_count_before:
                return f"Fillet: {radius}mm applied (FF3 Options={opts})"
                
            if result is not None:
                return f"Fillet: {radius}mm applied"
        except Exception as e:
            print(f"    DEBUG: FeatureFillet3 Options={opts} failed: {e}")
    
    # Try using SimpleFillet via ISimpleFilletFeatureData2
    for feat_type in [47, 52, 148, 149, 150]:
        try:
            print(f"    DEBUG: Trying CreateDefinition({feat_type})...")
            swFeatData = fm.CreateDefinition(feat_type)
            if swFeatData is not None:
                print(f"    DEBUG: CreateDefinition({feat_type}) returned object")
                try:
                    swFeatData.Initialize(0)
                except:
                    pass
                try:
                    swFeatData.DefaultRadius = r
                except:
                    pass
                try:
                    # Get edges from selection
                    edges = []
                    for i in range(1, sel_count + 1):
                        edge = selMgr.GetSelectedObject6(i, -1)
                        if edge:
                            edges.append(edge)
                    if edges:
                        swFeatData.Edges = tuple(edges)
                except:
                    pass
                
                result = fm.CreateFeature(swFeatData)
                feat_count_after = fm.GetFeatureCount(True)
                if feat_count_after > feat_count_before:
                    return f"Fillet: {radius}mm applied (type={feat_type})"
        except Exception as e:
            print(f"    DEBUG: CreateDefinition({feat_type}) failed: {e}")
    
    raise Exception(f"Fillet failed for {radius}mm. Feature was not created in model.")

def chamfer(distance, angle=45):
    """
    Chamfer selected edges.
    
    IMPORTANT: Pre-select the edge(s) before calling this function!
    Use select_edge_at_coordinate() first.
    
    Args:
        distance: Chamfer distance in mm
        angle: Chamfer angle in degrees (default 45)
    """
    _require_part()
    model = _model()
    fm = _fm()
    
    d = distance / 1000.0  # Convert mm to meters
    a = math.radians(angle)  # Convert degrees to radians
    
    # Check edges are selected
    selMgr = model.SelectionManager
    sel_count = selMgr.GetSelectedObjectCount2(-1)
    if sel_count == 0:
        raise Exception("No edges selected for chamfer!")
    
    print(f"    DEBUG: {sel_count} edge(s) selected for chamfer, distance={d}m, angle={angle}deg")
    
    # Get feature count before to verify chamfer creation
    feat_count_before = fm.GetFeatureCount(True)
    
    # VBA signature from user: InsertFeatureChamfer(swConstRadiusFillet, 6, 1, 0.01, 0.78539816339745, 0, 0, 0, 0)
    # That's 9 parameters but we got "Invalid number of parameters"
    # Let's try different parameter counts
    
    # Try 5 parameters: InsertFeatureChamfer(Type, Options, Distance, Angle, SecondDist)
    try:
        print("    DEBUG: Trying InsertFeatureChamfer with 5 params...")
        result = fm.InsertFeatureChamfer(0, 4, d, a, d)
        feat_count_after = fm.GetFeatureCount(True)
        if feat_count_after > feat_count_before:
            return f"Chamfer: {distance}mm @ {angle}deg applied (5 params)"
    except Exception as e:
        print(f"    DEBUG: 5 params failed: {e}")
    
    # Try 6 parameters
    try:
        print("    DEBUG: Trying InsertFeatureChamfer with 6 params...")
        result = fm.InsertFeatureChamfer(0, 4, 1, d, a, d)
        feat_count_after = fm.GetFeatureCount(True)
        if feat_count_after > feat_count_before:
            return f"Chamfer: {distance}mm @ {angle}deg applied (6 params)"
    except Exception as e:
        print(f"    DEBUG: 6 params failed: {e}")
    
    # Try 7 parameters (SolidWorks 2020+ style)
    try:
        print("    DEBUG: Trying InsertFeatureChamfer with 7 params...")
        result = fm.InsertFeatureChamfer(0, 6, 1, d, a, 0, 0)
        feat_count_after = fm.GetFeatureCount(True)
        if feat_count_after > feat_count_before:
            return f"Chamfer: {distance}mm @ {angle}deg applied (7 params)"
    except Exception as e:
        print(f"    DEBUG: 7 params failed: {e}")
    
    # Try 8 parameters
    try:
        print("    DEBUG: Trying InsertFeatureChamfer with 8 params...")
        result = fm.InsertFeatureChamfer(0, 6, 1, d, a, 0, 0, 0)
        feat_count_after = fm.GetFeatureCount(True)
        if feat_count_after > feat_count_before:
            # Rebuild model to refresh graphics
            try:
                model.ForceRebuild3(True)
            except:
                try:
                    model.EditRebuild3()
                except:
                    pass
            return f"Chamfer: {distance}mm @ {angle}deg applied (8 params)"
    except Exception as e:
        print(f"    DEBUG: 8 params failed: {e}")
    
    # Try 10 parameters (some versions)
    try:
        print("    DEBUG: Trying InsertFeatureChamfer with 10 params...")
        result = fm.InsertFeatureChamfer(0, 6, 1, d, a, 0, 0, 0, 0, 0)
        feat_count_after = fm.GetFeatureCount(True)
        if feat_count_after > feat_count_before:
            return f"Chamfer: {distance}mm @ {angle}deg applied (10 params)"
    except Exception as e:
        print(f"    DEBUG: 10 params failed: {e}")
    
    # Try FeatureChamfer (alternative method)
    try:
        print("    DEBUG: Trying FeatureChamfer...")
        result = fm.FeatureChamfer(d, a, False)
        feat_count_after = fm.GetFeatureCount(True)
        if feat_count_after > feat_count_before:
            return f"Chamfer: {distance}mm @ {angle}deg applied (FeatureChamfer)"
    except Exception as e:
        print(f"    DEBUG: FeatureChamfer failed: {e}")
    
    # Try the exact VBA signature
    try:
        print("    DEBUG: Trying exact VBA signature...")
        # swConstRadiusFillet = 0, options = 6, count = 1, distance = 0.01, angle = 0.785...
        result = fm.InsertFeatureChamfer(0, 6, 1, d, a, 0, 0, 0, 0)
        feat_count_after = fm.GetFeatureCount(True)
        if feat_count_after > feat_count_before:
            return f"Chamfer: {distance}mm @ {angle}deg applied (VBA style)"
    except Exception as e:
        print(f"    DEBUG: VBA style failed: {e}")
    
    raise Exception(f"Chamfer failed for {distance}mm. All parameter combinations failed.")

def linear_pattern(count, spacing):
    """Linear pattern."""
    _require_part()
    s = spacing / 1000.0
    _fm().FeatureLinearPattern3(count, 1, s, 0, False, False, "", "", False, False, True)
    return f"Linear pattern: {count} x {spacing}mm"

def circular_pattern(count, angle=360):
    """
    Circular pattern - patterns the LAST feature around the Y-axis (vertical).
    
    Creates a reference axis through the origin if needed, then patterns
    the last extrude/cut feature around it.
    
    Args:
        count: Number of instances (including original)
        angle: Total angle span in degrees (default 360 for full circle)
    """
    _require_part()
    model = _model()
    fm = _fm()
    nothing = get_nothing()
    
    print(f"    DEBUG: Circular pattern: {count} instances over {angle}°")
    
    feat_count_before = fm.GetFeatureCount(True)
    
    # NOTE: We select the cylindrical face of the base body as axis reference
    # in STEP 3 below (via SelectByRay). No need to find/create named axes.
    
    # ══════════════════════════════════════════════════════════════
    # STEP 2: Find the LAST feature to pattern
    # ══════════════════════════════════════════════════════════════
    last_feature_name = None
    
    # Try common feature names — DESCENDING order, BREAK on first match
    common_names = [
        "Boss-Extrude10", "Boss-Extrude9", "Boss-Extrude8", "Boss-Extrude7",
        "Boss-Extrude6", "Boss-Extrude5", "Boss-Extrude4", "Boss-Extrude3",
        "Boss-Extrude2",
        "Cut-Extrude10", "Cut-Extrude9", "Cut-Extrude8", "Cut-Extrude7",
        "Cut-Extrude6", "Cut-Extrude5", "Cut-Extrude4", "Cut-Extrude3",
        "Cut-Extrude2", "Cut-Extrude1",
        "Sweep1", "Revolve1", "Loft1",
        "Boss-Extrude1",
    ]
    
    for name in common_names:
        try:
            if model.Extension.SelectByID2(name, "BODYFEATURE", 0, 0, 0, False, 0, nothing, 0):
                last_feature_name = name
                print(f"    DEBUG: Feature found: '{name}'")
                model.ClearSelection2(True)
                break  # Take first match = highest numbered = LAST created
        except:
            pass
    
    if last_feature_name is None:
        raise Exception("No feature found to pattern!")
    
    print(f"    DEBUG: Will pattern '{last_feature_name}' using cylindrical face as axis")
    
    # ══════════════════════════════════════════════════════════════
    # STEP 3: Select feature (mark=4) and axis (mark=1), then execute
    # ══════════════════════════════════════════════════════════════
    model.ClearSelection2(True)
    
    # Select the feature to pattern (mark = 4)
    feat_selected = model.Extension.SelectByID2(
        last_feature_name, "BODYFEATURE", 0, 0, 0, False, 4, nothing, 0
    )
    print(f"    DEBUG: Feature selection: {feat_selected}")
    
    # Select the CYLINDRICAL FACE of the base body as axis reference (mark=1, append=True)
    # When you select a cylindrical face for circular pattern, SolidWorks
    # automatically uses the cylinder's central axis. This is the ONLY reliable method.
    axis_selected = False
    
    # Get bounding box to know the part height and radius for ray targeting
    try:
        box = model.GetPartBox()
        if box:
            # box = [xmin, ymin, zmin, xmax, ymax, zmax] in meters
            mid_y = (box[1] + box[4]) / 2  # mid-height
            max_x = box[3]  # rightmost edge
            max_z = box[5]  # front edge
            print(f"    DEBUG: BBox mid_y={mid_y*1000:.1f}mm, max_x={max_x*1000:.1f}mm")
        else:
            mid_y = 0.005  # default 5mm
            max_x = 0.03   # default 30mm
            max_z = 0.03
    except:
        mid_y = 0.005
        max_x = 0.03
        max_z = 0.03
    
    # Try selecting the cylindrical face from multiple directions
    ray_attempts = [
        # (origin_x, origin_y, origin_z, dir_x, dir_y, dir_z, description)
        (max_x + 0.01, mid_y, 0, -1, 0, 0, "from +X"),      # From right side
        (-max_x - 0.01, mid_y, 0, 1, 0, 0, "from -X"),       # From left side  
        (0, mid_y, max_z + 0.01, 0, 0, -1, "from +Z"),        # From front
        (0, mid_y, -max_z - 0.01, 0, 0, 1, "from -Z"),        # From back
    ]
    
    for ox, oy, oz, dx, dy, dz, desc in ray_attempts:
        try:
            axis_selected = model.Extension.SelectByRay(
                ox, oy, oz,    # Ray origin (outside the part)
                dx, dy, dz,    # Ray direction (toward center)
                0.001,         # Radius
                2,             # Type: 2 = FACE
                True,          # Append to existing selection  
                1,             # Mark = 1 (axis reference for pattern)
                0              # Option
            )
            if axis_selected:
                print(f"    DEBUG: Cylindrical face selected {desc} ✅")
                break
        except Exception as e:
            print(f"    DEBUG: Ray {desc} failed: {e}")
    
    # Fallback: try named axis entities (rarely works but worth trying)
    if not axis_selected:
        for fallback in ["Axis1", "Axis2", "Y Axis"]:
            try:
                axis_selected = model.Extension.SelectByID2(
                    fallback, "AXIS", 0, 0, 0, True, 1, nothing, 0
                )
                if axis_selected:
                    print(f"    DEBUG: Named axis '{fallback}' selected")
                    break
            except:
                pass
    
    if not feat_selected:
        raise Exception(f"Failed to select feature '{last_feature_name}' for pattern")
    if not axis_selected:
        raise Exception("Failed to select cylindrical face or axis for circular pattern.")
    
    # Execute circular pattern
    for strategy_name, strategy_fn in [
        ("FeatureCircularPattern4", lambda: fm.FeatureCircularPattern4(
            count, math.radians(angle), False, "", False, True
        )),
        ("FeatureCircularPattern3", lambda: fm.FeatureCircularPattern3(
            count, math.radians(angle), False, "", False, True
        )),
    ]:
        try:
            print(f"    DEBUG: Trying {strategy_name}...")
            strategy_fn()
            
            feat_count_after = fm.GetFeatureCount(True)
            if feat_count_after > feat_count_before:
                model.ForceRebuild3(True)
                print(f"    DEBUG: {strategy_name} SUCCESS! Features: {feat_count_before} → {feat_count_after}")
                return f"Circular pattern: {count} instances over {angle}°"
            else:
                print(f"    DEBUG: {strategy_name} returned OK but no new feature created")
        except Exception as e:
            print(f"    DEBUG: {strategy_name} failed: {e}")
    
    raise Exception(f"Circular pattern failed. Feature and axis were selected but pattern creation failed.")

def mirror_feature():
    """Mirror feature."""
    _require_part()
    _fm().InsertMirrorFeature2(False, True, False, False)
    return "Mirrored"

def get_feature_count():
    """Returns feature count."""
    _require_part()
    return _model().GetFeatureCount(False)

def thread(diameter=6, pitch=1.0, depth=10, size=None, right_handed=True, thread_method="cut"):
    """
    Creates a thread feature on the selected cylindrical edge.
    
    IMPORTANT: Pre-select the circular edge before calling this function!
    Use select_edge_at_coordinate() first.
    
    Args:
        diameter: Thread diameter in mm (default 6 for M6)
        pitch: Thread pitch in mm (default 1.0)
        depth: Thread depth in mm (default 10)
        size: Thread size string like "M6x1.0" (auto-generated if None)
        right_handed: True for right-hand thread, False for left-hand (default True)
        thread_method: "cut" for cut thread, "extrude" for extruded thread (default "cut")
    """
    _require_part()
    model = _model()
    fm = _fm()
    
    # Convert to meters
    d = diameter / 1000.0
    p = pitch / 1000.0
    depth_m = depth / 1000.0
    
    # Generate size string if not provided
    if size is None:
        size = f"M{int(diameter)}x{pitch}"
    
    # Check edge is selected
    selMgr = model.SelectionManager
    sel_count = selMgr.GetSelectedObjectCount2(-1)
    if sel_count == 0:
        raise Exception("No edge selected for thread! Select a circular edge first.")
    
    print(f"    DEBUG: {sel_count} edge(s) selected for thread, size={size}, depth={depth}mm")
    
    # Get feature count before
    feat_count_before = fm.GetFeatureCount(True)
    
    # Scan a wide range to find valid feature type IDs
    # First, find all valid IDs that return non-None objects
    print("    DEBUG: Scanning for valid CreateDefinition IDs...")
    valid_ids = []
    for type_id in range(0, 300):
        try:
            swFeatData = fm.CreateDefinition(type_id)
            if swFeatData is not None:
                valid_ids.append(type_id)
        except:
            pass
    
    print(f"    DEBUG: Found {len(valid_ids)} valid IDs: {valid_ids[:20]}...")  # Show first 20
    
    # Now try each valid ID and check if it has thread-related methods
    for type_id in valid_ids:
        try:
            swFeatData = fm.CreateDefinition(type_id)
            if swFeatData is None:
                continue
                
            # Try to call InitializeThreadData - if it works, this is a thread type
            try:
                swFeatData.InitializeThreadData()
                print(f"    DEBUG: Found thread type at ID {type_id}!")
                
                # Set all properties from VBA macro
                swFeatData.BlindDepth = depth_m
                swFeatData.DiameterOverride = False
                swFeatData.EndCondition = 0  # swThreadEndCondition_Blind
                swFeatData.EndConditionOffset = False
                swFeatData.EndConditionOffsetDistance = 0.001
                swFeatData.EndConditionOffsetReverse = False
                swFeatData.MaintainThreadLength = False
                swFeatData.MirrorProfile = False
                swFeatData.MirrorType = 0  # swThreadMirrorType_Horizontally
                swFeatData.MultipleStart = False
                swFeatData.NumberOfStarts = 2
                swFeatData.Offset = False
                swFeatData.OffsetDistance = 0.001
                swFeatData.PitchOverride = False
                swFeatData.ReverseDirection = False
                swFeatData.ReverseOffset = False
                swFeatData.Revolutions = int(depth / pitch) if pitch > 0 else 10
                swFeatData.RightHanded = right_handed
                swFeatData.RotationAngle = 0
                swFeatData.ThreadMethod = 0 if thread_method == "cut" else 1
                swFeatData.ThreadStartAngle = 0
                swFeatData.TrimEndFace = False
                swFeatData.TrimStartFace = False
                swFeatData.Type = r"C:\ProgramData\SolidWorks\SOLIDWORKS 2025\thread profiles\Metric Die.SLDLFP"
                swFeatData.Diameter = d
                swFeatData.Pitch = p
                swFeatData.Size = size
                
                result = fm.CreateFeature(swFeatData)
                
                feat_count_after = fm.GetFeatureCount(True)
                if feat_count_after > feat_count_before:
                    try:
                        model.ForceRebuild3(True)
                    except:
                        pass
                    return f"Thread {size} created with depth {depth}mm (type_id={type_id})"
                    
            except AttributeError:
                # No InitializeThreadData method, skip
                pass
            except Exception as e:
                # InitializeThreadData exists but failed - still might be thread type
                print(f"    DEBUG: ID {type_id} has issues: {e}")
                
        except Exception as e:
            pass
    
    # Clear selection and notify
    model.ClearSelection2(True)
    
    raise Exception(f"Thread creation failed for {size}. No valid thread feature type found in IDs 0-300.")

def thread_tap(diameter=6, pitch=1.0, depth=10, size=None, right_handed=True):
    """
    Creates an internal thread (tap) feature on the selected circular edge of a hole.
    This is used for NUTS and threaded holes.
    
    IMPORTANT: Pre-select the circular edge of a hole before calling this function!
    Use select_edge_at_coordinate() first.
    
    Args:
        diameter: Thread diameter in mm (default 6 for M6)
        pitch: Thread pitch in mm (default 1.0)
        depth: Thread depth in mm (default 10)
        size: Thread size string like "M6x1.0" (auto-generated if None)
        right_handed: True for right-hand thread, False for left-hand (default True)
    """
    _require_part()
    model = _model()
    fm = _fm()
    
    # Convert to meters
    d = diameter / 1000.0
    p = pitch / 1000.0
    depth_m = depth / 1000.0
    
    # Generate size string if not provided
    if size is None:
        size = f"M{int(diameter)}x{pitch}"
    
    # Check edge is selected
    selMgr = model.SelectionManager
    sel_count = selMgr.GetSelectedObjectCount2(-1)
    if sel_count == 0:
        raise Exception("No edge selected for thread! Select a circular edge of the hole first.")
    
    print(f"    DEBUG: {sel_count} edge(s) selected for tap thread, size={size}, depth={depth}mm")
    
    # Get feature count before
    feat_count_before = fm.GetFeatureCount(True)
    
    # Use the known thread type ID (87 = swFmSweepThread in SW 2025)
    thread_type_id = 87
    
    try:
        print(f"    DEBUG: Trying CreateDefinition({thread_type_id})...")
        swFeatData = fm.CreateDefinition(thread_type_id)
        print(f"    DEBUG: CreateDefinition returned: {swFeatData}")
        
        if swFeatData is None:
            # Fallback: scan for thread type
            print("    DEBUG: Scanning for thread type ID...")
            for type_id in range(0, 300):
                try:
                    swFeatData = fm.CreateDefinition(type_id)
                    if swFeatData is not None:
                        try:
                            swFeatData.InitializeThreadData()
                            thread_type_id = type_id
                            print(f"    DEBUG: Found thread type at ID {type_id}")
                            break
                        except:
                            swFeatData = None
                except:
                    pass
        
        if swFeatData is None:
            raise Exception("Could not find thread feature type")
        
        print(f"    DEBUG: Initializing thread data...")
        swFeatData.InitializeThreadData()
        
        print(f"    DEBUG: Setting thread properties...")
        # Set all properties for internal tap thread
        swFeatData.BlindDepth = depth_m
        swFeatData.DiameterOverride = False
        swFeatData.EndCondition = 0  # swThreadEndCondition_Blind
        swFeatData.EndConditionOffset = False
        swFeatData.EndConditionOffsetDistance = 0.001
        swFeatData.EndConditionOffsetReverse = False
        swFeatData.MaintainThreadLength = False
        swFeatData.MirrorProfile = False
        swFeatData.MirrorType = 0  # swThreadMirrorType_Horizontally
        swFeatData.MultipleStart = False
        swFeatData.NumberOfStarts = 2
        swFeatData.Offset = False
        swFeatData.OffsetDistance = 0.001
        swFeatData.PitchOverride = False
        swFeatData.ReverseDirection = False
        swFeatData.ReverseOffset = False
        swFeatData.Revolutions = int(depth / pitch) if pitch > 0 else 10
        swFeatData.RightHanded = right_handed
        swFeatData.RotationAngle = 0
        swFeatData.ThreadMethod = 0  # swThreadMethod_Cut
        swFeatData.ThreadStartAngle = 0
        swFeatData.TrimEndFace = False
        swFeatData.TrimStartFace = False
        # Use Metric TAP profile for internal threads (nuts)
        swFeatData.Type = r"C:\ProgramData\SolidWorks\SOLIDWORKS 2025\thread profiles\Metric Tap.SLDLFP"
        swFeatData.Diameter = d
        swFeatData.Pitch = p
        swFeatData.Size = size
        
        print(f"    DEBUG: Creating feature...")
        result = fm.CreateFeature(swFeatData)
        print(f"    DEBUG: CreateFeature returned: {result}")
        
        feat_count_after = fm.GetFeatureCount(True)
        print(f"    DEBUG: Feature count before={feat_count_before}, after={feat_count_after}")
        
        if feat_count_after > feat_count_before:
            try:
                model.ForceRebuild3(True)
            except:
                pass
            return f"Tap thread {size} created with depth {depth}mm (internal thread for nut)"
            
    except Exception as e:
        print(f"    DEBUG: Tap thread creation exception: {e}")
        import traceback
        traceback.print_exc()
    
    # Clear selection and notify
    model.ClearSelection2(True)
    
    raise Exception(f"Tap thread creation failed for {size}. Make sure you selected the circular edge of a hole.")

def sheet_metal_base_flange(thickness=1, bend_radius=1, depth=20, reverse_direction=False, k_factor=0.5):
    """
    Creates a sheet metal base flange from the active sketch profile.
    
    IMPORTANT: Draw a closed profile sketch first, then call this function.
    The sketch should be on a plane (Front/Top/Right).
    
    Args:
        thickness: Sheet metal thickness in mm (default 1)
        bend_radius: Default bend radius in mm (default 1)
        depth: Extrusion depth in mm (default 20)
        reverse_direction: Reverse extrusion direction (default False)
        k_factor: K-factor for bend allowance (default 0.5)
    """
    _require_part()
    model = _model()
    fm = _fm()
    
    # Convert to meters
    t = thickness / 1000.0
    r = bend_radius / 1000.0
    d = depth / 1000.0
    
    print(f"    DEBUG: Creating sheet metal base flange: thickness={thickness}mm, bend_radius={bend_radius}mm, depth={depth}mm")
    
    # Get feature count before
    feat_count_before = fm.GetFeatureCount(True)
    
    # swFmBaseFlange = 34 (found from scanning)
    base_flange_id = 34
    
    try:
        print(f"    DEBUG: Creating BaseFlange definition (ID={base_flange_id})...")
        swFeatData = fm.CreateDefinition(base_flange_id)
        
        if swFeatData is None:
            # Scan for the correct ID
            print("    DEBUG: CreateDefinition(34) returned None, scanning...")
            for type_id in range(0, 300):
                try:
                    swFeatData = fm.CreateDefinition(type_id)
                    if swFeatData is not None:
                        try:
                            _ = swFeatData.Thickness
                            _ = swFeatData.BendRadius
                            base_flange_id = type_id
                            print(f"    DEBUG: Found BaseFlange at ID {type_id}")
                            break
                        except:
                            swFeatData = None
                except:
                    pass
        
        if swFeatData is None:
            raise Exception("Could not find BaseFlange feature type")
        
        print(f"    DEBUG: Setting sheet metal properties...")
        
        # Set properties directly - skip Initialize/CustomBendAllowance
        try:
            swFeatData.BendRadius = r
        except Exception as e:
            print(f"    DEBUG: Setting BendRadius failed: {e}")
            
        try:
            swFeatData.D1EndConditionDistance = d
        except Exception as e:
            print(f"    DEBUG: Setting D1EndConditionDistance failed: {e}")
            
        try:
            swFeatData.D1EndConditionType = 1  # Blind
        except:
            pass
            
        try:
            swFeatData.D1ReverseOffset = False
        except:
            pass
            
        try:
            swFeatData.D2EndConditionDistance = d
        except:
            pass
            
        try:
            swFeatData.D2EndConditionType = 1  # Blind
        except:
            pass
            
        try:
            swFeatData.D2ReverseOffset = False
        except:
            pass
            
        try:
            swFeatData.OffsetDirections = 1
        except:
            pass
            
        try:
            swFeatData.ReverseDirection = reverse_direction
        except:
            pass
            
        try:
            swFeatData.ReverseThickness = False
        except:
            pass
            
        try:
            swFeatData.Thickness = t
        except Exception as e:
            print(f"    DEBUG: Setting Thickness failed: {e}")
        
        print("    DEBUG: Creating sheet metal feature...")
        result = fm.CreateFeature(swFeatData)
        print(f"    DEBUG: CreateFeature returned: {result}")
        
        feat_count_after = fm.GetFeatureCount(True)
        print(f"    DEBUG: Feature count before={feat_count_before}, after={feat_count_after}")
        
        if feat_count_after > feat_count_before:
            try:
                model.ForceRebuild3(True)
            except:
                pass
            return f"Sheet metal base flange created: {thickness}mm thick, {depth}mm deep"
            
    except Exception as e:
        print(f"    DEBUG: Sheet metal creation exception: {e}")
        import traceback
        traceback.print_exc()
    
    raise Exception(f"Sheet metal base flange creation failed. Make sure you have an active sketch with a closed profile.")

def edge_flange(length=20, angle=90, gap_distance=0):
    """
    Creates a sheet metal edge flange on the selected edge.
    
    IMPORTANT: Pre-select an edge of an existing sheet metal part first!
    Use select_edge_at_coordinate() to select the edge.
    
    Args:
        length: Flange length in mm (default 20)
        angle: Flange angle in degrees (default 90)
        gap_distance: Gap between flanges in mm (default 0)
    """
    _require_part()
    model = _model()
    fm = _fm()
    
    # Convert to meters and radians
    l = length / 1000.0
    a = math.radians(angle)
    g = gap_distance / 1000.0
    
    print(f"    DEBUG: Creating edge flange: length={length}mm, angle={angle}deg")
    
    # Check edge is selected
    selMgr = model.SelectionManager
    sel_count = selMgr.GetSelectedObjectCount2(-1)
    if sel_count == 0:
        raise Exception("No edge selected for edge flange! Select an edge first.")
    
    # Get feature count before
    feat_count_before = fm.GetFeatureCount(True)
    
    # Find the EdgeFlange feature type ID
    edge_flange_id = None
    
    for type_id in range(0, 300):
        try:
            swFeatData = fm.CreateDefinition(type_id)
            if swFeatData is not None:
                try:
                    # Check for edge flange specific properties
                    _ = swFeatData.FlangeLength
                    _ = swFeatData.Angle
                    edge_flange_id = type_id
                    print(f"    DEBUG: Found EdgeFlange type at ID {type_id}")
                    break
                except:
                    pass
        except:
            pass
    
    if edge_flange_id is None:
        # Try alternative: InsertSheetMetalEdgeFlange
        print("    DEBUG: EdgeFlange type not found, trying InsertSheetMetalEdgeFlange...")
        try:
            result = fm.InsertSheetMetalEdgeFlange2(
                l,      # Length
                a,      # Angle
                0,      # Offset distance
                False,  # Use relief
                False,  # Use gap
                g,      # Gap distance
                0,      # Relief ratio
                0,      # Relief depth
                0,      # Relief width
                1,      # Flange position
                False   # Reverse direction
            )
            
            feat_count_after = fm.GetFeatureCount(True)
            if feat_count_after > feat_count_before:
                try:
                    model.ForceRebuild3(True)
                except:
                    pass
                return f"Edge flange created: {length}mm long at {angle}deg"
                
        except Exception as e:
            print(f"    DEBUG: InsertSheetMetalEdgeFlange2 failed: {e}")
    
    try:
        print(f"    DEBUG: Creating EdgeFlange definition (ID={edge_flange_id})...")
        swFeatData = fm.CreateDefinition(edge_flange_id)
        
        if swFeatData is not None:
            # Set properties
            try:
                swFeatData.FlangeLength = l
            except:
                pass
            try:
                swFeatData.Angle = a
            except:
                pass
            try:
                swFeatData.GapDistance = g
            except:
                pass
            try:
                swFeatData.ReverseDirection = False
            except:
                pass
            
            print("    DEBUG: Creating edge flange feature...")
            result = fm.CreateFeature(swFeatData)
            print(f"    DEBUG: CreateFeature returned: {result}")
            
            feat_count_after = fm.GetFeatureCount(True)
            print(f"    DEBUG: Feature count before={feat_count_before}, after={feat_count_after}")
            
            if feat_count_after > feat_count_before:
                try:
                    model.ForceRebuild3(True)
                except:
                    pass
                return f"Edge flange created: {length}mm long at {angle}deg"
                
    except Exception as e:
        print(f"    DEBUG: Edge flange creation exception: {e}")
        import traceback
        traceback.print_exc()
    
    # Clear selection
    model.ClearSelection2(True)
    
    raise Exception(f"Edge flange creation failed. Make sure you selected an edge of a sheet metal part.")

# ============================================================
# DELETE / MODIFY FEATURES
# ============================================================

def delete_feature(feature_name):
    """
    Delete a feature from the model by name.
    
    Use get_feature_tree to see available feature names first.
    Common names: "Boss-Extrude1", "Boss-Extrude2", "Sketch1", "Fillet1", etc.
    
    Args:
        feature_name: Exact name of the feature to delete (e.g., "Boss-Extrude1")
    """
    import re
    
    _require_part()
    model = _model()
    nothing = get_nothing()
    
    # Step 0: Exit any active sketch using the proven method from exit_sketch()
    try:
        model.InsertSketch2(True)
    except:
        pass
    model.ClearSelection2(True)
    
    # Build list of name candidates: exact name + common SolidWorks variations
    name_candidates = [feature_name]
    
    name_variations = {
        "Extrude": ["Boss-Extrude", "Extrude"],
        "Cut": ["Cut-Extrude", "Cut"],
        "Revolve": ["Boss-Revolve", "Revolve"],
        "Loft": ["Loft-Boss", "Loft"],
        "Sweep": ["Sweep-Boss", "Sweep"],
    }
    
    for key, prefixes in name_variations.items():
        if key.lower() in feature_name.lower():
            num_match = re.search(r'(\d+)', feature_name)
            num = num_match.group(1) if num_match else "1"
            for prefix in prefixes:
                candidate = f"{prefix}{num}"
                if candidate != feature_name and candidate not in name_candidates:
                    name_candidates.append(candidate)
    
    print(f"    DEBUG: Trying to delete feature. Candidates: {name_candidates}")
    
    # Step 1: Try SelectByID2 with each candidate name + type (same as circular_pattern)
    selected = False
    used_name = None
    sel_types = ["BODYFEATURE", "SKETCH", "REFSURFACE", "REFPLANE", "SOLIDBODY"]
    
    for name in name_candidates:
        for sel_type in sel_types:
            try:
                if model.Extension.SelectByID2(name, sel_type, 0, 0, 0, False, 0, nothing, 0):
                    selected = True
                    used_name = name
                    print(f"    DEBUG: Selected '{name}' as {sel_type}")
                    break
            except Exception as e:
                print(f"    DEBUG: SelectByID2('{name}', '{sel_type}') error: {e}")
        if selected:
            break

    if not selected:
        raise Exception(
            f"Could not select feature '{feature_name}'. "
            f"Tried names: {name_candidates} with types: {sel_types}"
        )
    
    # Step 2: Delete the selected feature
    try:
        result = model.Extension.DeleteSelection2(0)  # 0 = delete absorbed features too
        print(f"    DEBUG: DeleteSelection2 result: {result}")
    except Exception as e:
        print(f"    DEBUG: DeleteSelection2 error: {e}")
        raise Exception(f"Failed to delete feature '{used_name}': {e}")
    
    # Step 3: Rebuild
    try:
        model.ForceRebuild3(True)
    except:
        try:
            model.EditRebuild3()
        except:
            pass
    
    model.ClearSelection2(True)
    
    return f"Deleted feature '{used_name}'"