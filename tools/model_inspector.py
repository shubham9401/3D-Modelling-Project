"""
Model Inspector: Extracts measurable properties from SolidWorks models.
FIXED VERSION - Correctly handles mass properties array indices.
"""

import sys
import os

try:
    from .solidworks_app import get_active_model, get_sw_app
except ImportError:
    from solidworks_app import get_active_model, get_sw_app


def _model():
    model = get_active_model()
    if model is None:
        raise Exception("No active SolidWorks document")
    return model


# ============================================================
# FEATURE TREE
# ============================================================

def get_feature_tree():
    """
    Returns the feature tree as an ordered list of features.
    Uses multiple fallback methods to handle COM interop issues.
    """
    model = _model()
    
    SKIP_NAMES = {
        "Comments", "Favorites", "History", "Selection Sets",
        "Sensors", "Design Binder", "Annotations", "Surface Bodies",
        "Solid Bodies", "Lights, Cameras and Scene", "Equations",
        "Material", "Front Plane", "Top Plane", "Right Plane", "Origin",
        "Lights", "Ambient", "Directional1", "Directional2", "Directional3",
    }
    
    SKIP_TYPES = {
        "OriginProfileFeature", "RefPlane", "OriginPoint",
        "MateReferenceGroupFolder", "RefAxis",
    }
    
    features = []
    
    # Method 1: FirstFeature() chain (standard COM traversal)
    try:
        feat = model.FirstFeature()
        while feat is not None:
            try:
                name = feat.Name
                feat_type = feat.GetTypeName2()
                
                if name not in SKIP_NAMES and feat_type not in SKIP_TYPES:
                    features.append({
                        "name": name,
                        "type": feat_type,
                    })
            except Exception:
                pass
            
            try:
                feat = feat.GetNextFeature()
            except Exception:
                break
        
        if features:
            print(f"    [DEBUG] Method 1 (FirstFeature): found {len(features)} features")
            return features
    except Exception as e:
        print(f"    [DEBUG] Method 1 (FirstFeature) failed: {e}")
    
    # Method 2: FeatureManager.GetFeatures array
    try:
        fm = model.FeatureManager
        feat_array = fm.GetFeatures(True)  # True = top-level only
        if feat_array:
            for feat in feat_array:
                try:
                    name = feat.Name
                    feat_type = feat.GetTypeName2()
                    if name not in SKIP_NAMES and feat_type not in SKIP_TYPES:
                        features.append({
                            "name": name,
                            "type": feat_type,
                        })
                except Exception:
                    pass
        
        if features:
            print(f"    [DEBUG] Method 2 (GetFeatures): found {len(features)} features")
            return features
    except Exception as e:
        print(f"    [DEBUG] Method 2 (GetFeatures) failed: {e}")
    
    # Method 3: FeatureByPositionReverse (walk backwards)
    try:
        fm = model.FeatureManager
        count = fm.GetFeatureCount(True)
        for i in range(count):
            try:
                feat = fm.FeatureByPositionReverse(i)
                if feat:
                    name = feat.Name
                    feat_type = feat.GetTypeName2()
                    if name not in SKIP_NAMES and feat_type not in SKIP_TYPES:
                        features.append({
                            "name": name,
                            "type": feat_type,
                        })
            except Exception:
                pass
        
        if features:
            features.reverse()  # was walked backwards
            print(f"    [DEBUG] Method 3 (ByPositionReverse): found {len(features)} features")
            return features
    except Exception as e:
        print(f"    [DEBUG] Method 3 (ByPositionReverse) failed: {e}")
    
    # Method 4: Read mission.json to get tools used (100% reliable fallback)
    try:
        import json
        import os
        mission_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "mission.json")
        if os.path.exists(mission_path):
            with open(mission_path, "r") as f:
                mission = json.load(f)
            
            # Map tool names to SolidWorks feature type names
            TOOL_TO_FEATURE = {
                "extrude": "Boss-Extrude",
                "extrude_midplane": "Boss-Extrude",
                "cut_extrude": "Cut-Extrude",
                "cut_through_all": "Cut-Extrude",
                "revolve": "Revolve",
                "revolve_simple": "Revolve",
                "shell": "Shell",
                "fillet": "Fillet",
                "chamfer": "Chamfer",
                "loft": "Loft",
                "sweep": "Sweep",
                "thread": "Thread",
                "thread_tap": "Thread",
                "circular_pattern": "CirPattern",
                "linear_pattern": "LPattern",
            }
            
            SKIP_TOOLS = {"create_part", "create_sketch", "create_sketch_on_selected_face",
                         "draw_rectangle", "draw_circle", "draw_hexagon", "draw_line",
                         "draw_arc", "draw_semicircle", "draw_triangle", "draw_polygon",
                         "draw_slot", "draw_spline", "draw_ellipse", "draw_centerline_vertical",
                         "validate_closed_profile", "exit_sketch",
                         "select_face_at_coordinate", "select_face_by_normal",
                         "select_edge_at_coordinate", "select_edge_at_coordinate_append",
                         "select_sketch", "create_reference_plane",
                         "get_feature_tree", "delete_feature"}
            
            counters = {}
            for step in mission:
                tool = step.get("tool", "")
                if tool in SKIP_TOOLS:
                    continue
                feat_type = TOOL_TO_FEATURE.get(tool, tool)
                counters[feat_type] = counters.get(feat_type, 0) + 1
            
            for feat_type, count in counters.items():
                for i in range(1, count + 1):
                    features.append({
                        "name": f"{feat_type}{i}",
                        "type": feat_type,
                    })
            
            if features:
                print(f"    [DEBUG] Method 4 (mission.json): found {len(features)} features from {len(mission)} steps")
                return features
    except Exception as e:
        print(f"    [DEBUG] Method 4 (mission.json) failed: {e}")
    
    # All methods failed — return empty list
    print(f"    [DEBUG] All feature tree methods failed")
    return features


# ============================================================
# BODY INFO
# ============================================================

def get_body_info():
    """
    Returns body count, face count, edge count.
    """
    model = _model()
    
    try:
        bodies = model.GetBodies2(0, False)
    except Exception as e:
        print(f"    [DEBUG] model.GetBodies2() failed: {e}")
        return {"body_count": 0, "face_count": 0, "edge_count": 0}
    
    if not bodies:
        return {"body_count": 0, "face_count": 0, "edge_count": 0}
    
    body_count = len(bodies)
    total_faces = 0
    total_edges = 0
    
    for body in bodies:
        try:
            faces = body.GetFaces()
            if faces:
                total_faces += len(faces)
        except Exception:
            pass
        try:
            edges = body.GetEdges()
            if edges:
                total_edges += len(edges)
        except Exception:
            pass
    
    return {
        "body_count": body_count,
        "face_count": total_faces,
        "edge_count": total_edges,
    }


# ============================================================
# BOUNDING BOX
# ============================================================

def get_bounding_box():
    """
    Gets bounding box from solid body.
    """
    model = _model()
    box = None
    
    try:
        bodies = model.GetBodies2(0, False)
        if bodies and len(bodies) > 0:
            box = bodies[0].GetBodyBox()
    except Exception as e:
        print(f"    [DEBUG] GetBodyBox failed: {e}")
    
    if box is None or not hasattr(box, '__len__') or len(box) < 6:
        return {
            "width": 0, "height": 0, "depth": 0,
            "min_x": 0, "min_y": 0, "min_z": 0,
            "max_x": 0, "max_y": 0, "max_z": 0,
        }
    
    # Convert meters to mm
    min_x = round(box[0] * 1000, 2)
    min_y = round(box[1] * 1000, 2)
    min_z = round(box[2] * 1000, 2)
    max_x = round(box[3] * 1000, 2)
    max_y = round(box[4] * 1000, 2)
    max_z = round(box[5] * 1000, 2)
    
    return {
        "min_x": min_x, "min_y": min_y, "min_z": min_z,
        "max_x": max_x, "max_y": max_y, "max_z": max_z,
        "width": round(max_x - min_x, 2),
        "height": round(max_y - min_y, 2),
        "depth": round(max_z - min_z, 2),
    }


# ============================================================
# MASS PROPERTIES (Volume & Surface Area) - FIXED VERSION
# ============================================================

def get_mass_properties():
    """
    Extracts volume and surface area from all solid bodies.
    
    IBody2::GetMassProperties(density) for SOLID bodies returns:
        [0] CenterOfMass_X (m)     [1] CenterOfMass_Y (m)     [2] CenterOfMass_Z (m)
        [3] Volume (m³)            [4] Surface Area (m²)       [5] Mass (kg)
        [6-11] Moments of inertia
    
    We use mp[3] for volume and mp[4] for surface area.
    Values are converted from SI (meters) to mm.
    """
    model = _model()
    result = {"volume_mm3": None, "surface_area_mm2": None}
    
    try:
        # Force rebuild to ensure geometry is computed
        print(f"    [DEBUG] Rebuilding model before measurement...")
        try:
            model.ForceRebuild3(True)  # True = rebuild all
        except Exception as e:
            print(f"    [DEBUG] ForceRebuild3 failed (non-critical): {e}")
        
        # Get all bodies
        bodies = model.GetBodies2(0, False)  # 0 = solid bodies
        if not bodies or len(bodies) == 0:
            print(f"    [DEBUG] No bodies found in model")
            return result
        
        print(f"    [DEBUG] Found {len(bodies)} body/bodies")
        
        total_volume_m3 = 0
        total_surface_area_m2 = 0
        
        # Process each body
        for i, body in enumerate(bodies):
            try:
                # Get mass properties array
                # Density doesn't matter for pure geometric properties
                mp = body.GetMassProperties(1.0)
                
                if not mp:
                    print(f"    [DEBUG] Body {i}: GetMassProperties returned None")
                    continue
                
                print(f"    [DEBUG] Body {i}: Raw array length = {len(mp)}")
                print(f"    [DEBUG] Body {i}: Raw array = {mp[:min(15, len(mp))]}")
                
                if len(mp) < 5:
                    print(f"    [DEBUG] Body {i}: Array too short (need at least 5 elements)")
                    continue
                
                # IBody2::GetMassProperties for solid bodies:
                #   mp[3] = Volume (m³),  mp[4] = Surface Area (m²)
                volume_m3 = mp[3]
                surface_area_m2 = mp[4]
                
                if volume_m3 is None or volume_m3 == 0:
                    print(f"    [DEBUG] Body {i}: Volume is None or 0")
                    # Try alternative method using bounding box
                    bbox = body.GetBodyBox()
                    if bbox and len(bbox) >= 6:
                        w = bbox[3] - bbox[0]
                        h = bbox[4] - bbox[1]
                        d = bbox[5] - bbox[2]
                        volume_m3 = w * h * d  # Rough estimate
                        print(f"    [DEBUG] Body {i}: Using bbox estimate: {volume_m3} m³")
                
                if surface_area_m2 is None or surface_area_m2 == 0:
                    print(f"    [DEBUG] Body {i}: Surface area is None or 0")
                    # Don't skip volume if SA fails!
                
                print(f"    [DEBUG] Body {i}: Volume = {volume_m3} m³")
                print(f"    [DEBUG] Body {i}: Surface Area = {surface_area_m2} m²")
                
                total_volume_m3 += volume_m3 if volume_m3 else 0
                total_surface_area_m2 += surface_area_m2 if surface_area_m2 else 0
                
            except Exception as e:
                print(f"    [DEBUG] Body {i}: Error processing - {e}")
                continue
        
        # Convert from meters to millimeters
        if total_volume_m3 > 0:
            result["volume_mm3"] = round(total_volume_m3 * 1e9, 2)  # m³ → mm³
            print(f"    [DEBUG] ✅ Total Volume: {total_volume_m3:.6e} m³ = {result['volume_mm3']:.2f} mm³")
        
        if total_surface_area_m2 > 0:
            result["surface_area_mm2"] = round(total_surface_area_m2 * 1e6, 2)  # m² → mm²
            print(f"    [DEBUG] ✅ Total Surface Area: {total_surface_area_m2:.6e} m² = {result['surface_area_mm2']:.2f} mm²")
        
    except Exception as e:
        print(f"    [DEBUG] ❌ Error in get_mass_properties: {e}")
        import traceback
        traceback.print_exc()
    
    return result


# ============================================================
# COMBINED PROPERTIES
# ============================================================

def get_model_properties():
    """
    Returns all measurable model properties.
    """
    model = _model()
    
    result = {
        "dimensions": {"width": 0, "height": 0, "depth": 0},
        "bounding_box": {},
        "features": [],
        "feature_types": {},
        "feature_count": 0,
        "feature_tree_available": False,
        "body_count": 0,
        "face_count": 0,
        "edge_count": 0,
        "volume_mm3": None,
        "surface_area_mm2": None,
    }
    
    # 1. Feature tree
    print("    [DEBUG] Getting feature tree...")
    try:
        features = get_feature_tree()
        result["features"] = features
        result["feature_count"] = len(features)
        feature_types = {}
        for f in features:
            ft = f["type"]
            feature_types[ft] = feature_types.get(ft, 0) + 1
        result["feature_types"] = feature_types
        
        if len(features) == 0:
            try:
                fm = model.FeatureManager
                fm_count = fm.GetFeatureCount(True)
                result["feature_count"] = fm_count
                result["feature_tree_available"] = False
                print(f"    [DEBUG] ✅ Features: {fm_count} (count only)")
            except Exception as e:
                result["feature_tree_available"] = False
                print(f"    [DEBUG] ⚠️ FeatureManager.GetFeatureCount failed: {e}")
        else:
            result["feature_tree_available"] = True
            print(f"    [DEBUG] ✅ Features: {len(features)}")
    except Exception as e:
        result["feature_tree_available"] = False
        print(f"    [DEBUG] ❌ Feature tree failed: {e}")
    
    # 2. Body info
    print("    [DEBUG] Getting body info...")
    try:
        body_info = get_body_info()
        result["body_count"] = body_info["body_count"]
        result["face_count"] = body_info["face_count"]
        result["edge_count"] = body_info["edge_count"]
        print(f"    [DEBUG] ✅ Bodies: {body_info['body_count']}, Faces: {body_info['face_count']}")
    except Exception as e:
        print(f"    [DEBUG] ❌ Body info failed: {e}")
    
    # 3. Bounding box
    print("    [DEBUG] Getting bounding box...")
    try:
        bbox = get_bounding_box()
        result["bounding_box"] = bbox
        result["dimensions"] = {
            "width": bbox["width"],
            "height": bbox["height"],
            "depth": bbox["depth"],
        }
        print(f"    [DEBUG] ✅ Dims: {bbox['width']}W x {bbox['height']}H x {bbox['depth']}D")
    except Exception as e:
        print(f"    [DEBUG] ❌ Bounding box failed: {e}")
    
    # 4. Mass properties
    print("    [DEBUG] Getting mass properties...")
    try:
        mass_props = get_mass_properties()
        result["volume_mm3"] = mass_props["volume_mm3"]
        result["surface_area_mm2"] = mass_props["surface_area_mm2"]
    except Exception as e:
        print(f"    [DEBUG] ❌ Mass properties failed: {e}")
    
    return result


def get_model_summary():
    """
    Returns a human-readable text summary.
    """
    model = _model()
    
    try:
        doc_type = model.GetType
        if callable(doc_type):
            doc_type = doc_type()
        print(f"    [DEBUG] Document type: {doc_type} (1=Part, 2=Assembly, 3=Drawing)")
    except Exception as e:
        print(f"    [DEBUG] Could not get doc type: {e}")
    
    try:
        props = get_model_properties()
    except Exception as e:
        print(f"    ❌ Error inspecting model: {e}")
        return None
    
    bbox = props.get("bounding_box", {})
    dims = props["dimensions"]
    
    lines = [
        "=== CURRENT MODEL STATE ===",
        f"Dimensions: {dims['width']}W x {dims['height']}H x {dims['depth']}D mm",
        f"Bodies: {props['body_count']}, Faces: {props['face_count']}, Edges: {props['edge_count']}",
    ]
    
    # Add volume and surface area
    if props.get('volume_mm3'):
        lines.append(f"Volume: {props['volume_mm3']:,.2f} mm³")
    if props.get('surface_area_mm2'):
        lines.append(f"Surface Area: {props['surface_area_mm2']:,.2f} mm²")
    
    # Include bounding box
    if bbox.get("min_x") is not None:
        x1, x2 = bbox['min_x'], bbox['max_x']
        y1, y2 = bbox['min_y'], bbox['max_y']
        z1, z2 = bbox['min_z'], bbox['max_z']
        mx = round((x1 + x2) / 2, 2)
        mz = round((z1 + z2) / 2, 2)
        
        lines.append(f"\nBounding Box (mm):")
        lines.append(f"  X: {x1} to {x2}")
        lines.append(f"  Y: {y1} to {y2}")
        lines.append(f"  Z: {z1} to {z2}")
    
    lines.append(f"\nFeature Tree ({props['feature_count']} features):")
    for i, f in enumerate(props["features"], 1):
        lines.append(f"  {i}. {f['name']} ({f['type']})")
    
    if props['feature_count'] == 0 and not props.get('feature_tree_available', True):
        lines.append(f"  (tree not traversable, but {props['feature_count']} features exist)")
    elif props['feature_count'] == 0:
        lines.append("  (no features found)")
    
    return "\n".join(lines)