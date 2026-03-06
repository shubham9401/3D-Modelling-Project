"""
Model Inspector: Extracts measurable properties from the active SolidWorks model.

Uses ONLY COM methods that are proven to work in the existing codebase.
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
    Uses same traversal as circular_pattern() in feature.py.
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
    
    try:
        feat = model.FirstFeature()
    except Exception as e:
        print(f"    [DEBUG] model.FirstFeature() not available: {e}")
        # Fallback: use FeatureManager.GetFeatureCount to at least get a count
        try:
            fm = model.FeatureManager
            count = fm.GetFeatureCount(True)
            print(f"    [DEBUG] FeatureManager reports {count} features (tree not traversable)")
        except Exception:
            pass
        return features
    
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
    
    return features


# ============================================================
# BODY INFO
# ============================================================

def get_body_info():
    """
    Returns body count, face count, edge count.
    Uses GetBodies2 (same as sketch.py line 245).
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
    Gets bounding box from solid body (body.GetBodyBox).
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
# COMBINED PROPERTIES
# ============================================================

def get_model_properties():
    """
    Returns all measurable model properties.
    Each section is wrapped in try/except to prevent one failure
    from killing the entire inspection.
    """
    model = _model()  # Get the model object once for this function
    
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
        
        # If tree traversal failed, get count from FeatureManager
        if len(features) == 0:
            try:
                fm = model.FeatureManager
                fm_count = fm.GetFeatureCount(True)
                result["feature_count"] = fm_count
                result["feature_tree_available"] = False
                print(f"    [DEBUG] ✅ Features: {fm_count} (count only, tree not available)")
            except Exception as e:
                result["feature_tree_available"] = False
                print(f"    [DEBUG] ⚠️ FeatureManager.GetFeatureCount also failed: {e}")
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
    
    return result


def get_model_summary():
    """
    Returns a human-readable text summary. Returns None on complete failure.
    Includes bounding box coordinates so the LLM can generate correct edge positions.
    """
    model = _model()
    
    # Check what type of document this is
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
    
    # Include bounding box so LLM knows exact coordinates
    if bbox.get("min_x") is not None:
        x1, x2 = bbox['min_x'], bbox['max_x']
        y1, y2 = bbox['min_y'], bbox['max_y']
        z1, z2 = bbox['min_z'], bbox['max_z']
        mx = round((x1 + x2) / 2, 2)  # midpoints
        mz = round((z1 + z2) / 2, 2)
        
        lines.append(f"\nBounding Box (mm):")
        lines.append(f"  X: {x1} to {x2}")
        lines.append(f"  Y: {y1} to {y2}")
        lines.append(f"  Z: {z1} to {z2}")
        
        # Edge MIDPOINTS — select_edge_at_coordinate works best at edge midpoints, NOT corners!
        lines.append(f"\nEdge Midpoints (use these for select_edge_at_coordinate):")
        lines.append(f"  IMPORTANT: Always use edge MIDPOINTS, never corners!")
        lines.append(f"  Top 4 edges (Y={y2}):")
        lines.append(f"    Front:  ({mx}, {y2}, {z1})")
        lines.append(f"    Back:   ({mx}, {y2}, {z2})")
        lines.append(f"    Left:   ({x1}, {y2}, {mz})")
        lines.append(f"    Right:  ({x2}, {y2}, {mz})")
        lines.append(f"  Bottom 4 edges (Y={y1}):")
        lines.append(f"    Front:  ({mx}, {y1}, {z1})")
        lines.append(f"    Back:   ({mx}, {y1}, {z2})")
        lines.append(f"    Left:   ({x1}, {y1}, {mz})")
        lines.append(f"    Right:  ({x2}, {y1}, {mz})")
        lines.append(f"  Vertical 4 edges:")
        lines.append(f"    Front-Left:  ({x1}, {round((y1+y2)/2,2)}, {z1})")
        lines.append(f"    Front-Right: ({x2}, {round((y1+y2)/2,2)}, {z1})")
        lines.append(f"    Back-Left:   ({x1}, {round((y1+y2)/2,2)}, {z2})")
        lines.append(f"    Back-Right:  ({x2}, {round((y1+y2)/2,2)}, {z2})")
        lines.append(f"  Top face center: ({mx}, {y2}, {mz})")
    
    lines.append(f"\nFeature Tree ({props['feature_count']} features):")
    for i, f in enumerate(props["features"], 1):
        lines.append(f"  {i}. {f['name']} ({f['type']})")
    
    if props['feature_count'] == 0 and not props.get('feature_tree_available', True):
        lines.append(f"  (tree not traversable, but {props['feature_count']} features exist)")
    elif props['feature_count'] == 0:
        lines.append("  (no features found)")
    
    return "\n".join(lines)
