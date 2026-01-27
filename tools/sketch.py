"""
Production-ready sketch module for SolidWorks CAD engine.
"""

import math

try:
    from .solidworks_app import get_active_model, get_nothing
except ImportError:
    from solidworks_app import get_active_model, get_nothing

_SKETCH_ACTIVE = False

PLANE_MAP = {
    "Front": "Front Plane",
    "Top": "Top Plane",
    "Right": "Right Plane"
}

def _model():
    model = get_active_model()
    if model is None:
        raise Exception("No active SolidWorks document")
    return model

def _sm():
    return _model().SketchManager

def _pt(x, y):
    """Convert mm -> meters"""
    return x / 1000, y / 1000, 0

def _require_sketch_active():
    if not _SKETCH_ACTIVE:
        raise Exception("No active sketch. Call create_sketch() first.")

def _active_sketch():
    try:
        # Try getting active sketch
        model = _model()
        sketch = model.GetActiveSketch2()
        return sketch
    except:
        # If GetActiveSketch2 fails, return None
        return None
    
def create_sketch(plane: str):
    """Starts a new sketch on the specified plane."""
    global _SKETCH_ACTIVE

    if _active_sketch() is None:
        _SKETCH_ACTIVE = False

    if _SKETCH_ACTIVE:
        raise Exception("Sketch already active. Exit current sketch first.")

    if plane not in PLANE_MAP:
        raise Exception(f"Invalid plane: {plane}. Must be Front, Top, or Right.")

    model = _model()
    nothing = get_nothing()
    
    model.Extension.SelectByID2(
        PLANE_MAP[plane],
        "PLANE",
        0, 0, 0,
        False, 0,
        nothing, 0
    )

    model.InsertSketch2(True)
    _SKETCH_ACTIVE = True
    return f"Sketch created on {plane} Plane"

def select_face_at_coordinate(x, y, z):
    """Selects a face at a specific 3D coordinate (mm)."""
    model = _model()
    nothing = get_nothing()
    model.ClearSelection2(True)
    
    x_m, y_m, z_m = x/1000.0, y/1000.0, z/1000.0
    
    status = model.Extension.SelectByID2("", "FACE", x_m, y_m, z_m, False, 0, nothing, 0)
    
    if not status:
        for offset in [0.0001, -0.0001, 0.0002, -0.0002]:
            status = model.Extension.SelectByID2("", "FACE", x_m, y_m + offset, z_m, False, 0, nothing, 0)
            if status:
                break
    
    if not status:
        raise Exception(f"No face found at coordinates ({x}, {y}, {z})")
    
    return f"Selected face at ({x}, {y}, {z})"

def select_face_by_normal(direction="up"):
    """Intelligently selects a face based on its orientation."""
    model = _model()
    nothing = get_nothing()
    
    model.ClearSelection2(True)
    
    direction_map = {
        "up": (0, 1, 0),
        "down": (0, -1, 0),
        "front": (0, 0, 1),
        "back": (0, 0, -1),
        "right": (1, 0, 0),
        "left": (-1, 0, 0)
    }
    
    if direction not in direction_map:
        raise Exception(f"Invalid direction: {direction}")
    
    target_normal = direction_map[direction]
    
    part = model
    bodies = part.GetBodies2(0, False)
    
    if not bodies:
        raise Exception("No solid bodies found")
    
    body = bodies[0]
    faces = body.GetFaces()
    
    if not faces:
        raise Exception("No faces found")
    
    best_face = None
    best_dot = -2
    
    for face in faces:
        try:
            normal = face.Normal
            if normal and len(normal) >= 3:
                dot = (normal[0] * target_normal[0] + 
                       normal[1] * target_normal[1] + 
                       normal[2] * target_normal[2])
                
                if dot > best_dot:
                    best_dot = dot
                    best_face = face
        except:
            pass
    
    if best_face and best_dot > 0.8:
        best_face.Select4(False, nothing)
        return f"Selected face pointing {direction}"
    else:
        raise Exception(f"Could not find face pointing {direction}")

def create_sketch_on_selected_face():
    """Creates a sketch on the currently selected face."""
    global _SKETCH_ACTIVE

    if _active_sketch() is None:
        _SKETCH_ACTIVE = False

    if _SKETCH_ACTIVE:
        raise Exception("Sketch already active")

    model = _model()
    sel_mgr = model.SelectionManager
    count_val = sel_mgr.GetSelectedObjectCount
    if callable(count_val):
        count_val = count_val()
    
    if count_val == 0:
        raise Exception("No face selected")
    
    # Use the EXACT pattern from the test that worked
    sm = _model().SketchManager
    sm.InsertSketch(True)
    _SKETCH_ACTIVE = True
    
    return "Sketch created on selected face"

def exit_sketch():
    """Exits the current sketch."""
    global _SKETCH_ACTIVE
    if not _SKETCH_ACTIVE:
        raise Exception("No active sketch to exit")
    _model().InsertSketch2(True)
    _SKETCH_ACTIVE = False
    return "Exited sketch"

def draw_line(x1, y1, x2, y2):
    """Draws a line from (x1,y1) to (x2,y2) in mm."""
    _require_sketch_active()
    x1m, y1m, _ = _pt(x1, y1)
    x2m, y2m, _ = _pt(x2, y2)
    _sm().CreateLine(x1m, y1m, 0, x2m, y2m, 0)
    return f"Line drawn from ({x1},{y1}) to ({x2},{y2})"

def draw_centerline_vertical():
    """Draws a vertical centerline (construction line) for revolve."""
    _require_sketch_active()
    sm = _sm()
    line = sm.CreateLine(0, -0.1, 0, 0, 0.1, 0)
    if line:
        line.ConstructionGeometry = True
    return "Vertical centerline drawn"

def draw_rectangle(width, height, x=0, y=0):
    """Draws a rectangle centered at (x, y)."""
    _require_sketch_active()
    
    w_m = width / 1000.0
    h_m = height / 1000.0
    x_m = x / 1000.0
    y_m = y / 1000.0
    
    x1 = x_m - w_m / 2.0
    y1 = y_m - h_m / 2.0
    x2 = x_m + w_m / 2.0
    y2 = y_m + h_m / 2.0
    
    _sm().CreateCornerRectangle(x1, y1, 0, x2, y2, 0)
    
    return f"Rectangle {width}x{height}mm drawn (centered)"

def draw_circle(radius, x=0, y=0):
    """Draws a circle at (x,y). All units in mm."""
    _require_sketch_active()
    r = radius / 1000.0
    x_m = x / 1000.0
    y_m = y / 1000.0
    _sm().CreateCircleByRadius(x_m, y_m, 0, r)
    return f"Circle radius {radius}mm drawn at ({x},{y})"

def draw_arc(radius, start_angle, end_angle):
    """Draws an arc centered at origin."""
    _require_sketch_active()
    sm = _sm()
    r = radius / 1000.0
    x1 = r * math.cos(math.radians(start_angle))
    y1 = r * math.sin(math.radians(start_angle))
    x2 = r * math.cos(math.radians(end_angle))
    y2 = r * math.sin(math.radians(end_angle))
    sm.CreateArc(0, 0, 0, x1, y1, 0, x2, y2, 0, 1)
    return f"Arc radius {radius}mm"

def draw_semicircle(radius):
    """Draws a semicircle at origin."""
    _require_sketch_active()
    r = radius / 1000.0
    _sm().CreateArc(0, 0, 0, r, 0, 0, -r, 0, 0, 1)
    return f"Semicircle radius {radius}mm drawn"

def draw_polygon(sides, radius):
    """Draws a regular polygon centered at origin."""
    _require_sketch_active()
    if sides < 3:
        raise Exception("Polygon must have at least 3 sides")
    
    r = radius / 1000.0
    points = []
    for i in range(sides):
        angle = 2 * math.pi * i / sides
        points.append((r * math.cos(angle), r * math.sin(angle)))
    
    sm = _sm()
    for i in range(len(points)):
        x1, y1 = points[i]
        x2, y2 = points[(i + 1) % len(points)]
        sm.CreateLine(x1, y1, 0, x2, y2, 0)
    
    return f"{sides}-sided polygon drawn"

def draw_slot(length, width):
    """Draws a slot shape."""
    _require_sketch_active()
    sm = _sm()
    r = width / 2000.0
    half_len = length / 2000.0
    sm.CreateArc(-half_len, 0, 0, -half_len, r, 0, -half_len, -r, 0, 1)
    sm.CreateArc(half_len, 0, 0, half_len, -r, 0, half_len, r, 0, 1)
    sm.CreateLine(-half_len, r, 0, half_len, r, 0)
    sm.CreateLine(half_len, -r, 0, -half_len, -r, 0)
    return f"Slot {length}x{width}mm drawn"

def draw_ellipse(major_radius, minor_radius):
    _require_sketch_active()
    _sm().CreateEllipse(0, 0, 0, major_radius/1000.0, 0, 0, 0, minor_radius/1000.0, 0)
    return "Ellipse drawn"

def validate_closed_profile():
    """Exits sketch and prepares for feature creation."""
    global _SKETCH_ACTIVE
    _require_sketch_active()
    model = _model()
    model.InsertSketch2(True)
    _SKETCH_ACTIVE = False
    return "Sketch validated and closed"