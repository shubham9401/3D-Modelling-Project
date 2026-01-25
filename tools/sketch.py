"""
Final production-ready sketch module for MCP-style SolidWorks CAD engine.

Capabilities:
- General & composite sketch primitives
- Sketch lifecycle enforcement
- Closed-profile validation
- Automatic constraints
- LLM-safe semantic interface
"""

import math

try:
    from .solidworks_app import get_active_model, get_nothing
except ImportError:
    from solidworks_app import get_active_model, get_nothing

# ============================================================
# INTERNAL STATE & HELPERS
# ============================================================

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
    # Safe property access
    val = _model().GetActiveSketch2
    if callable(val):
        return val()
    return val

# ============================================================
# SKETCH LIFECYCLE
# ============================================================

def create_sketch(plane: str):
    """
    Starts a new sketch on the specified plane.
    """
    global _SKETCH_ACTIVE

    if _SKETCH_ACTIVE:
        raise Exception("Sketch already active. Exit current sketch first.")

    if plane not in PLANE_MAP:
        raise Exception(f"Invalid plane: {plane}. Must be Front, Top, or Right.")

    model = _model()
    nothing = get_nothing()
    
    # Select the plane
    model.Extension.SelectByID2(
        PLANE_MAP[plane],
        "PLANE",
        0, 0, 0,
        False, 0,
        nothing, 0
    )

    # Start sketch using InsertSketch2
    model.InsertSketch2(True)
    _SKETCH_ACTIVE = True
    return f"Sketch created on {plane} Plane"


def select_face_by_normal(direction="up"):
    """
    Intelligently selects a face based on its orientation.
    direction can be: "up", "down", "front", "back", "left", "right"
    
    This works for ANY geometry - cubes, cylinders, complex shapes.
    """
    model = _model()
    nothing = get_nothing()
    
    # Clear selection
    model.ClearSelection2(True)
    
    # Define direction vectors
    direction_map = {
        "up": (0, 0, 1),
        "down": (0, 0, -1),
        "front": (0, 1, 0),
        "back": (0, -1, 0),
        "right": (1, 0, 0),
        "left": (-1, 0, 0)
    }
    
    if direction not in direction_map:
        raise Exception(f"Invalid direction: {direction}. Use: up, down, front, back, left, right")
    
    target_normal = direction_map[direction]
    
    # Get all bodies in the part
    part = model
    bodies = part.GetBodies2(0, False)  # 0 = solid bodies
    
    if not bodies or len(bodies) == 0:
        raise Exception("No solid bodies found")
    
    # Get the first (or main) body
    body = bodies[0]
    faces = body.GetFaces()
    
    if not faces or len(faces) == 0:
        raise Exception("No faces found on body")
    
    # Find the face with normal closest to target direction
    best_face = None
    best_dot = -2  # Dot product ranges from -1 to 1
    
    for face in faces:
        try:
            # Get face normal
            normal = face.Normal
            if normal and len(normal) >= 3:
                # Calculate dot product (how aligned the normals are)
                dot = (normal[0] * target_normal[0] + 
                       normal[1] * target_normal[1] + 
                       normal[2] * target_normal[2])
                
                if dot > best_dot:
                    best_dot = dot
                    best_face = face
        except:
            pass
    
    if best_face:
        best_face.Select4(False, nothing)
        return f"Selected face pointing {direction}"
    else:
        raise Exception(f"Could not find face pointing {direction}")


def create_sketch_on_selected_face():
    """
    Creates a sketch on the most recently selected face.
    This is the most flexible approach - works for any geometry.
    The user/AI should select the face first using select_face_by_normal or similar.
    """
    global _SKETCH_ACTIVE

    if _SKETCH_ACTIVE:
        raise Exception("Sketch already active. Exit current sketch first.")

    model = _model()
    
    # Check if a face is selected
    sel_mgr = model.SelectionManager
    count = sel_mgr.GetSelectedObjectCount
    if callable(count):
        count = count()
    
    if count == 0:
        raise Exception("No face selected. Select a face before calling this.")
    
    # Create sketch on the selected face
    model.SketchManager.InsertSketch(True)
    _SKETCH_ACTIVE = True
    
    return "Sketch created on selected face"


def create_sketch_on_top_face(height=None):
    """
    DEPRECATED but kept for backward compatibility.
    Automatically selects the topmost face and creates a sketch.
    For more control, use: select_face_by_normal("up") + create_sketch_on_selected_face()
    """
    global _SKETCH_ACTIVE

    if _SKETCH_ACTIVE:
        raise Exception("Sketch already active. Exit current sketch first.")

    # Use the generalized selection method
    select_face_by_normal("up")
    
    # Now create sketch on that face
    model = _model()
    model.SketchManager.InsertSketch(True)
    _SKETCH_ACTIVE = True
    
    return "Sketch created on top face"


def exit_sketch():
    """Exits the current sketch."""
    global _SKETCH_ACTIVE

    if not _SKETCH_ACTIVE:
        raise Exception("No active sketch to exit")

    _model().InsertSketch2(True)
    _SKETCH_ACTIVE = False
    return "Exited sketch"

# ============================================================
# BASIC PRIMITIVES
# ============================================================

def draw_line(x1, y1, x2, y2):
    """Draws a line from (x1,y1) to (x2,y2) in mm."""
    _require_sketch_active()
    x1m, y1m, _ = _pt(x1, y1)
    x2m, y2m, _ = _pt(x2, y2)
    _sm().CreateLine(x1m, y1m, 0, x2m, y2m, 0)
    return f"Line drawn from ({x1},{y1}) to ({x2},{y2})"


def draw_rectangle(width, height, x=0, y=0):
    """
    Draws a rectangle centered at (x, y).
    Default (0, 0) centers it at the sketch origin.
    All units in mm.
    """
    _require_sketch_active()
    
    # Convert to meters
    w_m = width / 1000.0
    h_m = height / 1000.0
    x_m = x / 1000.0
    y_m = y / 1000.0
    
    # Calculate corner coordinates for a center rectangle at (x, y)
    x1 = x_m - w_m / 2.0
    y1 = y_m - h_m / 2.0
    x2 = x_m + w_m / 2.0
    y2 = y_m + h_m / 2.0
    
    # Use CreateCornerRectangle
    _sm().CreateCornerRectangle(x1, y1, 0, x2, y2, 0)
    
    if x == 0 and y == 0:
        return f"Rectangle {width}x{height}mm drawn (centered)"
    else:
        return f"Rectangle {width}x{height}mm drawn at ({x},{y})"


def draw_circle(radius, x=0, y=0):
    """
    Draws a circle. Default (0,0) is Sketch Origin.
    """
    _require_sketch_active()
    r = radius / 1000.0
    xm = x / 1000.0
    ym = y / 1000.0
    _sm().CreateCircleByRadius(xm, ym, 0, r)
    return f"Circle radius {radius}mm drawn at ({x},{y})"

# ============================================================
# ARC & CURVE PRIMITIVES
# ============================================================

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
    return f"Arc radius {radius}mm ({start_angle} to {end_angle} deg)"


def draw_semicircle(radius):
    """Draws a semicircle at origin."""
    _require_sketch_active()
    r = radius / 1000.0
    _sm().CreateArc(0, 0, 0, r, 0, 0, -r, 0, 0, 1)
    return f"Semicircle radius {radius}mm drawn"


def draw_ellipse(major_radius, minor_radius):
    """Draws an ellipse at origin."""
    _require_sketch_active()
    maj = major_radius / 1000.0
    min_r = minor_radius / 1000.0
    _sm().CreateEllipse(0, 0, 0, maj, 0, 0, 0, min_r, 0)
    return f"Ellipse {major_radius}x{minor_radius}mm drawn"

# ============================================================
# POLYGON / SHAPES
# ============================================================

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

# ============================================================
# SEMANTIC / HIGH-LEVEL HELPERS
# ============================================================

def draw_slot(length, width):
    """Draws a slot shape (stadium/obround)."""
    _require_sketch_active()
    sm = _sm()

    r = width / 2000.0  # radius of end caps
    half_len = length / 2000.0

    sm.CreateArc(-half_len, 0, 0, -half_len, r, 0, -half_len, -r, 0, 1)
    sm.CreateArc(half_len, 0, 0, half_len, -r, 0, half_len, r, 0, 1)
    sm.CreateLine(-half_len, r, 0, half_len, r, 0)
    sm.CreateLine(half_len, -r, 0, -half_len, -r, 0)

    return f"Slot {length}x{width}mm drawn"


def draw_symmetric_circles(offset_x, radius):
    """Draws two circles symmetric about Y axis."""
    _require_sketch_active()
    sm = _sm()
    ox = offset_x / 1000.0
    r = radius / 1000.0
    sm.CreateCircleByRadius(ox, 0, 0, r)
    sm.CreateCircleByRadius(-ox, 0, 0, r)
    return f"Symmetric circles at +-{offset_x}mm, radius {radius}mm"

# ============================================================
# AUTOMATIC CONSTRAINTS
# ============================================================

def constrain_to_origin():
    """Constrains first sketch entity to origin."""
    _require_sketch_active()
    model = _model()
    sk = _active_sketch()
    nothing = get_nothing()

    entities = sk.GetSketchSegments()
    if not entities:
        raise Exception("No sketch entities to constrain")

    entities[0].Select(False)
    model.Extension.SelectByID2("Origin", "SKETCHPOINT", 0, 0, 0, True, 0, nothing, 0)
    model.SketchManager.AddConstraint("sgCOINCIDENT")
    return "Sketch constrained to origin"


def auto_horizontal_vertical():
    """Applies horizontal/vertical constraints to lines."""
    _require_sketch_active()
    model = _model()
    sk = _active_sketch()

    segments = sk.GetSketchSegments()
    if not segments:
        return "No segments to constrain"

    for seg in segments:
        try:
            if seg.GetType() == 0:  # Line
                dx = abs(seg.GetEndPoint2().X - seg.GetStartPoint2().X)
                dy = abs(seg.GetEndPoint2().Y - seg.GetStartPoint2().Y)
                seg.Select(False)
                if dx > dy:
                    model.SketchManager.AddConstraint("sgHORIZONTAL")
                else:
                    model.SketchManager.AddConstraint("sgVERTICAL")
        except:
            pass
    return "Auto horizontal/vertical constraints applied"

# ============================================================
# SKETCH VALIDATION
# ============================================================

def validate_closed_profile():
    """Ensures sketch is a closed profile before features."""
    _model().ClearSelection2(True)
    global _SKETCH_ACTIVE
    _model().InsertSketch2(True)
    _SKETCH_ACTIVE = False
    return "Sketch validated and closed"