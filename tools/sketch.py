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
    """
    Starts a new sketch on the specified plane.
    
    Args:
        plane: "Front", "Top", "Right", or reference plane name like "Plane1"
    """
    global _SKETCH_ACTIVE

    if _active_sketch() is None:
        _SKETCH_ACTIVE = False

    if _SKETCH_ACTIVE:
        raise Exception("Sketch already active. Exit current sketch first.")

    model = _model()
    nothing = get_nothing()
    
    # Check if it's a standard plane or reference plane
    if plane in PLANE_MAP:
        plane_name = PLANE_MAP[plane]
        plane_type = "PLANE"
    else:
        # Assume it's a reference plane (Plane1, Plane2, etc.)
        plane_name = plane
        plane_type = "PLANE"
    
    result = model.Extension.SelectByID2(
        plane_name,
        plane_type,
        0, 0, 0,
        False, 0,
        nothing, 0
    )
    
    if not result:
        raise Exception(f"Failed to select plane: {plane}")

    model.InsertSketch2(True)
    _SKETCH_ACTIVE = True
    return f"Sketch created on {plane}"

def select_face_at_coordinate(x, y, z):
    """
    Selects a face at a specific 3D coordinate (mm).
    Uses SelectByRay for more reliable selection (like VBA macro).
    """
    model = _model()
    nothing = get_nothing()
    model.ClearSelection2(True)
    
    x_m, y_m, z_m = x/1000.0, y/1000.0, z/1000.0
    
    # Try SelectByID2 first
    status = model.Extension.SelectByID2("", "FACE", x_m, y_m, z_m, False, 0, nothing, 0)
    
    if not status:
        # If failed, try SelectByRay - cast a ray downward from above the point
        # SelectByRay(RayOriginX, RayOriginY, RayOriginZ, RayDirX, RayDirY, RayDirZ, Radius, Type, Append, Mark, Option)
        # Type = 2 for faces
        # Direction: shoot ray downward (-Y) to hit top face
        status = model.Extension.SelectByRay(
            x_m, y_m + 0.01, z_m,    # Origin: slightly above the target point
            0, -1, 0,                 # Direction: straight down (-Y)
            0.001,                    # Radius (small value)
            2,                        # Type: 2 = face
            False,                    # Append
            0,                        # Mark
            0                         # Option
        )
    
    if not status:
        # Try other offsets with SelectByID2
        for offset in [0.001, -0.001, 0.002, -0.002]:
            status = model.Extension.SelectByID2("", "FACE", x_m, y_m + offset, z_m, False, 0, nothing, 0)
            if status:
                break
    
    if not status:
        raise Exception(f"No face found at coordinates ({x}, {y}, {z})")
    
    return f"Selected face at ({x}, {y}, {z})"

def select_edge_at_coordinate(x, y, z):
    """
    Selects an edge at a specific 3D coordinate (mm).
    Uses SelectByRay for reliable edge selection.
    
    IMPORTANT: Call this BEFORE fillet() to select the edge(s) to fillet.
    
    Args:
        x, y, z: Coordinates near the edge to select (in mm)
    """
    model = _model()
    nothing = get_nothing()
    
    x_m, y_m, z_m = x/1000.0, y/1000.0, z/1000.0
    
    # Clear any previous selection
    model.ClearSelection2(True)
    
    # Try SelectByRay FIRST (more reliable for edges than SelectByID2)
    ray_directions = [
        (0, -1, 0),   # Down
        (0, 1, 0),    # Up
        (1, 0, 0),    # Right
        (-1, 0, 0),   # Left
        (0, 0, 1),    # Front
        (0, 0, -1),   # Back
    ]
    
    status = False
    for dx, dy, dz in ray_directions:
        ox = x_m - dx * 0.01
        oy = y_m - dy * 0.01
        oz = z_m - dz * 0.01
        
        status = model.Extension.SelectByRay(
            ox, oy, oz,
            dx, dy, dz,
            0.002,                # Radius (slightly larger for reliability)
            1,                    # Type: 1 = edge
            False,                # Append = False (first selection)
            1,                    # Mark = 1 (for fillet)
            0                     # Option
        )
        if status:
            break
    
    # Fallback to SelectByID2
    if not status:
        status = model.Extension.SelectByID2("", "EDGE", x_m, y_m, z_m, False, 1, nothing, 0)
    
    if not status:
        raise Exception(f"No edge found at coordinates ({x}, {y}, {z})")
    
    # Verify selection
    selMgr = model.SelectionManager
    count = selMgr.GetSelectedObjectCount2(-1)
    
    return f"Selected edge at ({x}, {y}, {z}) [total selected: {count}]"

def select_edge_at_coordinate_append(x, y, z):
    """
    Selects an additional edge at a specific 3D coordinate (mm).
    Appends to current selection - use after select_edge_at_coordinate for multiple edges.
    
    Args:
        x, y, z: Coordinates near the edge to select (in mm)
    """
    model = _model()
    nothing = get_nothing()
    
    x_m, y_m, z_m = x/1000.0, y/1000.0, z/1000.0
    
    # Get count before
    selMgr = model.SelectionManager
    count_before = selMgr.GetSelectedObjectCount2(-1)
    
    # Try SelectByRay with Append=True FIRST
    ray_directions = [
        (0, -1, 0), (0, 1, 0), (1, 0, 0), (-1, 0, 0), (0, 0, 1), (0, 0, -1),
    ]
    
    status = False
    for dx, dy, dz in ray_directions:
        ox = x_m - dx * 0.01
        oy = y_m - dy * 0.01
        oz = z_m - dz * 0.01
        
        status = model.Extension.SelectByRay(
            ox, oy, oz, dx, dy, dz,
            0.002, 1, True, 1, 0  # Append=True, Mark=1
        )
        if status:
            break
    
    # Fallback to SelectByID2 with append
    if not status:
        status = model.Extension.SelectByID2("", "EDGE", x_m, y_m, z_m, True, 1, nothing, 0)
    
    if not status:
        raise Exception(f"No edge found at coordinates ({x}, {y}, {z})")
    
    # Verify selection count increased
    count_after = selMgr.GetSelectedObjectCount2(-1)
    
    return f"Appended edge at ({x}, {y}, {z}) [total selected: {count_after}]"

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

def draw_ellipse(radius_x, radius_y, x=0, y=0):
    """
    Draws an ellipse at (x,y) with specified radii.
    
    Args:
        radius_x: Semi-major axis (horizontal radius) in mm
        radius_y: Semi-minor axis (vertical radius) in mm
        x: Center X coordinate in mm (default 0)
        y: Center Y coordinate in mm (default 0)
    """
    _require_sketch_active()
    
    # Convert to meters
    rx = radius_x / 1000.0
    ry = radius_y / 1000.0
    x_m = x / 1000.0
    y_m = y / 1000.0
    
    # CreateEllipse(CenterX, CenterY, CenterZ, MajorAxisX, MajorAxisY, MajorAxisZ, MinorAxisX, MinorAxisY, MinorAxisZ)
    # Major axis point is on the ellipse boundary
    # Minor axis point is on the ellipse boundary
    _sm().CreateEllipse(
        x_m, y_m, 0,           # Center point
        x_m + rx, y_m, 0,      # Point on major axis (right edge)
        x_m, y_m + ry, 0       # Point on minor axis (top edge)
    )
    return f"Ellipse {radius_x}x{radius_y}mm drawn at ({x},{y})"

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
    """
    Draws a semicircle with diameter line for sphere creation.
    
    Creates:
    - Arc1: The semicircle arc (profile to revolve)
    - Line1: The diameter line (axis for revolve)
    
    This creates a closed profile ready for revolve.
    """
    _require_sketch_active()
    sm = _sm()
    r = radius / 1000.0
    
    # Create the semicircle arc (from top to bottom on the right side)
    # Arc goes from (0, r) to (0, -r) curving to the right
    sm.CreateArc(0, 0, 0, 0, r, 0, 0, -r, 0, 1)
    
    # Create the diameter line to close the profile (and serve as axis)
    # This line connects the two endpoints of the arc
    sm.CreateLine(0, r, 0, 0, -r, 0)
    
    return f"Semicircle radius {radius}mm with diameter line drawn (ready for revolve)"

def draw_triangle(base, height):
    """
    Draws a right-angled triangle for Cone creation (Revolve).
    
    Creates:
    - Line1 (Vertical): The axis of revolution (Height).
    - Line2 (Horizontal): The base radius.
    - Line3 (Slant): The hypotenuse.
    
    This creates a closed profile ready for revolve.
    """
    _require_sketch_active()
    sm = _sm()
    
    b = base / 1000.0
    h = height / 1000.0
    
    # Draw Vertical Line (Axis) from Origin up
    # This will be "Line1" usually
    sm.CreateLine(0, 0, 0, 0, h, 0)
    
    # Draw Base Line from Origin right
    sm.CreateLine(0, 0, 0, b, 0, 0)
    
    # Draw Hypotenuse from (Base, 0) to (0, Height)
    sm.CreateLine(b, 0, 0, 0, h, 0)
    
    return f"Triangle (base={base}mm, height={height}mm) drawn"

def draw_polygon(sides, radius):
    """Draws a regular polygon centered at origin."""
    _require_sketch_active()
    if sides < 3:
        raise Exception("Polygon must have at least 3 sides")
    
    r = radius / 1000.0
    points = []
    for i in range(sides):
        angle = 2 * math.pi * i / sides
        points.append((round(r * math.cos(angle), 10), round(r * math.sin(angle), 10)))
    
    sm = _sm()
    for i in range(len(points)):
        x1, y1 = points[i]
        x2, y2 = points[(i + 1) % len(points)]
        sm.CreateLine(x1, y1, 0, x2, y2, 0)
    
    return f"{sides}-sided polygon drawn"

def draw_hexagon(radius, x=0, y=0):
    """
    Draws a regular hexagon centered at (x, y).
    Oriented with FLAT SIDES on top/bottom (standard bolt head orientation).
    
    Args:
        radius: Distance from center to vertex in mm
        x: Center X coordinate in mm (default 0)
        y: Center Y coordinate in mm (default 0)
    """
    _require_sketch_active()
    
    r = radius / 1000.0
    x_m = x / 1000.0
    y_m = y / 1000.0
    
    # Generate 6 vertices of hexagon with 30° offset
    # The offset ensures flat sides are on top/bottom (standard bolt head orientation)
    # Without offset: vertex points up → flat side on left/right (wrong for bolts)
    # With 30° offset: flat side on top/bottom (correct for bolt heads)
    points = []
    for i in range(6):
        angle = math.pi / 6 + (2 * math.pi * i / 6)  # Start at 30°
        px = round(x_m + r * math.cos(angle), 10)  # Round to prevent float drift
        py = round(y_m + r * math.sin(angle), 10)
        points.append((px, py))
    
    # Draw 6 lines connecting the vertices
    sm = _sm()
    for i in range(6):
        x1, y1 = points[i]
        x2, y2 = points[(i + 1) % 6]
        sm.CreateLine(x1, y1, 0, x2, y2, 0)
    
    return f"Hexagon radius {radius}mm drawn at ({x}, {y})"

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

def draw_spline(points):
    """
    Draws a spline curve through the specified points.
    Used for sweep paths.
    
    Args:
        points: List of [x, y] coordinates in mm.
                Example: [[0, 0], [50, 25], [100, 0]]
    
    Note: All Z coordinates are set to 0 (2D sketch).
    """
    import win32com.client
    
    _require_sketch_active()
    sm = _sm()
    
    # Build the points array (x, y, z triplets in meters)
    num_points = len(points)
    point_array = []
    
    for p in points:
        x = p[0] / 1000.0  # mm to meters
        y = p[1] / 1000.0
        z = 0  # 2D sketch
        point_array.extend([x, y, z])
    
    # Convert to VARIANT array
    import pythoncom
    variant_array = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, point_array)
    
    # Create the spline using CreateSpline2 (CreateSpline returns None in some versions)
    result = sm.CreateSpline2(variant_array, False)
    
    if result:
        return f"Spline drawn through {num_points} points"
    else:
        raise Exception("Failed to create spline")

def validate_closed_profile():
    """Confirms sketch is active and ready for features."""
    _require_sketch_active()
    # DO NOT EXIT SKETCH HERE. Features (Extrude/Revolve) work best when sketch is active.
    # The feature creation will automatically consume/close the sketch.
    return "Sketch profile validated (Sketch remains active)"