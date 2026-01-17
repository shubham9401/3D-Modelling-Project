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
from solidworks_app import get_active_model

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
    """Convert mm → meters"""
    return x / 1000, y / 1000, 0

def _require_sketch_active():
    if not _SKETCH_ACTIVE:
        raise Exception("No active sketch. Call create_sketch() first.")

def _active_sketch():
    return _model().GetActiveSketch2()

# ============================================================
# SKETCH LIFECYCLE
# ============================================================

def create_sketch(plane: str):
    global _SKETCH_ACTIVE

    if _SKETCH_ACTIVE:
        raise Exception("Sketch already active")

    if plane not in PLANE_MAP:
        raise Exception(f"Invalid plane: {plane}")

    model = _model()
    model.Extension.SelectByID2(
        PLANE_MAP[plane],
        "PLANE",
        0, 0, 0,
        False, 0, None, 0
    )

    model.SketchManager.InsertSketch(True)
    _SKETCH_ACTIVE = True
    return f"Sketch created on {plane}"

def exit_sketch():
    global _SKETCH_ACTIVE

    if not _SKETCH_ACTIVE:
        raise Exception("No active sketch to exit")

    _model().SketchManager.InsertSketch(True)
    _SKETCH_ACTIVE = False
    return "Exited sketch"

# ============================================================
# BASIC PRIMITIVES
# ============================================================

def draw_line(x1, y1, x2, y2):
    _require_sketch_active()
    _sm().CreateLine(*_pt(x1, y1), *_pt(x2, y2))
    return "Line drawn"

def draw_rectangle(width, height):
    _require_sketch_active()
    _sm().CreateCenterRectangle(
        0, 0, 0,
        width / 2000,
        height / 2000,
        0
    )
    return f"Rectangle {width}x{height} drawn"

def draw_circle(radius):
    _require_sketch_active()
    _sm().CreateCircleByRadius(0, 0, 0, radius / 1000)
    return f"Circle radius {radius} drawn"

# ============================================================
# ARC & CURVE PRIMITIVES
# ============================================================

def draw_arc(radius, start_angle, end_angle):
    _require_sketch_active()
    sm = _sm()

    x1 = radius * math.cos(math.radians(start_angle))
    y1 = radius * math.sin(math.radians(start_angle))
    x2 = radius * math.cos(math.radians(end_angle))
    y2 = radius * math.sin(math.radians(end_angle))

    sm.CreateArc(
        0, 0, 0,
        x1 / 1000, y1 / 1000, 0,
        x2 / 1000, y2 / 1000, 0,
        1
    )
    return f"Arc radius {radius} ({start_angle}° → {end_angle}°)"

def draw_semicircle(radius):
    _require_sketch_active()
    _sm().CreateArc(
        0, 0, 0,
        radius / 1000, 0, 0,
        -radius / 1000, 0, 0,
        1
    )
    return f"Semicircle radius {radius} drawn"

def draw_ellipse(major_radius, minor_radius):
    _require_sketch_active()
    _sm().CreateEllipse(
        0, 0, 0,
        major_radius / 1000, 0, 0,
        0, minor_radius / 1000, 0
    )
    return f"Ellipse {major_radius}x{minor_radius} drawn"

# ============================================================
# POLYGON / SHAPES
# ============================================================

def draw_polygon(sides, radius):
    _require_sketch_active()

    if sides < 3:
        raise Exception("Polygon must have at least 3 sides")

    points = []
    for i in range(sides):
        angle = 2 * math.pi * i / sides
        points.append((
            radius * math.cos(angle),
            radius * math.sin(angle)
        ))

    for i in range(len(points)):
        x1, y1 = points[i]
        x2, y2 = points[(i + 1) % len(points)]
        draw_line(x1, y1, x2, y2)

    return f"{sides}-sided polygon drawn"

# ============================================================
# SEMANTIC / HIGH-LEVEL HELPERS
# ============================================================

def draw_slot(length, width):
    _require_sketch_active()
    sm = _sm()

    r = width / 2
    half_len = length / 2

    sm.CreateArc(
        -half_len / 1000, 0, 0,
        -half_len / 1000, r / 1000, 0,
        -half_len / 1000, -r / 1000, 0,
        1
    )

    sm.CreateArc(
        half_len / 1000, 0, 0,
        half_len / 1000, -r / 1000, 0,
        half_len / 1000, r / 1000, 0,
        1
    )

    sm.CreateLine(
        -half_len / 1000, r / 1000, 0,
        half_len / 1000, r / 1000, 0
    )
    sm.CreateLine(
        half_len / 1000, -r / 1000, 0,
        -half_len / 1000, -r / 1000, 0
    )

    return f"Slot {length}x{width} drawn"

def draw_symmetric_circles(offset_x, radius):
    _require_sketch_active()
    sm = _sm()
    sm.CreateCircleByRadius(offset_x / 1000, 0, 0, radius / 1000)
    sm.CreateCircleByRadius(-offset_x / 1000, 0, 0, radius / 1000)
    return f"Symmetric circles radius {radius} drawn"

# ============================================================
# AUTOMATIC CONSTRAINTS
# ============================================================

def constrain_to_origin():
    _require_sketch_active()
    model = _model()
    sm = model.SketchManager
    sk = _active_sketch()

    entities = sk.GetSketchSegments()
    if not entities:
        raise Exception("No sketch entities to constrain")

    entities[0].Select(False)
    model.Extension.SelectByID2(
        "Origin",
        "SKETCHPOINT",
        0, 0, 0,
        True, 0, None, 0
    )

    sm.AddConstraint("sgCOINCIDENT")
    return "Sketch constrained to origin"

def auto_horizontal_vertical():
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
    """
    Ensures sketch is a closed profile before features
    """
    _require_sketch_active()
    sk = _active_sketch()

    if sk is None:
        raise Exception("No active sketch found")

    if not sk.IsClosed():
        raise Exception("Sketch is NOT a closed profile")

    return "Sketch validated: closed profile"
