"""
Training Data Generator for SolidWorks AI Agent Fine-Tuning

Builds a JSONL dataset from:
  1. Existing examples in system_prompt.py
  2. mission_*_test.json files
  3. Programmatic dimension variations
  4. Manual examples from training_examples/

Usage:
    python generate_training_data.py              # Generate training_data.jsonl
    python generate_training_data.py --validate    # Validate the generated JSONL
    python generate_training_data.py --stats       # Print dataset statistics
"""

import json
import os
import re
import sys
import random
import hashlib
from pathlib import Path

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

OUTPUT_FILE = "training_data.jsonl"
MANUAL_DIR = "training_examples"
PROJECT_DIR = Path(__file__).parent

# Known tools from the system (used for validation)
KNOWN_TOOLS = {
    "create_part", "create_sketch", "create_sketch_on_selected_face",
    "select_face_by_normal", "select_face_at_coordinate",
    "select_edge_at_coordinate", "select_edge_at_coordinate_append",
    "draw_centerline_vertical", "draw_line", "draw_rectangle",
    "draw_circle", "draw_ellipse", "draw_arc", "draw_semicircle",
    "draw_triangle", "draw_polygon", "draw_hexagon", "draw_slot",
    "draw_spline",
    "validate_closed_profile", "exit_sketch",
    "extrude", "extrude_midplane", "cut_extrude", "cut_through_all",
    "revolve", "revolve_simple",
    "shell", "loft", "sweep",
    "create_reference_plane", "select_sketch",
    "fillet", "chamfer",
    "thread", "thread_tap",
    "sheet_metal_base_flange", "edge_flange",
    "linear_pattern", "circular_pattern",
    "delete_feature", "get_feature_tree",
}

# Condensed system prompt for fine-tuning (rules + tools only, NO examples)
CONDENSED_SYSTEM_PROMPT = """You are a SolidWorks CAD expert. Convert natural language requests into JSON arrays of SolidWorks tool calls.

RULES:
1. Origin (0,0,0) is the center of your first sketch.
2. Top Plane (XZ) -> Height is Y. Front Plane (XY) -> Depth is Z. Right Plane (YZ) -> Width is X.
3. ALWAYS call validate_closed_profile after drawing and BEFORE extrude/cut/revolve.
4. Use select_face_at_coordinate(x,y,z) to pick exact faces. NEVER guess.
5. If extruded X mm UP, top face is at (0, X, 0).
6. For symmetric features (legs, holes), use ±W/2 and ±H/2 coordinates.
7. BOLTS use thread (external). NUTS use thread_tap (internal) after cutting center hole.
8. Output ONLY a JSON array. No explanations, no markdown, no text before or after.

FEATURES:
- EXTRUDE: straight constant cross-section (boxes, cylinders, plates)
- REVOLVE: circular shapes around axis (spheres, cones, bowls)
- LOFT: transition between profiles (funnels, tapered shapes)
- SWEEP: follow a path (handles, pipes, tubes)

SWEEP: Profile on Front/RefPlane, Path on Right/Front, both share origin point.
LOFT: Two sketches on parallel planes, select both, then loft.

draw_hexagon radius = center-to-VERTEX. For across-flats AF: radius = AF / 1.732

AVAILABLE TOOLS:
create_part()
create_sketch(plane: "Front"|"Top"|"Right"|"PlaneN")
create_sketch_on_selected_face()
select_face_at_coordinate(x, y, z)
select_face_by_normal(direction: "up"|"down"|"front"|"back"|"left"|"right")
select_edge_at_coordinate(x, y, z)
select_edge_at_coordinate_append(x, y, z)
draw_centerline_vertical()
draw_line(x1, y1, x2, y2)
draw_rectangle(width, height, x=0, y=0)
draw_circle(radius, x=0, y=0)
draw_ellipse(radius_x, radius_y, x=0, y=0)
draw_arc(radius, start_angle, end_angle)
draw_semicircle(radius)
draw_triangle(base, height)
draw_polygon(sides, radius)
draw_hexagon(radius, x=0, y=0)
draw_slot(length, width)
draw_spline(points=[[x1,y1],[x2,y2],...])
validate_closed_profile()
exit_sketch()
extrude(depth)
extrude_midplane(depth)
cut_extrude(depth)
cut_through_all()
revolve(angle=360)
revolve_simple(angle=360)
shell(thickness)
loft()
sweep()
create_reference_plane(offset, plane)
select_sketch(sketch_name, mark, append)
fillet(radius)
chamfer(distance, angle)
thread(diameter, pitch, depth)
thread_tap(diameter, pitch, depth)
sheet_metal_base_flange(thickness, bend_radius, depth)
edge_flange(length, angle)
linear_pattern(count, spacing)
circular_pattern(count, angle)
delete_feature(feature_name)
get_feature_tree()
"""


# ---------------------------------------------------------------------------
# Source 1: Extract examples from system_prompt.py
# ---------------------------------------------------------------------------

def extract_examples_from_system_prompt():
    """Parse system_prompt.py to extract prompt->JSON example pairs."""
    examples = []

    prompt_path = PROJECT_DIR / "system_prompt.py"
    with open(prompt_path, "r", encoding="utf-8") as f:
        content = f.read()

    # Extract the SYSTEM_INSTRUCTION string
    match = re.search(r'SYSTEM_INSTRUCTION\s*=\s*"""(.*?)"""', content, re.DOTALL)
    if not match:
        print("⚠️  Could not find SYSTEM_INSTRUCTION in system_prompt.py")
        return examples

    instruction = match.group(1)

    # Pattern: **Description:** followed by JSON array
    # Find labeled JSON blocks like "**Simple Box:**" or "**Sphere (20mm diameter):**"
    pattern = r'\*\*([^*]+?)(?:\(([^)]+)\))?:\*\*\s*(?:NOTE:[^\n]*\n)?(?:MANDATORY:[^\n]*\n)?\s*\[(\s*\{.*?\}(?:\s*,\s*\{.*?\})*\s*)\]'

    blocks = re.finditer(pattern, instruction, re.DOTALL)

    for block in blocks:
        title = block.group(1).strip()
        params = block.group(2) or ""
        json_str = "[" + block.group(3) + "]"

        # Clean up the JSON
        try:
            actions = json.loads(json_str)
        except json.JSONDecodeError:
            # Try to fix common issues
            json_str = json_str.replace("'", '"')
            try:
                actions = json.loads(json_str)
            except json.JSONDecodeError:
                continue

        # Generate a natural language prompt from the title
        prompt = _title_to_prompt(title, params)
        if prompt and actions:
            examples.append({
                "prompt": prompt,
                "actions": actions,
                "source": "system_prompt"
            })

    return examples


def _title_to_prompt(title, params):
    """Convert an example title to a natural language prompt."""
    title = title.strip().rstrip(":")

    # Map known titles to natural prompts
    title_map = {
        "Simple Box": "Create a simple box",
        "Sphere": f"Create a sphere ({params})" if params else "Create a sphere",
        "Cone": f"Create a cone ({params})" if params else "Create a cone",
        "Mug with Curved Handle": "Create a mug with a curved handle",
        "Vase": f"Create a vase ({params})" if params else "Create a vase",
        "Washer": f"Create a washer ({params})" if params else "Create a washer",
        "Box with hole": "Create a box with a hole in the top",
        "Chair": "Create a chair with 4 legs and a backrest",
        "Stool": "Create a round stool with 4 legs",
        "Table": "Create a table with 4 legs",
        "Table with Rounded Edges": "Create a table with rounded edges and 4 legs",
        "Cupboard": "Create a cupboard with a shelf and a door",
        "Spur Gear": f"Create a spur gear ({params})" if params else "Create a spur gear",
        "Box with Filleted Edges": f"Create a box with filleted edges ({params})" if params else "Create a box with filleted edges",
        "Threaded Bolt": f"Create a threaded bolt ({params})" if params else "Create a threaded bolt",
        "Hex Head Bolt": f"Create a hex head bolt ({params})" if params else "Create a hex head bolt",
        "Hex Nut": f"Create a hex nut ({params})" if params else "Create a hex nut",
        "Curved Handle": "Create a curved handle using sweep",
        "Funnel": "Create a funnel using loft",
    }

    # Try exact match first
    for key, prompt in title_map.items():
        if key.lower() == title.lower():
            return prompt

    # Try partial match
    for key, prompt in title_map.items():
        if key.lower() in title.lower():
            return prompt

    # Fallback: use title as prompt
    if title and len(title) > 3:
        return f"Create a {title.lower()}"
    return None


# ---------------------------------------------------------------------------
# Source 2: Extract from mission test files
# ---------------------------------------------------------------------------

def extract_examples_from_test_missions():
    """Load mission_*_test.json files as training examples."""
    examples = []

    test_files = {
        "mission_nut_test.json": "Create an M6 hex nut with internal threads",
        "mission_loft_test.json": "Create a loft from a 100mm square to a 60mm circle",
        "mission_handle_test.json": "Create a curved handle using sweep",
        "mission_sweep_test.json": "Create a swept tube along a curved path",
        "mission_chamfer_test.json": "Create a box with chamfered edges",
        "mission_thread_test.json": "Create a threaded cylinder",
        "mission_handle_offset.json": "Create a handle with offset reference plane",
    }

    for filename, prompt in test_files.items():
        filepath = PROJECT_DIR / filename
        if filepath.exists():
            try:
                with open(filepath, "r", encoding="utf-8") as f:
                    actions = json.load(f)
                examples.append({
                    "prompt": prompt,
                    "actions": actions,
                    "source": "test_mission"
                })
            except (json.JSONDecodeError, IOError) as e:
                print(f"⚠️  Could not load {filename}: {e}")

    return examples


# ---------------------------------------------------------------------------
# Source 3: Programmatic dimension variations
# ---------------------------------------------------------------------------

def generate_variations():
    """Generate parametric variations of common shapes to expand the dataset."""
    examples = []

    # --- Simple boxes with varying dimensions ---
    box_sizes = [
        (30, 30, 20), (50, 50, 30), (80, 60, 40), (100, 100, 50),
        (120, 80, 60), (150, 100, 25), (200, 150, 30), (40, 40, 40),
        (75, 50, 35), (60, 60, 15),
    ]
    for w, h, d in box_sizes:
        prompt = f"Create a {w}x{h}mm rectangular plate, {d}mm thick"
        actions = [
            {"tool": "create_part", "args": {}},
            {"tool": "create_sketch", "args": {"plane": "Top"}},
            {"tool": "draw_rectangle", "args": {"width": w, "height": h}},
            {"tool": "validate_closed_profile", "args": {}},
            {"tool": "extrude", "args": {"depth": d}},
        ]
        examples.append({"prompt": prompt, "actions": actions, "source": "variation"})

    # --- Cubes ---
    for size in [10, 20, 30, 50, 75, 100, 150]:
        prompt = f"Create a {size}mm cube"
        actions = [
            {"tool": "create_part", "args": {}},
            {"tool": "create_sketch", "args": {"plane": "Top"}},
            {"tool": "draw_rectangle", "args": {"width": size, "height": size}},
            {"tool": "validate_closed_profile", "args": {}},
            {"tool": "extrude", "args": {"depth": size}},
        ]
        examples.append({"prompt": prompt, "actions": actions, "source": "variation"})

    # --- Cylinders ---
    cyl_sizes = [
        (5, 20), (10, 30), (15, 50), (20, 40), (25, 60),
        (30, 100), (40, 80), (50, 120),
    ]
    for r, h in cyl_sizes:
        prompt = f"Create a cylinder with {r*2}mm diameter and {h}mm height"
        actions = [
            {"tool": "create_part", "args": {}},
            {"tool": "create_sketch", "args": {"plane": "Top"}},
            {"tool": "draw_circle", "args": {"radius": r}},
            {"tool": "validate_closed_profile", "args": {}},
            {"tool": "extrude", "args": {"depth": h}},
        ]
        examples.append({"prompt": prompt, "actions": actions, "source": "variation"})

    # --- Spheres ---
    for r in [5, 10, 15, 20, 25, 30, 50]:
        prompt = f"Create a sphere with {r*2}mm diameter"
        actions = [
            {"tool": "create_part", "args": {}},
            {"tool": "create_sketch", "args": {"plane": "Front"}},
            {"tool": "draw_centerline_vertical", "args": {}},
            {"tool": "draw_semicircle", "args": {"radius": r}},
            {"tool": "validate_closed_profile", "args": {}},
            {"tool": "revolve", "args": {"angle": 360}},
        ]
        examples.append({"prompt": prompt, "actions": actions, "source": "variation"})

    # --- Washers ---
    washer_sizes = [
        (10, 4, 2, "M8"),  (12, 5, 2.5, "M10"), (8, 3.5, 1.5, "M6"),
        (16, 6, 3, "M12"), (20, 8, 3, "M16"),
    ]
    for od_r, id_r, thick, label in washer_sizes:
        prompt = f"Create an {label} washer (OD={od_r*2}mm, ID={id_r*2}mm, {thick}mm thick)"
        actions = [
            {"tool": "create_part", "args": {}},
            {"tool": "create_sketch", "args": {"plane": "Top"}},
            {"tool": "draw_circle", "args": {"radius": od_r}},
            {"tool": "validate_closed_profile", "args": {}},
            {"tool": "extrude", "args": {"depth": thick}},
            {"tool": "select_face_at_coordinate", "args": {"x": 0, "y": thick, "z": 0}},
            {"tool": "create_sketch_on_selected_face", "args": {}},
            {"tool": "draw_circle", "args": {"radius": id_r}},
            {"tool": "validate_closed_profile", "args": {}},
            {"tool": "cut_through_all", "args": {}},
        ]
        examples.append({"prompt": prompt, "actions": actions, "source": "variation"})

    # --- Tables with varying dimensions ---
    table_sizes = [
        (600, 400, 25, 20, 600),
        (1000, 700, 30, 25, 750),
        (1200, 800, 35, 30, 720),
        (500, 500, 20, 20, 500),
    ]
    for tw, th, tt, lr, lh in table_sizes:
        inset = lr + 10
        lx = tw // 2 - inset
        ly = th // 2 - inset
        prompt = f"Create a {tw}x{th}mm table with {tt}mm thick top and {lh}mm tall legs"
        actions = [
            {"tool": "create_part", "args": {}},
            {"tool": "create_sketch", "args": {"plane": "Top"}},
            {"tool": "draw_rectangle", "args": {"width": tw, "height": th}},
            {"tool": "validate_closed_profile", "args": {}},
            {"tool": "extrude", "args": {"depth": tt}},
            {"tool": "create_sketch", "args": {"plane": "Top"}},
            {"tool": "draw_circle", "args": {"radius": lr, "x": lx, "y": ly}},
            {"tool": "draw_circle", "args": {"radius": lr, "x": -lx, "y": ly}},
            {"tool": "draw_circle", "args": {"radius": lr, "x": lx, "y": -ly}},
            {"tool": "draw_circle", "args": {"radius": lr, "x": -lx, "y": -ly}},
            {"tool": "validate_closed_profile", "args": {}},
            {"tool": "extrude", "args": {"depth": -lh}},
        ]
        examples.append({"prompt": prompt, "actions": actions, "source": "variation"})

    # --- Hex nuts with varying sizes ---
    nut_specs = [
        ("M3", 3, 0.5, 3.46, 1.5, 2.4),   # M3x0.5
        ("M4", 4, 0.7, 4.04, 2.0, 3.2),   # M4x0.7
        ("M5", 5, 0.8, 4.62, 2.5, 4.0),   # M5x0.8
        ("M8", 8, 1.25, 7.51, 4.0, 6.5),  # M8x1.25
        ("M10", 10, 1.5, 9.24, 5.0, 8.0), # M10x1.5
    ]
    for label, dia, pitch, hex_r, hole_r, thick in nut_specs:
        prompt = f"Create an {label} hex nut"
        actions = [
            {"tool": "create_part", "args": {}},
            {"tool": "create_sketch", "args": {"plane": "Top"}},
            {"tool": "draw_hexagon", "args": {"radius": round(hex_r, 2)}},
            {"tool": "validate_closed_profile", "args": {}},
            {"tool": "extrude", "args": {"depth": thick}},
            {"tool": "select_face_at_coordinate", "args": {"x": 0, "y": thick, "z": 0}},
            {"tool": "create_sketch_on_selected_face", "args": {}},
            {"tool": "draw_circle", "args": {"radius": hole_r}},
            {"tool": "validate_closed_profile", "args": {}},
            {"tool": "cut_through_all", "args": {}},
            {"tool": "select_edge_at_coordinate", "args": {"x": hole_r, "y": thick, "z": 0}},
            {"tool": "thread_tap", "args": {"diameter": dia, "pitch": pitch, "depth": thick}},
        ]
        examples.append({"prompt": prompt, "actions": actions, "source": "variation"})

    # --- Hex bolts with varying sizes ---
    bolt_specs = [
        ("M6", 6, 1.0, 5.77, 3, 5, 30, 25),
        ("M8", 8, 1.25, 7.51, 4, 6, 40, 35),
        ("M10", 10, 1.5, 9.24, 5, 8, 50, 45),
        ("M12", 12, 1.75, 10.97, 6, 10, 60, 50),
    ]
    for label, dia, pitch, hex_r, shaft_r, head_h, shaft_len, thread_depth in bolt_specs:
        prompt = f"Create an {label} hex head bolt, {shaft_len}mm long"
        actions = [
            {"tool": "create_part", "args": {}},
            {"tool": "create_sketch", "args": {"plane": "Top"}},
            {"tool": "draw_hexagon", "args": {"radius": round(hex_r, 2)}},
            {"tool": "validate_closed_profile", "args": {}},
            {"tool": "extrude", "args": {"depth": head_h}},
            {"tool": "select_face_at_coordinate", "args": {"x": 0, "y": head_h, "z": 0}},
            {"tool": "create_sketch_on_selected_face", "args": {}},
            {"tool": "draw_circle", "args": {"radius": shaft_r}},
            {"tool": "validate_closed_profile", "args": {}},
            {"tool": "extrude", "args": {"depth": shaft_len}},
            {"tool": "select_edge_at_coordinate", "args": {"x": shaft_r, "y": head_h, "z": 0}},
            {"tool": "thread", "args": {"diameter": dia, "pitch": pitch, "depth": thread_depth}},
        ]
        examples.append({"prompt": prompt, "actions": actions, "source": "variation"})

    # --- Cones ---
    for base_r, height in [(5, 30), (10, 50), (15, 60), (20, 80), (25, 100)]:
        prompt = f"Create a cone with {base_r*2}mm base diameter and {height}mm height"
        actions = [
            {"tool": "create_part", "args": {}},
            {"tool": "create_sketch", "args": {"plane": "Front"}},
            {"tool": "draw_triangle", "args": {"base": base_r, "height": height}},
            {"tool": "validate_closed_profile", "args": {}},
            {"tool": "revolve", "args": {"angle": 360}},
        ]
        examples.append({"prompt": prompt, "actions": actions, "source": "variation"})

    # --- Boxes with holes ---
    for w, h, d, hole_r in [(60, 60, 30, 10), (80, 80, 40, 15), (100, 100, 50, 20), (120, 80, 35, 12)]:
        prompt = f"Create a {w}x{h}x{d}mm box with a {hole_r*2}mm hole through the top"
        actions = [
            {"tool": "create_part", "args": {}},
            {"tool": "create_sketch", "args": {"plane": "Top"}},
            {"tool": "draw_rectangle", "args": {"width": w, "height": h}},
            {"tool": "validate_closed_profile", "args": {}},
            {"tool": "extrude", "args": {"depth": d}},
            {"tool": "select_face_at_coordinate", "args": {"x": 0, "y": d, "z": 0}},
            {"tool": "create_sketch_on_selected_face", "args": {}},
            {"tool": "draw_circle", "args": {"radius": hole_r}},
            {"tool": "validate_closed_profile", "args": {}},
            {"tool": "cut_through_all", "args": {}},
        ]
        examples.append({"prompt": prompt, "actions": actions, "source": "variation"})

    # --- Lofted shapes (vases/funnels) ---
    for base_r, top_r, height in [(40, 20, 150), (30, 15, 100), (50, 25, 200), (35, 10, 120)]:
        prompt = f"Create a vase {base_r*2}mm base, {top_r*2}mm top, {height}mm tall"
        actions = [
            {"tool": "create_part", "args": {}},
            {"tool": "create_sketch", "args": {"plane": "Top"}},
            {"tool": "draw_circle", "args": {"radius": base_r}},
            {"tool": "validate_closed_profile", "args": {}},
            {"tool": "exit_sketch", "args": {}},
            {"tool": "create_reference_plane", "args": {"offset": height, "plane": "Top"}},
            {"tool": "create_sketch", "args": {"plane": "Plane1"}},
            {"tool": "draw_circle", "args": {"radius": top_r}},
            {"tool": "validate_closed_profile", "args": {}},
            {"tool": "exit_sketch", "args": {}},
            {"tool": "select_sketch", "args": {"sketch_name": "Sketch1", "mark": 1, "append": False}},
            {"tool": "select_sketch", "args": {"sketch_name": "Sketch2", "mark": 1, "append": True}},
            {"tool": "loft", "args": {}},
            {"tool": "select_face_at_coordinate", "args": {"x": 0, "y": height, "z": 0}},
            {"tool": "shell", "args": {"thickness": 3}},
        ]
        examples.append({"prompt": prompt, "actions": actions, "source": "variation"})

    # --- Cups/mugs with handle ---
    for r, h in [(35, 90), (40, 100), (45, 110), (50, 120)]:
        prompt = f"Create a coffee mug with {r*2}mm diameter and {h}mm height with a handle"
        actions = [
            {"tool": "create_part", "args": {}},
            {"tool": "create_sketch", "args": {"plane": "Top"}},
            {"tool": "draw_circle", "args": {"radius": r}},
            {"tool": "validate_closed_profile", "args": {}},
            {"tool": "extrude", "args": {"depth": h}},
            {"tool": "select_face_at_coordinate", "args": {"x": 0, "y": h, "z": 0}},
            {"tool": "shell", "args": {"thickness": 3}},
            {"tool": "create_reference_plane", "args": {"plane": "Right", "offset": r}},
            {"tool": "create_sketch", "args": {"plane": "Plane1"}},
            {"tool": "draw_circle", "args": {"radius": 5, "x": 0, "y": round(h * 0.25)}},
            {"tool": "validate_closed_profile", "args": {}},
            {"tool": "exit_sketch", "args": {}},
            {"tool": "create_sketch", "args": {"plane": "Front"}},
            {"tool": "draw_spline", "args": {"points": [
                [r, round(h * 0.25)],
                [r + 25, round(h * 0.5)],
                [r, round(h * 0.75)]
            ]}},
            {"tool": "exit_sketch", "args": {}},
            {"tool": "select_sketch", "args": {"sketch_name": "Sketch2", "mark": 1, "append": False}},
            {"tool": "select_sketch", "args": {"sketch_name": "Sketch3", "mark": 4, "append": True}},
            {"tool": "sweep", "args": {}},
        ]
        examples.append({"prompt": prompt, "actions": actions, "source": "variation"})

    # --- Boxes with fillets ---
    for size, fillet_r in [(40, 3), (50, 5), (60, 8), (80, 10), (100, 12)]:
        depth = size // 2 + 10
        prompt = f"Create a {size}x{size}x{depth}mm box with {fillet_r}mm fillet on top edges"
        actions = [
            {"tool": "create_part", "args": {}},
            {"tool": "create_sketch", "args": {"plane": "Top"}},
            {"tool": "draw_rectangle", "args": {"width": size, "height": size}},
            {"tool": "validate_closed_profile", "args": {}},
            {"tool": "extrude", "args": {"depth": depth}},
            {"tool": "select_edge_at_coordinate", "args": {"x": size // 2, "y": depth, "z": 0}},
            {"tool": "fillet", "args": {"radius": fillet_r}},
        ]
        examples.append({"prompt": prompt, "actions": actions, "source": "variation"})

    # --- Pipes (hollow cylinders) ---
    for od_r, wall, length in [(15, 2, 100), (20, 3, 150), (25, 3, 200), (30, 4, 250)]:
        prompt = f"Create a pipe OD={od_r*2}mm, wall thickness {wall}mm, {length}mm long"
        actions = [
            {"tool": "create_part", "args": {}},
            {"tool": "create_sketch", "args": {"plane": "Top"}},
            {"tool": "draw_circle", "args": {"radius": od_r}},
            {"tool": "validate_closed_profile", "args": {}},
            {"tool": "extrude", "args": {"depth": length}},
            {"tool": "select_face_at_coordinate", "args": {"x": 0, "y": length, "z": 0}},
            {"tool": "shell", "args": {"thickness": wall}},
        ]
        examples.append({"prompt": prompt, "actions": actions, "source": "variation"})

    # --- Stepped shafts ---
    shaft_specs = [
        (20, 30, 15, 50, 10, 20),
        (25, 40, 18, 60, 12, 30),
        (30, 50, 20, 70, 15, 25),
        (15, 25, 10, 40, 8, 15),
    ]
    for d1, l1, d2, l2, d3, l3 in shaft_specs:
        r1, r2, r3 = d1//2, d2//2, d3//2
        prompt = f"Create a stepped shaft: {d1}mm x {l1}mm, then {d2}mm x {l2}mm, then {d3}mm x {l3}mm"
        actions = [
            {"tool": "create_part", "args": {}},
            {"tool": "create_sketch", "args": {"plane": "Top"}},
            {"tool": "draw_circle", "args": {"radius": r1}},
            {"tool": "validate_closed_profile", "args": {}},
            {"tool": "extrude", "args": {"depth": l1}},
            {"tool": "select_face_at_coordinate", "args": {"x": 0, "y": l1, "z": 0}},
            {"tool": "create_sketch_on_selected_face", "args": {}},
            {"tool": "draw_circle", "args": {"radius": r2}},
            {"tool": "validate_closed_profile", "args": {}},
            {"tool": "extrude", "args": {"depth": l2}},
            {"tool": "select_face_at_coordinate", "args": {"x": 0, "y": l1 + l2, "z": 0}},
            {"tool": "create_sketch_on_selected_face", "args": {}},
            {"tool": "draw_circle", "args": {"radius": r3}},
            {"tool": "validate_closed_profile", "args": {}},
            {"tool": "extrude", "args": {"depth": l3}},
        ]
        examples.append({"prompt": prompt, "actions": actions, "source": "variation"})

    # --- Bowls ---
    for r, h in [(40, 30), (50, 40), (60, 50), (80, 60)]:
        prompt = f"Create a bowl {r*2}mm diameter {h}mm deep with 3mm walls"
        actions = [
            {"tool": "create_part", "args": {}},
            {"tool": "create_sketch", "args": {"plane": "Top"}},
            {"tool": "draw_circle", "args": {"radius": r}},
            {"tool": "validate_closed_profile", "args": {}},
            {"tool": "extrude", "args": {"depth": h}},
            {"tool": "select_face_at_coordinate", "args": {"x": 0, "y": h, "z": 0}},
            {"tool": "shell", "args": {"thickness": 3}},
        ]
        examples.append({"prompt": prompt, "actions": actions, "source": "variation"})

    # --- L-brackets ---
    for h_arm, v_arm, thick in [(60, 40, 3), (80, 60, 5), (100, 80, 5), (120, 100, 8)]:
        prompt = f"Create an L-bracket {h_arm}x{v_arm}mm with {thick}mm thickness"
        actions = [
            {"tool": "create_part", "args": {}},
            {"tool": "create_sketch", "args": {"plane": "Front"}},
            {"tool": "draw_line", "args": {"x1": 0, "y1": 0, "x2": h_arm, "y2": 0}},
            {"tool": "draw_line", "args": {"x1": h_arm, "y1": 0, "x2": h_arm, "y2": thick}},
            {"tool": "draw_line", "args": {"x1": h_arm, "y1": thick, "x2": thick, "y2": thick}},
            {"tool": "draw_line", "args": {"x1": thick, "y1": thick, "x2": thick, "y2": v_arm}},
            {"tool": "draw_line", "args": {"x1": thick, "y1": v_arm, "x2": 0, "y2": v_arm}},
            {"tool": "draw_line", "args": {"x1": 0, "y1": v_arm, "x2": 0, "y2": 0}},
            {"tool": "validate_closed_profile", "args": {}},
            {"tool": "extrude", "args": {"depth": 30}},
        ]
        examples.append({"prompt": prompt, "actions": actions, "source": "variation"})

    # --- Mounting blocks with holes ---
    for w, h, d, hole_r in [(40, 30, 20, 2.5), (60, 40, 30, 3), (80, 60, 40, 4), (100, 80, 50, 5)]:
        inset_x, inset_y = w//2 - 8, h//2 - 8
        prompt = f"Create a {w}x{h}x{d}mm mounting block with 4 corner holes"
        actions = [
            {"tool": "create_part", "args": {}},
            {"tool": "create_sketch", "args": {"plane": "Top"}},
            {"tool": "draw_rectangle", "args": {"width": w, "height": h}},
            {"tool": "validate_closed_profile", "args": {}},
            {"tool": "extrude", "args": {"depth": d}},
            {"tool": "select_face_at_coordinate", "args": {"x": 0, "y": d, "z": 0}},
            {"tool": "create_sketch_on_selected_face", "args": {}},
            {"tool": "draw_circle", "args": {"radius": hole_r, "x": inset_x, "y": inset_y}},
            {"tool": "draw_circle", "args": {"radius": hole_r, "x": -inset_x, "y": inset_y}},
            {"tool": "draw_circle", "args": {"radius": hole_r, "x": inset_x, "y": -inset_y}},
            {"tool": "draw_circle", "args": {"radius": hole_r, "x": -inset_x, "y": -inset_y}},
            {"tool": "validate_closed_profile", "args": {}},
            {"tool": "cut_through_all", "args": {}},
        ]
        examples.append({"prompt": prompt, "actions": actions, "source": "variation"})

    # --- Discs / spacers ---
    for od, id_r, thick in [(20, 8, 5), (30, 12, 8), (40, 15, 10), (50, 20, 12), (60, 25, 15)]:
        prompt = f"Create a spacer ring OD={od}mm ID={id_r*2}mm {thick}mm thick"
        actions = [
            {"tool": "create_part", "args": {}},
            {"tool": "create_sketch", "args": {"plane": "Top"}},
            {"tool": "draw_circle", "args": {"radius": od // 2}},
            {"tool": "validate_closed_profile", "args": {}},
            {"tool": "extrude", "args": {"depth": thick}},
            {"tool": "select_face_at_coordinate", "args": {"x": 0, "y": thick, "z": 0}},
            {"tool": "create_sketch_on_selected_face", "args": {}},
            {"tool": "draw_circle", "args": {"radius": id_r}},
            {"tool": "validate_closed_profile", "args": {}},
            {"tool": "cut_through_all", "args": {}},
        ]
        examples.append({"prompt": prompt, "actions": actions, "source": "variation"})

    # --- Alternative phrasings for common shapes ---
    alt_prompts = [
        ("Make a 50mm cube", [
            {"tool": "create_part", "args": {}},
            {"tool": "create_sketch", "args": {"plane": "Top"}},
            {"tool": "draw_rectangle", "args": {"width": 50, "height": 50}},
            {"tool": "validate_closed_profile", "args": {}},
            {"tool": "extrude", "args": {"depth": 50}},
        ]),
        ("Build a cylinder 40mm wide and 100mm tall", [
            {"tool": "create_part", "args": {}},
            {"tool": "create_sketch", "args": {"plane": "Top"}},
            {"tool": "draw_circle", "args": {"radius": 20}},
            {"tool": "validate_closed_profile", "args": {}},
            {"tool": "extrude", "args": {"depth": 100}},
        ]),
        ("I need a flat circular disc 80mm diameter 3mm thick", [
            {"tool": "create_part", "args": {}},
            {"tool": "create_sketch", "args": {"plane": "Top"}},
            {"tool": "draw_circle", "args": {"radius": 40}},
            {"tool": "validate_closed_profile", "args": {}},
            {"tool": "extrude", "args": {"depth": 3}},
        ]),
        ("Generate a hollow tube 30mm outer 20mm inner 150mm long", [
            {"tool": "create_part", "args": {}},
            {"tool": "create_sketch", "args": {"plane": "Top"}},
            {"tool": "draw_circle", "args": {"radius": 15}},
            {"tool": "validate_closed_profile", "args": {}},
            {"tool": "extrude", "args": {"depth": 150}},
            {"tool": "select_face_at_coordinate", "args": {"x": 0, "y": 150, "z": 0}},
            {"tool": "shell", "args": {"thickness": 5}},
        ]),
        ("A simple rectangular block 80x40x20", [
            {"tool": "create_part", "args": {}},
            {"tool": "create_sketch", "args": {"plane": "Top"}},
            {"tool": "draw_rectangle", "args": {"width": 80, "height": 40}},
            {"tool": "validate_closed_profile", "args": {}},
            {"tool": "extrude", "args": {"depth": 20}},
        ]),
    ]
    for prompt, actions in alt_prompts:
        examples.append({"prompt": prompt, "actions": actions, "source": "variation"})

    return examples


# ---------------------------------------------------------------------------
# Source 4: Manual examples from training_examples/ directory
# ---------------------------------------------------------------------------

def load_manual_examples():
    """Load hand-curated examples from training_examples/ directory."""
    examples = []
    manual_dir = PROJECT_DIR / MANUAL_DIR

    if not manual_dir.exists():
        return examples

    # Look for .txt + .json pairs
    for txt_file in manual_dir.glob("*.txt"):
        json_file = txt_file.with_suffix(".json")
        if json_file.exists():
            try:
                prompt = txt_file.read_text(encoding="utf-8").strip()
                actions = json.loads(json_file.read_text(encoding="utf-8"))
                if prompt and isinstance(actions, list):
                    examples.append({
                        "prompt": prompt,
                        "actions": actions,
                        "source": "manual"
                    })
            except (json.JSONDecodeError, IOError) as e:
                print(f"⚠️  Could not load {txt_file.stem}: {e}")

    return examples


# ---------------------------------------------------------------------------
# JSONL Builder
# ---------------------------------------------------------------------------

def build_jsonl(examples, output_path=None):
    """Convert examples to JSONL format and write to file."""
    output_path = output_path or (PROJECT_DIR / OUTPUT_FILE)

    # Deduplicate by prompt hash
    seen = set()
    unique_examples = []
    for ex in examples:
        key = hashlib.md5(ex["prompt"].lower().encode()).hexdigest()
        if key not in seen:
            seen.add(key)
            unique_examples.append(ex)

    # Shuffle for training
    random.seed(42)
    random.shuffle(unique_examples)

    lines = []
    for ex in unique_examples:
        row = {
            "messages": [
                {"role": "system", "content": CONDENSED_SYSTEM_PROMPT.strip()},
                {"role": "user", "content": ex["prompt"]},
                {"role": "assistant", "content": json.dumps(ex["actions"])},
            ]
        }
        lines.append(json.dumps(row, ensure_ascii=False))

    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")

    print(f"✅ Generated {len(lines)} training examples → {output_path}")
    return unique_examples


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------

def validate_jsonl(filepath=None):
    """Validate the generated JSONL file."""
    filepath = filepath or (PROJECT_DIR / OUTPUT_FILE)

    if not Path(filepath).exists():
        print(f"❌ File not found: {filepath}")
        print("   Run 'python generate_training_data.py' first.")
        return False

    errors = []
    total = 0
    tool_coverage = set()

    with open(filepath, "r", encoding="utf-8") as f:
        for line_num, line in enumerate(f, 1):
            total += 1
            line = line.strip()
            if not line:
                continue

            # Check valid JSON
            try:
                row = json.loads(line)
            except json.JSONDecodeError as e:
                errors.append(f"  Line {line_num}: Invalid JSON - {e}")
                continue

            # Check structure
            if "messages" not in row:
                errors.append(f"  Line {line_num}: Missing 'messages' key")
                continue

            messages = row["messages"]
            roles = [m.get("role") for m in messages]

            if roles != ["system", "user", "assistant"]:
                errors.append(f"  Line {line_num}: Expected [system, user, assistant], got {roles}")
                continue

            # Check assistant content is valid JSON array
            try:
                actions = json.loads(messages[2]["content"])
                if not isinstance(actions, list):
                    errors.append(f"  Line {line_num}: Assistant content is not a JSON array")
                    continue
            except json.JSONDecodeError:
                errors.append(f"  Line {line_num}: Assistant content is not valid JSON")
                continue

            # Check tool names
            for action in actions:
                tool = action.get("tool", "")
                tool_coverage.add(tool)
                if tool not in KNOWN_TOOLS:
                    errors.append(f"  Line {line_num}: Unknown tool '{tool}'")

    # Report
    print(f"\n{'='*50}")
    print(f"VALIDATION REPORT: {filepath}")
    print(f"{'='*50}")
    print(f"Total examples: {total}")
    print(f"Tool coverage:  {len(tool_coverage)}/{len(KNOWN_TOOLS)} tools")

    if errors:
        print(f"\n❌ {len(errors)} errors found:")
        for e in errors[:20]:
            print(e)
        if len(errors) > 20:
            print(f"  ... and {len(errors) - 20} more")
        return False
    else:
        print(f"\n✅ All {total} examples are valid!")
        if total < 50:
            print(f"⚠️  Dataset size ({total}) is below recommended minimum (50)")
        return True


# ---------------------------------------------------------------------------
# Statistics
# ---------------------------------------------------------------------------

def print_stats(examples):
    """Print dataset statistics."""
    print(f"\n{'='*50}")
    print("DATASET STATISTICS")
    print(f"{'='*50}")

    # By source
    sources = {}
    for ex in examples:
        src = ex.get("source", "unknown")
        sources[src] = sources.get(src, 0) + 1

    print(f"\nTotal examples: {len(examples)}")
    print(f"\nBy source:")
    for src, count in sorted(sources.items()):
        print(f"  {src:20s}: {count}")

    # Tool usage
    tool_counts = {}
    total_steps = 0
    for ex in examples:
        for action in ex["actions"]:
            tool = action.get("tool", "unknown")
            tool_counts[tool] = tool_counts.get(tool, 0) + 1
            total_steps += 1

    print(f"\nTotal steps across all examples: {total_steps}")
    print(f"Average steps per example: {total_steps / len(examples):.1f}")

    print(f"\nTool usage (top 15):")
    for tool, count in sorted(tool_counts.items(), key=lambda x: -x[1])[:15]:
        print(f"  {tool:35s}: {count}")

    # Tools NOT covered
    used_tools = set(tool_counts.keys())
    unused = KNOWN_TOOLS - used_tools
    if unused:
        print(f"\n⚠️  Tools NOT covered in training data ({len(unused)}):")
        for t in sorted(unused):
            print(f"  - {t}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    print("🔧 SolidWorks AI Agent — Training Data Generator")
    print("=" * 50)

    # Collect from all sources
    print("\n📥 Collecting examples...")

    examples = []

    src1 = extract_examples_from_system_prompt()
    print(f"  system_prompt.py:    {len(src1)} examples")
    examples.extend(src1)

    src2 = extract_examples_from_test_missions()
    print(f"  mission test files:  {len(src2)} examples")
    examples.extend(src2)

    src3 = generate_variations()
    print(f"  variations:          {len(src3)} examples")
    examples.extend(src3)

    src4 = load_manual_examples()
    print(f"  manual examples:     {len(src4)} examples")
    examples.extend(src4)

    print(f"\n  Total (pre-dedup):   {len(examples)}")

    # Build JSONL
    print("\n📝 Building JSONL...")
    unique = build_jsonl(examples)

    # Handle flags
    if "--validate" in sys.argv:
        validate_jsonl()

    if "--stats" in sys.argv:
        print_stats(unique)

    if "--validate" not in sys.argv and "--stats" not in sys.argv:
        # Always show stats on generation
        print_stats(unique)


if __name__ == "__main__":
    main()
