"""
Dispatcher: Executes the AI-generated mission plan in SolidWorks.
"""

import json
import sys
import os

# Add project root to path for imports
current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(current_dir)
sys.path.append(project_root)

from tools import part, sketch, feature

# ============================================================
# TOOL REGISTRY
# ============================================================

TOOL_REGISTRY = {
    # Part lifecycle
    "create_part": part.create_part,
    "save_part": part.save_part,
    
    # Sketch lifecycle
    "create_sketch": sketch.create_sketch,
    "create_sketch_on_top_face": sketch.create_sketch_on_top_face,
    "create_sketch_on_selected_face": sketch.create_sketch_on_selected_face,
    "exit_sketch": sketch.exit_sketch,
    
    # Face selection
    "select_face_by_normal": sketch.select_face_by_normal,
    
    # Sketch primitives
    "draw_line": sketch.draw_line,
    "draw_centerline_vertical": sketch.draw_centerline_vertical,
    "draw_rectangle": sketch.draw_rectangle,
    "draw_circle": sketch.draw_circle,
    "draw_semicircle": sketch.draw_semicircle,
    "draw_arc": sketch.draw_arc,
    "draw_ellipse": sketch.draw_ellipse,
    "draw_polygon": sketch.draw_polygon,
    "draw_slot": sketch.draw_slot,
    
    # Sketch validation
    "validate_closed_profile": sketch.validate_closed_profile,
    
    # Features
    "extrude": feature.extrude,
    "extrude_midplane": feature.extrude_midplane,
    "cut_extrude": feature.cut_extrude,
    "cut_through_all": feature.cut_through_all,
    "revolve": feature.revolve,
    "fillet": feature.fillet,
    "chamfer": feature.chamfer,
    
    # Patterns
    "linear_pattern": feature.linear_pattern,
    "circular_pattern": feature.circular_pattern,
    "mirror_feature": feature.mirror_feature,
}

# ============================================================
# DISPATCHER LOGIC
# ============================================================

def load_mission(filepath: str) -> list:
    with open(filepath, "r") as f:
        return json.load(f)

def execute_action(action: dict) -> str:
    tool_name = action.get("tool")
    args = action.get("args", {})
    
    if tool_name not in TOOL_REGISTRY:
        raise ValueError(f"Unknown tool: {tool_name}")
    
    func = TOOL_REGISTRY[tool_name]
    result = func(**args)
    return result

def run_mission(filepath: str):
    print(f"📂 Loading mission from: {filepath}")
    
    try:
        actions = load_mission(filepath)
    except Exception as e:
        print(f"❌ Error loading mission: {e}")
        return False
    
    print(f"📋 Mission contains {len(actions)} steps\n")
    
    for i, action in enumerate(actions, 1):
        tool_name = action.get("tool", "unknown")
        args = action.get("args", {})
        
        print(f"[{i}/{len(actions)}] Executing: {tool_name}")
        print(f"         Args: {args}")
        
        try:
            result = execute_action(action)
            print(f"         ✅ {result}\n")
        except Exception as e:
            print(f"         ❌ FAILED: {e}\n")
            print("⛔ Mission aborted due to error.")
            return False
    
    print("🎉 Mission completed successfully!")
    return True

if __name__ == "__main__":
    mission_file = os.path.join(project_root, "mission.json")
    if len(sys.argv) > 1:
        mission_file = sys.argv[1]
    
    print("=" * 50)
    print("🚀 SOLIDWORKS AI DISPATCHER")
    print("=" * 50)
    
    run_mission(mission_file)