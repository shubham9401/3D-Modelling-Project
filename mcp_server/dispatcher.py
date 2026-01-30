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
    "create_sketch_on_selected_face": sketch.create_sketch_on_selected_face,
    "exit_sketch": sketch.exit_sketch,
    
    # Face selection
    "select_face_by_normal": sketch.select_face_by_normal,
    "select_face_at_coordinate": sketch.select_face_at_coordinate,
    
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
    "draw_triangle": sketch.draw_triangle,
    
    # Sketch validation
    "validate_closed_profile": sketch.validate_closed_profile,
    
    # Features
    "extrude": feature.extrude,
    "extrude_midplane": feature.extrude_midplane,
    "cut_extrude": feature.cut_extrude,
    "cut_through_all": feature.cut_through_all,
    "revolve": feature.revolve,
    "revolve_simple": feature.revolve_simple,
    "shell": feature.shell,
    "loft": feature.loft,
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
        content = f.read()
        # Basic cleanup if Markdown is present
        if "```json" in content:
            content = content.split("```json")[1].split("```")[0].strip()
        elif "```" in content:
            content = content.split("```")[1].strip()
        return json.loads(content)

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
        
        # ADD THIS DEBUG CODE:
        if tool_name == "create_sketch_on_selected_face":
            from tools.solidworks_app import get_active_model
            model = get_active_model()
            sel_count = model.SelectionManager.GetSelectedObjectCount
            if callable(sel_count):
                sel_count = sel_count()
            print(f"         🔍 DEBUG: Selection count before execution = {sel_count}")
        
        try:
            result = execute_action(action)
            print(f"         ✅ {result}\n")
        except Exception as e:
            print(f"         ❌ FAILED: {e}\n")
            
            # ADD THIS TO SEE FULL ERROR:
            import traceback
            print("Full traceback:")
            traceback.print_exc()
            
            print("⛔ Mission aborted due to error.")
            return False
    
    print("🎉 Mission completed successfully!")
    return True