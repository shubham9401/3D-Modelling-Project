"""
LLM Client: Unified interface for any OpenAI-compatible LLM provider.

Configuration (in .env):
    LLM_API_KEY   = your_api_key_here
    LLM_MODEL     = llama-3.3-70b-versatile
    LLM_BASE_URL  = https://api.groq.com/openai/v1

Common base URLs:
    Groq:    https://api.groq.com/openai/v1
    Gemini:  https://generativelanguage.googleapis.com/v1beta/openai/
    OpenAI:  https://api.openai.com/v1
"""

import json
import os
from system_prompt import SYSTEM_INSTRUCTION, AVAILABLE_TOOLS

# --- CONFIGURATION (read from .env) ---

LLM_API_KEY  = os.environ.get("LLM_API_KEY", "")
LLM_MODEL    = os.environ.get("LLM_MODEL", "llama-3.3-70b-versatile")
LLM_BASE_URL = os.environ.get("LLM_BASE_URL", "https://api.groq.com/openai/v1")
OUTPUT_FILE  = "mission.json"


# --- UNIFIED LLM CALL ---

def _call_llm(system_prompt, user_prompt):
    """
    Calls any OpenAI-compatible API (Groq, Gemini, OpenAI, etc.)
    using the openai Python package.
    """
    from openai import OpenAI

    # Read config fresh from env (not module-level cache)
    api_key  = os.environ.get("LLM_API_KEY", "")
    model    = os.environ.get("LLM_MODEL", "llama-3.3-70b-versatile")
    base_url = os.environ.get("LLM_BASE_URL", "https://api.groq.com/openai/v1")

    if not api_key:
        raise ValueError(
            "LLM_API_KEY is not set!\n"
            "Add to .env:\n"
            '  LLM_API_KEY=your_key_here\n'
            '  LLM_MODEL=llama-3.3-70b-versatile\n'
            '  LLM_BASE_URL=https://api.groq.com/openai/v1'
        )

    client = OpenAI(
        api_key=api_key,
        base_url=base_url,
    )

    completion = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user",   "content": user_prompt},
        ],
        temperature=0.1,
    )

    content = completion.choices[0].message.content

    # Databricks (and some other APIs) can return content as a list of blocks, e.g. [{"type": "text", "text": "..."}]
    if isinstance(content, list):
        parts = []
        for block in content:
            if isinstance(block, str):
                parts.append(block)
            elif isinstance(block, dict) and "text" in block:
                parts.append(block["text"])
            elif hasattr(block, "text"):
                parts.append(block.text)
        content = "\n".join(parts) if parts else ""

    return content


# --- COMPLETENESS VALIDATION ---

# Rules: (keyword_in_request, required_tool, description)
COMPLETENESS_RULES = [
    # Fasteners - threads are mandatory
    (["nut"],                  "thread_tap",  "internal thread (thread_tap)"),
    (["bolt", "screw"],        "thread",      "external thread (thread)"),
    # Hollow objects need shell
    (["cup", "mug", "bowl", "vase", "container", "hollow"],
                               "shell",       "shell (hollow interior)"),
    # Furniture needs legs
    (["table"],                "extrude",     "legs (negative extrude)"),
    (["chair"],                "extrude",     "legs (negative extrude)"),
    # Every model needs create_part
    ([],                       "create_part", "create_part"),
]

def _check_completeness(actions, user_request):
    """
    Checks if the LLM output is complete based on the user request.
    Returns a list of missing items, or empty list if complete.
    """
    if not actions:
        return ["No actions generated"]

    tools_used = [a.get("tool", "") for a in actions]
    request_lower = user_request.lower()
    missing = []

    for keywords, required_tool, description in COMPLETENESS_RULES:
        # Skip rules that don't match the request
        if keywords and not any(kw in request_lower for kw in keywords):
            continue
        # Check if the required tool is present
        if required_tool not in tools_used:
            missing.append(description)

    # Special check: nut/bolt should have enough steps (not truncated)
    if any(kw in request_lower for kw in ["nut", "bolt"]):
        if len(actions) < 8:
            missing.append(f"too few steps ({len(actions)}) for a fastener - likely truncated")

    return missing


def get_agent_response(user_request, max_retries=1):
    """
    Sends user request to the LLM and returns CAD commands as JSON.
    Includes completeness validation with automatic retry.
    Applies post-processing to fix common LLM math errors.
    """
    full_system_message = f"{SYSTEM_INSTRUCTION}\n\nAVAILABLE TOOLS:\n{AVAILABLE_TOOLS}"

    print(f"🧠 Processing: '{user_request}' (model: {LLM_MODEL})...")

    try:
        content = _call_llm(full_system_message, user_request)
        actions = clean_and_validate_json(content)

        if not actions:
            return None

        # Check completeness
        missing = _check_completeness(actions, user_request)

        if missing and max_retries > 0:
            missing_str = ", ".join(missing)
            print(f"⚠️  Incomplete output detected! Missing: {missing_str}")
            print(f"🔄 Retrying with feedback...")

            retry_prompt = (
                f"Your previous output for \"{user_request}\" was INCOMPLETE.\n"
                f"MISSING STEPS: {missing_str}\n\n"
                f"Regenerate the COMPLETE JSON array with ALL steps including the missing ones.\n"
                f"Original request: {user_request}"
            )

            retry_content = _call_llm(full_system_message, retry_prompt)
            retry_actions = clean_and_validate_json(retry_content)

            if retry_actions:
                retry_missing = _check_completeness(retry_actions, user_request)
                if not retry_missing or len(retry_actions) > len(actions):
                    print(f"✅ Retry successful! {len(retry_actions)} steps (was {len(actions)})")
                    actions = retry_actions
                else:
                    print(f"⚠️  Retry still incomplete. Using best result.")
                    actions = retry_actions if len(retry_actions) >= len(actions) else actions

        elif missing:
            missing_str = ", ".join(missing)
            print(f"⚠️  WARNING: Output may be incomplete. Missing: {missing_str}")

        # ═══════════════════════════════════════════════════
        # POST-PROCESSING: Fix common LLM math errors
        # ═══════════════════════════════════════════════════
        actions = _postprocess_actions(actions, user_request)

        return actions

    except Exception as e:
        print(f"❌ API Error: {type(e).__name__}: {e}")
        if hasattr(e, 'status_code'):
            print(f"   HTTP Status: {e.status_code}")
        import traceback
        traceback.print_exc()
        return None



def get_modification_response(modification_request, model_summary):
    """
    Generates delta CAD commands to modify an existing model.
    Sends modification request + current model state to the LLM.
    """
    full_system_message = f"{SYSTEM_INSTRUCTION}\n\nAVAILABLE TOOLS:\n{AVAILABLE_TOOLS}"

    combined_prompt = f"""CURRENT MODEL STATE:
{model_summary}

MODIFICATION REQUEST:
{modification_request}

IMPORTANT MODIFICATION RULES:
1. The model already exists in SolidWorks and is open.
2. NEVER use create_part. The part is already open. You are ONLY adding, removing, or changing features on the EXISTING part.
3. For ADDITIVE modifications (fillet, chamfer, hole, pattern, etc.):
   - Generate ONLY the new steps needed (select edges/faces, then apply the feature).
4. For SHAPE changes (e.g., changing a circular seat to rectangular):
   - Use delete_feature to remove the old feature(s) that need to change (e.g., the circular sketch/extrude).
   - Then create a NEW sketch on the appropriate plane or face and draw the new shape.
   - Then extrude/revolve/loft as needed.
   - Do NOT create a new part. Work on the existing model.
5. For DIMENSION changes (resize, change height, etc.):
   - Use delete_feature to remove the feature that needs resizing.
   - Recreate it with the new dimensions on the same plane/face.
6. Study the MODEL STATE above carefully. Identify which features to keep and which to modify.
   - Use the EXACT feature names from the MODEL STATE (e.g., "Boss-Extrude1", NOT "Extrude1").
   - SolidWorks naming convention: "Boss-Extrude1", "Cut-Extrude1", "Boss-Revolve1", "Fillet1", etc.
7. If you need to select a face or edge, use select_face_at_coordinate or select_edge_at_coordinate.
8. Output ONLY the JSON array of steps. No explanations.
"""

    print(f"🔧 Generating modification steps (model: {LLM_MODEL})...")

    try:
        content = _call_llm(full_system_message, combined_prompt)
        actions = clean_and_validate_json(content)
        if actions:
            actions = _postprocess_actions(actions, modification_request)
        return actions
    except Exception as e:
        print(f"❌ API Error: {e}")
        return None


# ============================================================
# POST-PROCESSING: Fix LLM Math Errors Programmatically
# ============================================================

import math
import re

def _postprocess_actions(actions, user_request):
    """
    Post-processes the LLM-generated actions to fix common math errors.
    The LLM is bad at arithmetic — Python corrects coordinates here.
    """
    if not actions:
        return actions
    
    actions = _fix_loft_workflow(actions)
    actions = _fix_bottle_profile(actions, user_request)
    actions = _fix_gear_proportions(actions, user_request)
    actions = _fix_hole_boundaries(actions)
    
    return actions


def _fix_bottle_profile(actions, user_request):
    """
    Detects bottle/vase-like requests where the LLM used stacked extrusions
    and replaces with a proper revolve half-profile for a realistic shape.
    
    A realistic bottle has: flat bottom → body wall → tapered shoulder → neck → lip
    """
    request_lower = user_request.lower()
    
    # Check if this looks like a bottle/vase request
    bottle_keywords = ["bottle", "vase", "flask", "jar", "jug", "canteen"]
    if not any(kw in request_lower for kw in bottle_keywords):
        return actions
    
    # Check if LLM used stacked extrusions (the BAD approach)
    tools = [a.get("tool", "") for a in actions]
    extrude_count = tools.count("extrude")
    has_shell = "shell" in tools
    has_revolve = "revolve" in tools or "revolve_simple" in tools
    
    # If already using revolve, don't override
    if has_revolve:
        return actions
    
    # If using extrude-stack approach (2+ extrudes + shell), replace with revolve
    if extrude_count < 2:
        return actions
    
    # Parse dimensions from the user request
    body_radius = 35.0  # default
    neck_radius = 13.0
    body_height = 150.0
    total_height = 200.0
    wall_thickness = 2.0
    
    # Try to extract body diameter/radius
    m = re.search(r'(\d+(?:\.\d+)?)\s*mm\s*(?:body|base|bottom|diameter|wide)', request_lower)
    if m:
        val = float(m.group(1))
        body_radius = val / 2 if val > 50 else val
    
    m = re.search(r'(\d+(?:\.\d+)?)\s*mm\s*(?:neck|top|opening|mouth)', request_lower)
    if m:
        val = float(m.group(1))
        neck_radius = val / 2 if val > 30 else val
    
    m = re.search(r'(\d+(?:\.\d+)?)\s*mm\s*(?:tall|height|high)', request_lower)
    if m:
        total_height = float(m.group(1))
        body_height = total_height * 0.7  # Body is ~70% of total height
    
    # Build the revolve half-profile using draw_line
    # Profile on Front plane (X=radial, Y=height)
    shoulder_start = body_height
    shoulder_end = body_height + (total_height - body_height) * 0.4
    neck_start = shoulder_end
    neck_end = total_height - 5  # Lip starts 5mm below top
    lip_height = total_height
    lip_radius = neck_radius + 2  # Lip is slightly wider
    
    print(f"\n🔧 BOTTLE POST-PROCESSING:")
    print(f"   Replacing {extrude_count} stacked extrusions → revolve profile")
    print(f"   Body: r={body_radius}mm, h={body_height}mm")
    print(f"   Shoulder: {shoulder_start}→{shoulder_end}mm (tapered)")
    print(f"   Neck: r={neck_radius}mm, {neck_start}→{neck_end}mm") 
    print(f"   Total height: {total_height}mm")
    
    # Generate the revolve-based bottle profile
    # Half-profile: draw as lines going counterclockwise from origin
    new_actions = [
        {"tool": "create_part", "args": {}},
        {"tool": "create_sketch", "args": {"plane": "Front"}},
        # Bottom line (origin to body radius)
        {"tool": "draw_line", "args": {"x1": 0, "y1": 0, "x2": body_radius, "y2": 0}},
        # Body wall (vertical up)
        {"tool": "draw_line", "args": {"x1": body_radius, "y1": 0, "x2": body_radius, "y2": shoulder_start}},
        # Shoulder taper (diagonal inward to neck)
        {"tool": "draw_line", "args": {"x1": body_radius, "y1": shoulder_start, "x2": neck_radius, "y2": neck_start}},
        # Neck wall (vertical up)
        {"tool": "draw_line", "args": {"x1": neck_radius, "y1": neck_start, "x2": neck_radius, "y2": neck_end}},
        # Lip (slight outward flare)
        {"tool": "draw_line", "args": {"x1": neck_radius, "y1": neck_end, "x2": lip_radius, "y2": lip_height}},
        # Close back to axis (top to origin Y)
        {"tool": "draw_line", "args": {"x1": lip_radius, "y1": lip_height, "x2": 0, "y2": lip_height}},
        # Axis line (close the profile - top to bottom on Y axis)
        {"tool": "draw_line", "args": {"x1": 0, "y1": lip_height, "x2": 0, "y2": 0}},
        {"tool": "validate_closed_profile", "args": {}},
        {"tool": "revolve", "args": {"angle": 360}},
    ]
    
    # Add shell if the original had it
    if has_shell:
        # Shell from the top face (the opening)
        new_actions.append({"tool": "select_face_at_coordinate", "args": {"x": 0, "y": lip_height, "z": 0}})
        new_actions.append({"tool": "shell", "args": {"thickness": wall_thickness}})
    
    print(f"   Generated {len(new_actions)} revolve steps (was {len(actions)} extrude steps)")
    
    return new_actions

def _fix_loft_workflow(actions):
    """
    Fixes loft/bottle workflows by inserting exit_sketch before create_reference_plane.
    
    Problem: validate_closed_profile doesn't exit the sketch, but create_reference_plane
    needs no sketch active. This auto-inserts exit_sketch where needed.
    """
    if not actions:
        return actions
    
    tools = [a.get("tool", "") for a in actions]
    
    # Only apply if this is a loft workflow (has create_reference_plane + select_sketch + loft)
    has_ref_plane = "create_reference_plane" in tools
    has_loft = "loft" in tools
    
    if not (has_ref_plane and has_loft):
        return actions
    
    # Scan through and insert exit_sketch before create_reference_plane
    # when a sketch is still active (after validate_closed_profile without extrude)
    fixed = []
    sketch_active = False
    insertions = 0
    
    for a in actions:
        tool = a.get("tool", "")
        
        if tool in ("create_sketch", "create_sketch_on_selected_face"):
            sketch_active = True
        elif tool in ("extrude", "extrude_midplane", "cut_extrude", "cut_through_all", 
                       "revolve", "revolve_simple", "exit_sketch"):
            sketch_active = False
        
        # If sketch is active and we're about to create a reference plane, insert exit_sketch
        if sketch_active and tool == "create_reference_plane":
            fixed.append({"tool": "exit_sketch", "args": {}})
            sketch_active = False
            insertions += 1
        
        # If sketch is active and we're about to select_sketch (for loft), insert exit_sketch
        if sketch_active and tool == "select_sketch":
            fixed.append({"tool": "exit_sketch", "args": {}})
            sketch_active = False
            insertions += 1
        
        fixed.append(a)
    
    if insertions > 0:
        print(f"\n🔧 LOFT WORKFLOW POST-PROCESSING:")
        print(f"   Inserted {insertions} exit_sketch calls before create_reference_plane/select_sketch")
        print(f"   Steps: {len(actions)} → {len(fixed)}")
    
    return fixed


def _fix_gear_proportions(actions, user_request):
    """
    Detects gear patterns (circle + rectangle + circular_pattern)
    and recalculates dimensions using proper gear module formulas.
    """
    # Check if this looks like a gear request
    tools = [a.get("tool", "") for a in actions]
    has_circle = "draw_circle" in tools
    has_rect = "draw_rectangle" in tools
    has_pattern = "circular_pattern" in tools
    
    if not (has_circle and has_rect and has_pattern):
        return actions
    
    # Extract gear parameters from the user request
    request_lower = user_request.lower()
    if not any(kw in request_lower for kw in ["gear", "sprocket", "cog", "teeth"]):
        return actions
    
    # Parse tooth count from circular_pattern or from prompt
    tooth_count = None
    pattern_idx = None
    for i, a in enumerate(actions):
        if a.get("tool") == "circular_pattern":
            tooth_count = a.get("args", {}).get("count")
            pattern_idx = i
            break
    
    # Also try to parse from user request
    count_match = re.search(r'(\d+)\s*(?:teeth|tooth)', request_lower)
    if count_match and tooth_count is None:
        tooth_count = int(count_match.group(1))
    
    if tooth_count is None or tooth_count < 4:
        return actions
    
    # Parse OD from user request
    od = None
    od_patterns = [
        r'(?:outer\s*)?(?:diameter|od)\s*(?:of\s*)?(\d+(?:\.\d+)?)\s*(?:mm)?',
        r'(\d+(?:\.\d+)?)\s*mm\s*(?:outer\s*)?(?:diameter|od)',
        r'(\d+(?:\.\d+)?)\s*mm\s*(?:dia|diam)',
    ]
    for pat in od_patterns:
        m = re.search(pat, request_lower)
        if m:
            od = float(m.group(1))
            break
    
    # Fallback: infer OD from draw_circle radius
    if od is None:
        for a in actions:
            if a.get("tool") == "draw_circle":
                r = a.get("args", {}).get("radius")
                if r and not a.get("args", {}).get("x") and not a.get("args", {}).get("y"):
                    od = r * 2  # Base circle radius × 2
                    break
    
    if od is None:
        return actions
    
    # ═══════════════════════════════════════════════════════════
    # CORRECT GEAR GEOMETRY:
    #
    #   Pitch circle radius = m * N / 2
    #   Addendum (tip above pitch) = 1.0 * m
    #   Dedendum (root below pitch) = 1.25 * m
    #   Tip radius (OD/2) = pitch_r + addendum
    #   Root radius = pitch_r - dedendum
    #
    #   For the TOOTH RECTANGLE:
    #     - width (radial) = addendum + dedendum = 2.25 * m
    #     - The inner edge must OVERLAP into the base by ~1mm for SolidWorks merge
    #     - x_center = root_radius + width/2 - overlap
    #     
    #   This ensures teeth are VISIBLE and proportional!
    # ═══════════════════════════════════════════════════════════
    
    m_module = od / (tooth_count + 2)
    pitch_radius = m_module * tooth_count / 2
    addendum = m_module           # tooth above pitch circle
    dedendum = 1.25 * m_module    # tooth below pitch circle  
    tip_radius = od / 2           # = pitch + addendum
    root_radius = round(pitch_radius - dedendum, 1)
    
    # Tooth dimensions
    tooth_radial_height = round(addendum + dedendum, 1)  # Total tooth height
    tooth_tangential = round(m_module * 1.5, 1)          # Tooth thickness (~1.5*module)
    
    # CRITICAL: Position the rectangle so inner edge overlaps base by 1mm for merge
    overlap = 1.0  # mm overlap into base disk for SolidWorks body merge
    tooth_x = round(root_radius + tooth_radial_height / 2 - overlap, 1)
    
    # Ensure minimum dimensions
    tooth_radial_height = max(tooth_radial_height, 3.0)
    tooth_tangential = max(tooth_tangential, 2.0)
    root_radius = max(root_radius, 5.0)
    
    print(f"\n🔧 GEAR POST-PROCESSING:")
    print(f"   OD={od}mm, {tooth_count} teeth, module={m_module:.2f}mm")
    print(f"   Pitch radius = {pitch_radius:.1f}mm")
    print(f"   Root radius = {root_radius}mm (base disk circle)")
    print(f"   Tip radius = {tip_radius:.1f}mm (tooth tips reach here = OD/2)")
    print(f"   Tooth: {tooth_radial_height}mm radial × {tooth_tangential}mm tangential")
    print(f"   Tooth center X = {tooth_x}mm (overlap={overlap}mm into base)")
    
    # Find and fix the base circle
    for a in actions:
        if a.get("tool") == "draw_circle":
            args = a.get("args", {})
            # Only fix the base/center circle (not offset circles)
            if not args.get("x") and not args.get("y"):
                old_r = args.get("radius")
                if old_r and abs(old_r - root_radius) > 0.5:
                    print(f"   ✏️ Fixed base circle: radius {old_r} → {root_radius}")
                    args["radius"] = root_radius
                break
    
    # Find and replace the tooth rectangle with a TRAPEZOID for realistic tapered teeth
    # Trapezoid: wider at base (root), narrower at tip — like real involute gear teeth
    tooth_base_width = round(tooth_tangential * 1.5, 1)  # Base is wider
    tooth_tip_width = round(tooth_tangential * 0.7, 1)   # Tip is narrower
    
    rect_fixed = False
    for i, a in enumerate(actions):
        if a.get("tool") == "draw_rectangle":
            args = a.get("args", {})
            x = args.get("x")
            if x is not None and x > 0:  # Tooth rectangle is off-center
                old_tool = "draw_rectangle"
                
                # Replace with trapezoid
                actions[i] = {
                    "tool": "draw_trapezoid",
                    "args": {
                        "base_width": tooth_base_width,
                        "tip_width": tooth_tip_width,
                        "height": tooth_radial_height,
                        "x": tooth_x,
                        "y": 0
                    }
                }
                
                print(f"   ✏️ Replaced {old_tool} → draw_trapezoid")
                print(f"      base={tooth_base_width}mm, tip={tooth_tip_width}mm, height={tooth_radial_height}mm, x={tooth_x}mm")
                rect_fixed = True
                break
    
    if not rect_fixed:
        print(f"   ⚠️  Could not find tooth rectangle to replace")
    
    return actions


def _fix_hole_boundaries(actions):
    """
    Checks if cut circles exceed the plate boundary and rescales positions.
    Works by finding the plate rectangle dimensions, then clamping all
    draw_circle positions to stay within boundary - margin.
    """
    # Find the plate dimensions from draw_rectangle
    plate_width = None
    plate_depth = None
    found_cut = False
    
    for a in actions:
        if a.get("tool") == "draw_rectangle":
            args = a.get("args", {})
            if not args.get("x") and not args.get("y"):  # Centered rectangle = plate
                plate_width = args.get("width")
                plate_depth = args.get("height")
        if a.get("tool") in ("cut", "cut_through_all"):
            found_cut = True
    
    if plate_width is None or plate_depth is None:
        return actions
    
    # Find all draw_circle calls that come after the plate (hole cuts)
    # and check if any exceed boundaries
    half_w = plate_width / 2
    half_d = plate_depth / 2
    violations = 0
    
    # Collect all circles with their radii and positions
    circles = []
    past_plate = False
    for i, a in enumerate(actions):
        if a.get("tool") == "draw_rectangle" and not a.get("args", {}).get("x"):
            past_plate = True
            continue
        if past_plate and a.get("tool") == "draw_circle":
            args = a.get("args", {})
            x = args.get("x", 0)
            y = args.get("y", 0)
            r = args.get("radius", 0)
            if x != 0 or y != 0:  # Off-center circle = hole
                circles.append((i, x, y, r))
    
    if not circles:
        return actions
    
    # Check each circle for boundary violations
    margin = 3  # 3mm margin from edge
    for idx, x, y, r in circles:
        max_x = half_w - r - margin
        max_y = half_d - r - margin
        
        if max_x < 0 or max_y < 0:
            # Holes too big for plate — reduce radius
            new_r = min(half_w, half_d) / 4
            actions[idx]["args"]["radius"] = round(new_r, 1)
            r = new_r
            max_x = half_w - r - margin
            max_y = half_d - r - margin
            violations += 1
        
        if abs(x) > max_x or abs(y) > max_y:
            # Clamp positions
            new_x = max(-max_x, min(max_x, x))
            new_y = max(-max_y, min(max_y, y))
            actions[idx]["args"]["x"] = round(new_x, 1)
            actions[idx]["args"]["y"] = round(new_y, 1)
            violations += 1
    
    if violations > 0:
        print(f"\n🔧 HOLE BOUNDARY POST-PROCESSING:")
        print(f"   Plate: {plate_width}×{plate_depth}mm, boundary: ±{half_w-margin:.1f} × ±{half_d-margin:.1f}")
        print(f"   Fixed {violations} circles that would have exceeded plate edges")
        
        # Verify and print corrected positions
        for idx, x, y, r in circles:
            args = actions[idx]["args"]
            print(f"   • Circle at ({args.get('x',0)}, {args.get('y',0)}) r={args.get('radius',0)}")
    
    return actions


# --- UTILITIES ---

def clean_and_validate_json(raw_text):
    """Cleans markdown from LLM output and parses JSON array."""
    try:
        clean_text = raw_text.replace("```json", "").replace("```", "").strip()

        start_idx = clean_text.find('[')
        end_idx = clean_text.rfind(']')

        if start_idx == -1 or end_idx == -1:
            raise ValueError("No JSON array found in response")

        json_text = clean_text[start_idx:end_idx + 1]
        data = json.loads(json_text)

        if not isinstance(data, list):
            raise ValueError("Output is not a list of actions")

        return data

    except json.JSONDecodeError:
        print(f"❌ Failed to parse JSON. Raw output:\n{raw_text}")
        return None


def save_mission(data):
    """Saves the JSON mission file for the dispatcher."""
    if not data:
        return

    with open(OUTPUT_FILE, "w") as f:
        json.dump(data, f, indent=4)
    print(f"✅ Success! Mission saved to '{OUTPUT_FILE}'")
    print(f"   (Contains {len(data)} steps for SolidWorks)")


# --- MAIN EXECUTION ---
if __name__ == "__main__":
    try:
        from dotenv import load_dotenv
        load_dotenv()
    except ImportError:
        pass

    # Re-read after loading .env
    LLM_API_KEY  = os.environ.get("LLM_API_KEY", "")
    LLM_MODEL    = os.environ.get("LLM_MODEL", "llama-3.3-70b-versatile")
    LLM_BASE_URL = os.environ.get("LLM_BASE_URL", "https://api.groq.com/openai/v1")

    if not LLM_API_KEY:
        print("❌ LLM_API_KEY not set! Add it to your .env file.")
        exit()

    print(f"--- SOLIDWORKS AI AGENT ---")
    print(f"    Model: {LLM_MODEL}")
    print(f"    Base URL: {LLM_BASE_URL}")
    user_input = input("Enter design request: ")

    actions = get_agent_response(user_input)
    save_mission(actions)
