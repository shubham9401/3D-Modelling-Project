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
                    return retry_actions
                else:
                    print(f"⚠️  Retry still incomplete. Using best result.")
                    return retry_actions if len(retry_actions) >= len(actions) else actions

        elif missing:
            missing_str = ", ".join(missing)
            print(f"⚠️  WARNING: Output may be incomplete. Missing: {missing_str}")

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
        return clean_and_validate_json(content)
    except Exception as e:
        print(f"❌ API Error: {e}")
        return None


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
