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


# --- MAIN FUNCTIONS ---

def get_agent_response(user_request):
    """
    Sends user request to the LLM and returns CAD commands as JSON.
    """
    full_system_message = f"{SYSTEM_INSTRUCTION}\n\nAVAILABLE TOOLS:\n{AVAILABLE_TOOLS}"

    print(f"🧠 Processing: '{user_request}' (model: {LLM_MODEL})...")

    try:
        content = _call_llm(full_system_message, user_request)
        return clean_and_validate_json(content)
    except Exception as e:
        print(f"❌ API Error: {e}")
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
