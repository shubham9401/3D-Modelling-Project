import json
import os
from groq import Groq
from system_prompt import SYSTEM_INSTRUCTION, AVAILABLE_TOOLS

# --- CONFIGURATION ---

API_KEY = process.env.API_KEY   
OUTPUT_FILE = "mission.json"

def get_agent_response(user_request):
    """
    Connects to Llama 3 to generate the CAD commands.
    """
    client = Groq(api_key=API_KEY)
    
    # Combine the Role (Instruction) with the Tools and the User Request
    full_system_message = f"{SYSTEM_INSTRUCTION}\n\nAVAILABLE TOOLS:\n{AVAILABLE_TOOLS}"
    
    print(f"🧠 Processing: '{user_request}'...")

    try:
        completion = client.chat.completions.create(
            model="llama-3.3-70b-versatile", # High-intelligence model
            messages=[
                {"role": "system", "content": full_system_message},
                {"role": "user", "content": user_request}
            ],
            temperature=0.1, # Low temperature = Precise, non-creative code
            stop=None
        )
        
        # Extract the content
        content = completion.choices[0].message.content
        return clean_and_validate_json(content)

    except Exception as e:
        print(f"❌ API Error: {e}")
        return None

def clean_and_validate_json(raw_text):
    """
    Ensures the AI returned valid JSON, removing any markdown text.
    """
    try:
        # Remove markdown backticks if present
        clean_text = raw_text.replace("```json", "").replace("```", "").strip()
        
        # Parse JSON
        data = json.loads(clean_text)
        
        # Basic check: It must be a list of actions
        if not isinstance(data, list):
            raise ValueError("Output is not a list of actions")
            
        return data
        
    except json.JSONDecodeError:
        print(f"❌ Failed to parse JSON. Raw output:\n{raw_text}")
        return None

def save_mission(data):
    """Saves the JSON for the MCP Server."""
    if not data:
        return
        
    with open(OUTPUT_FILE, "w") as f:
        json.dump(data, f, indent=4)
    print(f"✅ Success! Mission saved to '{OUTPUT_FILE}'")
    print(f"   (Contains {len(data)} steps for SolidWorks)")

# --- MAIN EXECUTION ---
if __name__ == "__main__":
    # Check if key is set
    if "YOUR_GROQ_API_KEY" in API_KEY:
        print("❌ ERROR: You need to paste your Groq API Key in line 8.")
        exit()

    print("--- SOLIDWORKS AI AGENT (Powered by Llama 3) ---")
    user_input = input("Enter design request: ")
    
    # 1. Get real AI response
    actions = get_agent_response(user_input)
    
    # 2. Save file
    save_mission(actions)