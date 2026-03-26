"""
SolidWorks AI Agent - Main Entry Point

Usage:
    python main.py

This script provides a complete workflow with three modes:
1. CREATE  - Generate a new 3D model from a text prompt
2. MODIFY  - Modify an existing open model
3. VALIDATE - Validate the current model against a prompt
"""

import os
import sys
# Load environment variables from .env file
try:
    from dotenv import load_dotenv
    load_dotenv(override=True)
except ImportError:
    pass  # python-dotenv not installed, will use system env vars

from llm_client import get_agent_response, get_modification_response, save_mission
from mcp_server.dispatcher import run_mission

# ============================================================
# CONFIGURATION
# ============================================================

MISSION_FILE = "mission.json"


def check_prerequisites():
    """Verify environment is set up correctly."""
    if not os.environ.get("LLM_API_KEY"):
        print("=" * 50)
        print("❌ LLM_API_KEY environment variable not set!")
        print("=" * 50)
        print("\nAdd these to your .env file:")
        print('  LLM_API_KEY=your_key_here')
        print('  LLM_MODEL=llama-3.3-70b-versatile')
        print('  LLM_BASE_URL=https://api.groq.com/openai/v1')
        return False
    
    return True


def check_solidworks():
    """Check if SolidWorks is running."""
    try:
        from tools.solidworks_app import get_sw_app
        sw = get_sw_app()
        if sw is None:
            print("⚠️  SolidWorks is not running.")
            print("   Please open SolidWorks before executing the mission.")
            return False
        print("✅ SolidWorks is connected.")
        return True
    except Exception as e:
        print(f"⚠️  Could not connect to SolidWorks: {e}")
        return False


# ============================================================
# MODE 1: CREATE NEW MODEL
# ============================================================

def mode_create():
    """Create a new 3D model from a text prompt."""
    print("\nDescribe the 3D model you want to create:")
    print("(Example: 'Create a 100x50mm rectangular plate, 10mm thick')")
    print()
    user_request = input("Your request: ").strip()
    
    if not user_request:
        print("❌ No request provided.")
        return
    
    # Step 1: Generate mission from AI
    print("\n" + "-" * 40)
    print("STEP 1: Generating CAD commands with AI...")
    print("-" * 40)
    
    actions = get_agent_response(user_request)
    
    if not actions:
        print("❌ Failed to generate mission. Please try again.")
        return
    
    save_mission(actions)
    
    # Step 2: Execute in SolidWorks
    print("\n" + "-" * 40)
    print("STEP 2: Ready to execute in SolidWorks")
    print("-" * 40)
    
    sw_ready = check_solidworks()
    
    if not sw_ready:
        print("\n📁 Mission saved to 'mission.json'")
        print("   Run 'python mcp_server/dispatcher.py' after opening SolidWorks.")
        return
    
    print("\nExecute in SolidWorks now? (y/n): ", end="")
    confirm = input().strip().lower()
    
    if confirm in ['y', 'yes']:
        print("\n" + "-" * 40)
        print("STEP 3: Executing in SolidWorks...")
        print("-" * 40)
        success = run_mission(MISSION_FILE)
        
        # Step 3: Auto-validate after execution
        if success:
            print("\n" + "-" * 40)
            print("STEP 4: Auto-validating the result...")
            print("-" * 40)
            # Allow SolidWorks to finish processing geometry
            import time
            time.sleep(1)
            try:
                from validator import run_validation
                run_validation(user_request)
            except Exception as e:
                print(f"⚠️  Validation skipped: {e}")
    else:
        print("\n📁 Mission saved to 'mission.json'")
        print("   Run 'python mcp_server/dispatcher.py' when ready.")


# ============================================================
# MODE 2: MODIFY EXISTING MODEL
# ============================================================

def mode_modify():
    """Modify the currently open SolidWorks model."""
    sw_ready = check_solidworks()
    if not sw_ready:
        print("❌ SolidWorks must be running with a model open.")
        return
    
    # Inspect current model
    print("\n🔍 Inspecting current model...")
    try:
        from tools.model_inspector import get_model_summary
        summary = get_model_summary()
        if summary is None:
            print("❌ Could not inspect model. Make sure a part with solid bodies is open.")
            return
        print(summary)
    except Exception as e:
        print(f"❌ Failed to inspect model: {e}")
        return
    
    # Get modification request
    print("\nDescribe the modification you want to make:")
    print("(Example: 'Add a 5mm fillet to all top edges')")
    print()
    modification = input("Modification: ").strip()
    
    if not modification:
        print("❌ No modification provided.")
        return
    
    # Generate delta steps
    print("\n" + "-" * 40)
    print("Generating modification steps...")
    print("-" * 40)
    
    actions = get_modification_response(modification, summary)
    
    if not actions:
        print("❌ Failed to generate modification steps.")
        return
    
    save_mission(actions)
    print(f"   Generated {len(actions)} modification steps")
    
    # Execute
    print("\nApply modification now? (y/n): ", end="")
    confirm = input().strip().lower()
    
    if confirm in ['y', 'yes']:
        print("\n" + "-" * 40)
        print("Applying modification...")
        print("-" * 40)
        run_mission(MISSION_FILE)
    else:
        print("\n📁 Modification steps saved to 'mission.json'")


# ============================================================
# MODE 3: VALIDATE CURRENT MODEL
# ============================================================

def mode_validate():
    """Validate the currently open model against a design prompt."""
    sw_ready = check_solidworks()
    if not sw_ready:
        print("❌ SolidWorks must be running with a model open.")
        return
    
    print("\nEnter the original design prompt to validate against:")
    print("(Example: 'Create a 100x100mm box, 50mm tall')")
    print()
    prompt = input("Prompt: ").strip()
    
    if not prompt:
        print("❌ No prompt provided.")
        return
    
    from validator import run_validation
    result = run_validation(prompt)
    
    if result:
        score = result["score"]
        if score >= 80:
            print(f"\n🎉 Great result! Score: {score}/100")
        elif score >= 60:
            print(f"\n⚠️  Acceptable result. Score: {score}/100")
        else:
            print(f"\n❌ Poor result. Score: {score}/100")
            print("   Consider regenerating or modifying the design.")


# ============================================================
# MAIN ENTRY POINT
# ============================================================

def main():
    """Main entry point for the AI CAD agent."""
    model_name = os.environ.get("LLM_MODEL", "unknown")
    
    print("=" * 60)
    print("🔧 SOLIDWORKS AI AGENT")
    print(f"   Model: {model_name}")
    print("=" * 60)
    
    # Check prerequisites
    if not check_prerequisites():
        return
    
    # Continuous session loop
    while True:
        print("\n" + "-" * 40)
        print("Choose a mode:")
        print("  [1] 🆕  Create new model")
        print("  [2] ✏️   Modify existing model")
        print("  [3] ✅  Validate current model")
        print("  [q] 🚪  Quit")
        print()
        
        choice = input("Enter choice (1/2/3/q): ").strip().lower()
        
        if choice == "1":
            mode_create()
        elif choice == "2":
            mode_modify()
        elif choice == "3":
            mode_validate()
        elif choice in ["q", "quit", "exit"]:
            print("\n👋 Goodbye!")
            break
        else:
            print(f"❌ Invalid choice: '{choice}'. Please enter 1, 2, 3, or q.")


if __name__ == "__main__":
    main()
