"""
SolidWorks AI Agent - Main Entry Point

Usage:
    python main.py

This script provides a complete workflow:
1. Takes a natural language prompt from the user
2. Sends it to Llama 3 (via Groq) to generate CAD commands
3. Executes those commands in SolidWorks
"""

import os
import sys
# Load environment variables from .env file
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass  # python-dotenv not installed, will use system env vars
from llm_client import get_agent_response, save_mission
from mcp_server.dispatcher import run_mission

# ============================================================
# CONFIGURATION
# ============================================================

MISSION_FILE = "mission.json"


def check_prerequisites():
    """Verify environment is set up correctly."""
    api_key = os.environ.get("GROQ_API_KEY")
    
    if not api_key:
        print("=" * 50)
        print("❌ GROQ_API_KEY environment variable not set!")
        print("=" * 50)
        print("\nTo fix this, run in PowerShell:")
        print('  $env:GROQ_API_KEY = "gsk_your_key_here"')
        print("\nGet your free key at: https://console.groq.com/")
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


def main():
    """Main entry point for the AI CAD agent."""
    print("=" * 60)
    print("🔧 SOLIDWORKS AI AGENT")
    print("   Powered by Llama 3 (via Groq)")
    print("=" * 60)
    
    # Check prerequisites
    if not check_prerequisites():
        return
    
    # Get user input
    print("\nDescribe the 3D model you want to create:")
    print("(Example: 'Create a 100x50mm rectangular plate, 10mm thick')")
    print()
    user_request = input("Your request: ").strip()
    
    if not user_request:
        print("❌ No request provided. Exiting.")
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
    
    # Step 2: Ask user to execute
    print("\n" + "-" * 40)
    print("STEP 2: Ready to execute in SolidWorks")
    print("-" * 40)
    
    # Check SolidWorks connection
    sw_ready = check_solidworks()
    
    if not sw_ready:
        print("\n📁 Mission saved to 'mission.json'")
        print("   Run 'python mcp_server/dispatcher.py' after opening SolidWorks.")
        return
    
    # Ask for confirmation
    print("\nExecute in SolidWorks now? (y/n): ", end="")
    confirm = input().strip().lower()
    
    if confirm in ['y', 'yes']:
        print("\n" + "-" * 40)
        print("STEP 3: Executing in SolidWorks...")
        print("-" * 40)
        run_mission(MISSION_FILE)
    else:
        print("\n📁 Mission saved to 'mission.json'")
        print("   Run 'python mcp_server/dispatcher.py' when ready.")


if __name__ == "__main__":
    main()
