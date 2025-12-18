import win32com.client
import json
import os
import sys
from executor import execute

def main():
    # 1. Connect to SolidWorks
    try:
        app = win32com.client.Dispatch("SldWorks.Application")
        app.Visible = True
    except:
        print("Error: SolidWorks is not open. Please launch it first.")
        return

    # 2. Get or Create Active Document
    model = app.ActiveDoc

    # Check if we need to open a new file (Type 1 is Part)
    if model is None or (model.GetType != 1):
        print("No active Part found. Creating new...")

        # Primary Method: Use your specific template path
        template_path = r"C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2025\templates\Part.prtdot"

        # Fallback Method: Auto-detect from settings if primary path is missing
        if not os.path.exists(template_path):
            template_path = app.GetUserPreferenceStringValue(7)

        # Create the new document
        app.NewDocument(template_path, 0, 0, 0)
        
        # Update the model variable to the new file
        model = app.ActiveDoc

        if model is None:
            print("CRITICAL ERROR: Could not create a new Part.")
            return

    # 3. Load Data
    json_path = os.path.join("swtest","3D-Modelling-Project", "data", "sphere_001.json")
    
    if not os.path.exists(json_path):
        print(f"Error: JSON file not found at {json_path}")
        return

    with open(json_path, "r") as f:
        design_data = json.load(f)
    print(f"Loaded instructions from {json_path}")

    # 4. Execute Logic
    if isinstance(design_data, list):
        for part in design_data:
            execute(model, part)
    else:
        execute(model, design_data)

    print("Execution finished.")

if __name__ == "__main__":
    main()