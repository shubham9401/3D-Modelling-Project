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
        template_path = r"C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2021\templates\Part.prtdot"

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
    data_folder = os.path.join("data")
    
    # This list comprehension creates a list of every .json file in that folder
    all_files = [f for f in os.listdir(data_folder) if f.endswith('.json')]
    for filename in all_files:
        json_path = os.path.join(data_folder, filename)
        
        try:
            with open(json_path, "r") as f:
                design_data = json.load(f)

            # ... Calls the executor ...
            execute(model, design_data)
            print("{filename} processed successfully")
            #To make sure different json files are made in different part 
            app.NewDocument(template_path, 0, 0, 0)
            model = app.ActiveDoc
            
                
        except Exception as e:
            print(f"Failed to process {filename}: {e}")

    print("All files processed successfully.")        

if __name__ == "__main__":
    main()