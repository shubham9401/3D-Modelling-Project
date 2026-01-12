import win32com.client
import json
import os
import sys
import argparse
from executor import execute
from command_parser import parse_prompt, parse_kv_params

TEMPLATE_PATH = r"C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2021\templates\Part.prtdot"

def save_stl(model, out_path):
    """
    Save the active model as STL.
    For finer control, retrieve ExportData (GetExportFileData(1)) and set Binary/Deviation/Angle.
    """
    ok = model.Extension.SaveAs(out_path, 0, None, None, None, None)
    if not ok:
        if not model.SaveAs3(out_path, 0, 0):
            raise RuntimeError(f"Failed to save STL: {out_path}")

def new_part(app):
    template_path = TEMPLATE_PATH if os.path.exists(TEMPLATE_PATH) else app.GetUserPreferenceStringValue(7)
    app.NewDocument(template_path, 0, 0, 0)
    return app.ActiveDoc

def build_once(app, model, design_data, out_dir="out", name="model"):
    os.makedirs(out_dir, exist_ok=True)
    execute(model, design_data)
    stl_out = os.path.join(out_dir, f"{name}.stl")
    save_stl(model, stl_out)
    print(f"Built '{design_data['shape']}' -> {stl_out}")
    return stl_out

def main():
    parser = argparse.ArgumentParser(description="SolidWorks COM generator")
    parser.add_argument("--prompt", type=str, help="Free text like: 'make a cube 40mm on top plane'")
    parser.add_argument("--shape", type=str, help="Explicit shape name (cube, cuboid, cylinder, sphere, cone, wedge, cup)")
    parser.add_argument("--params", type=str, help="Override params as key=value pairs separated by commas (e.g., edge_mm=50,plane=Top)")
    parser.add_argument("--out-dir", default="out", help="Output folder for STL when using prompt/shape modes")
    parser.add_argument("--name", default=None, help="Base output name")
    args = parser.parse_args()

    # 1. Connect to SolidWorks
    try:
        app = win32com.client.Dispatch("SldWorks.Application")
        app.Visible = True
    except:
        print("Error: SolidWorks is not open. Please launch it first.")
        return

    # 2. Get or Create Active Document (Type 1 is Part)
    model = app.ActiveDoc
    if model is None or (model.GetType() != 1):
        print("No active Part found. Creating new...")
        model = new_part(app)
        if model is None:
            print("CRITICAL ERROR: Could not create a new Part.")
            return

    # Prompt or explicit shape mode
    if args.prompt or args.shape:
        # Parse design_data either from prompt or explicit shape + params
        if args.prompt:
            design_data = parse_prompt(args.prompt)
        else:
            # shape + optional params
            design_data = {"shape": args.shape.lower()}
            if args.params:
                design_data.update(parse_kv_params(args.params))

        # Build once
        base_name = args.name or design_data["shape"]
        build_once(app, model, design_data, out_dir=args.out_dir, name=base_name)
        return

    # 3. JSON batch mode (existing behavior)
    data_folder = os.path.join("data")
    if not os.path.isdir(data_folder):
        print(f"Data folder '{data_folder}' not found.")
        return

    all_files = [f for f in os.listdir(data_folder) if f.endswith('.json')]
    for filename in all_files:
        json_path = os.path.join(data_folder, filename)
        try:
            with open(json_path, "r") as f:
                design_data = json.load(f)
            # Build the part
            execute(model, design_data)
            # Export STL next to the JSON
            stem = os.path.splitext(filename)[0]
            stl_out = os.path.join(data_folder, f"{stem}.stl")
            save_stl(model, stl_out)
            print(f"{filename} processed successfully -> {stl_out}")
            # Fresh part for next file
            model = new_part(app)
        except Exception as e:
            print(f"Failed to process {filename}: {e}")

    print("All files processed successfully.")        

if __name__ == "__main__":
    main()