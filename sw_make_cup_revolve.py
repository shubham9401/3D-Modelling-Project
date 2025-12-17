import win32com.client
import re
import math
import time
import os  # <--- Added this to handle file paths properly

print("--- SOLIDWORKS AI ASSISTANT V2 (Auto-Save Enabled) ---")
print("Supported Commands:")
print("  1. Cylinder (Radius, Height)")
print("  2. Cone     (Radius, Height)")
print("  3. Cube     (Side)")
print("  4. Cuboid   (Length, Width, Height)")
print("  5. Pyramid  (Base_Side, Height)")
print("----------------------------------")

def get_solidworks():
    try:
        swApp = win32com.client.GetActiveObject("SldWorks.Application")
        model = swApp.ActiveDoc
        return swApp, model
    except:
        return None, None

def extract_numbers(text):
    """ Finds numbers (integers or decimals) in text """
    numbers = re.findall(r"[-+]?\d*\.\d+|\d+", text)
    return [float(n) for n in numbers]

# --- NEW HELPER FUNCTION FOR SAVING ---
def save_dynamic(model, shape_type, dims):
    """ 
    Saves the part with a name based on its dimensions.
    Example: Cylinder_R50_H100.SLDPRT 
    """
    # 1. DEFINE YOUR FOLDER HERE (Check this path!)
    save_folder = r"D:\3D Model Project"
    
    # 2. Generate a unique filename string
    # dims is a list of numbers, e.g., [50, 100]
    # This creates a string like "_50_100"
    dim_str = "_".join([str(int(d)) for d in dims])
    
    filename = f"{shape_type}_{dim_str}.SLDPRT"
    full_path = os.path.join(save_folder, filename)
    
    print(f"   -> Saving to: {filename}")
    
    # 3. Save command
    try:
        model.SaveAs(full_path)
        print("   -> [Saved Successfully]")
    except Exception as e:
        print(f"   -> [Save Failed]: {e}")

# --- SHAPE FUNCTIONS (UPDATED) ---

def make_cylinder(model, r, h):
    print(f" -> Building Cylinder: R={r}, H={h}")
    model.SketchManager.InsertSketch(True)
    model.SketchManager.CreateCircle(0, 0, 0, r/1000.0, 0, 0)
    model.FeatureManager.FeatureExtrusion2(True, False, False, 0, 0, h/1000.0, 0, False, False, False, False, 0, 0, False, False, False, False, True, True, True, 0, 0, False)
    
    # CALL THE SAVE FUNCTION
    save_dynamic(model, "Cylinder", [r, h])

def make_cone(model, r, h):
    print(f" -> Building Cone: R={r}, H={h}")
    angle_deg = math.degrees(math.atan(r / h))
    model.SketchManager.InsertSketch(True)
    model.SketchManager.CreateCircle(0, 0, 0, r/1000.0, 0, 0)
    model.FeatureManager.FeatureExtrusion2(True, False, False, 0, 0, h/1000.0, 0, True, False, False, False, angle_deg * (3.14159/180), 0, False, False, False, False, True, True, True, 0, 0, False)
    
    # SAVE
    save_dynamic(model, "Cone", [r, h])

def make_cube(model, side):
    print(f" -> Building Cube: Side={side}")
    model.SketchManager.InsertSketch(True)
    half = (side / 1000.0) / 2
    model.SketchManager.CreateCenterRectangle(0, 0, 0, half, half, 0)
    model.FeatureManager.FeatureExtrusion2(True, False, False, 0, 0, side/1000.0, 0, False, False, False, False, 0, 0, False, False, False, False, True, True, True, 0, 0, False)
    
    # SAVE
    save_dynamic(model, "Cube", [side])

def make_cuboid(model, length, width, height):
    print(f" -> Building Cuboid: L={length}, W={width}, H={height}")
    model.SketchManager.InsertSketch(True)
    half_l = (length / 1000.0) / 2
    half_w = (width / 1000.0) / 2
    model.SketchManager.CreateCenterRectangle(0, 0, 0, half_l, half_w, 0)
    model.FeatureManager.FeatureExtrusion2(True, False, False, 0, 0, height/1000.0, 0, False, False, False, False, 0, 0, False, False, False, False, True, True, True, 0, 0, False)
    
    # SAVE
    save_dynamic(model, "Cuboid", [length, width, height])

def make_pyramid(model, side, height):
    print(f" -> Building Pyramid: Base={side}, H={height}")
    half_side = side / 2.0
    angle_deg = math.degrees(math.atan(half_side / height))
    
    model.SketchManager.InsertSketch(True)
    half_m = (side / 1000.0) / 2
    model.SketchManager.CreateCenterRectangle(0, 0, 0, half_m, half_m, 0)
    
    model.FeatureManager.FeatureExtrusion2(
        True, False, False, 0, 0, height/1000.0, 0, 
        True, False, False, False, 
        angle_deg * (3.14159/180), 
        0, False, False, False, False, 
        True, True, True, 0, 0, False
    )
    
    # SAVE
    save_dynamic(model, "Pyramid", [side, height])

# --- MAIN LOOP ---
swApp, model = get_solidworks()

if not model:
    print("ERROR: Open a New Part in SolidWorks first.")
else:
    while True:
        user_input = input("\nAI: What should I build? (or 'exit'): ").lower()
        if "exit" in user_input: break
            
        nums = extract_numbers(user_input)
        
        # RESET PLANE
        feature = model.FirstFeature
        while feature:
            if "Front Plane" == feature.Name:
                feature.Select2(False, 0)
                break
            feature = feature.GetNextFeature
        
        # DECISION TREE
        if "pyramid" in user_input:
            if len(nums) >= 2: make_pyramid(model, nums[0], nums[1])
            else: print("Error: Pyramid needs Base and Height")
            
        elif "cuboid" in user_input or "brick" in user_input:
            if len(nums) >= 3: make_cuboid(model, nums[0], nums[1], nums[2])
            else: print("Error: Cuboid needs Length, Width, and Height")

        elif "cone" in user_input:
            if len(nums) >= 2: make_cone(model, nums[0], nums[1])
            
        elif "cylinder" in user_input:
            if len(nums) >= 2: make_cylinder(model, nums[0], nums[1])
            
        elif "cube" in user_input:
            if len(nums) >= 1: make_cube(model, nums[0])
            
        else:
            print("Unknown shape.")

        model.ViewZoomtofit2()
        
        