import win32com.client
import pythoncom
import os

pythoncom.CoInitialize()

class Cylinder:
    def __init__(self,model):
        self.model = model
        self.nothing = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)

    # -------- Plane selection --------
    def Plane(self, name):
        plane_map = {
            "Top": "Top Plane",
            "Front": "Front Plane",
            "Right": "Right Plane"
        }

        if name not in plane_map:
            raise ValueError("Plane must be Top, Front, or Right")

        self.model.Extension.SelectByID2(
            plane_map[name],
            "PLANE",
            0, 0, 0,
            False, 0,
            self.nothing, 0
        )

        self.model.InsertSketch2(True)
        print(f"Started sketch on: {plane_map[name]}")

    # -------- Cylinder creation --------
    def create(self, diameter_mm,height_mm):
        diameter = diameter_mm / 1000.0  # mm → meters
        radius =  diameter/ 2
        height = height_mm/1000.0

        sk = self.model.SketchManager
        fm = self.model.FeatureManager

        # Center circle
        sk.CreateCircle(0, 0, 0,radius , 0, 0)
        

        # Extrude
        fm.FeatureExtrusion2(
            True, False, False,
            0, 0,
            height, 0,
            False, False, False, False,
            0, 0,
            False, False, False, False,
            True, True, True,
            0, 0, False
        )
        

# Create and use Part
class CylinderBuilder:
    def build(self, model, data):
        # 1. Initialize logic
        cylinder = Cylinder(model)
        # 2. Plane selection
        cylinder.Plane(data["plane"])
        # 3. Create cylinder
        cylinder.create(data["diameter_mm"],data["height_mm"])