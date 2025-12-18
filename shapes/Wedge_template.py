import win32com.client
import pythoncom
import os

pythoncom.CoInitialize()

class Wedge:
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

    # -------- Wedge creation (triangular prism) --------
    def create(self, base_mm, height_mm, depth_mm):
        base = base_mm / 1000.0  # mm → meters
        height = height_mm / 1000.0
        depth = depth_mm / 1000.0

        sk = self.model.SketchManager
        fm = self.model.FeatureManager

        # Create right triangle (wedge cross-section)
        # Points: origin (0,0), right along base, up at height
        sk.CreateLine(0, 0, 0, base, 0, 0)      # Base (horizontal)
        sk.CreateLine(base, 0, 0, 0, height, 0) # Hypotenuse (slant)
        sk.CreateLine(0, height, 0, 0, 0, 0)    # Vertical side
        
        print(f"Created triangle: base={base} m, height={height} m")

        # Extrude the triangle to create wedge
        fm.FeatureExtrusion2(
            True, False, False,
            0, 0,
            depth, 0,
            False, False, False, False,
            0, 0,
            False, False, False, False,
            True, True, True,
            0, 0, False
        )
        print(f"Extruded wedge to depth: {depth} m (={depth_mm} mm)")

class WedgeBuilder:
    def build(self, model, data):
        # 1. Initialize logic
        wedge = Wedge(model)
        # 2. Plane selection
        wedge.Plane(data["plane"])
        # 3. Create wedge
        wedge.create(data["base_mm"], data["height_mm"], data["depth_mm"])