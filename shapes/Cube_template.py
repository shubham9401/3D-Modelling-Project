import win32com.client
import pythoncom
import os

pythoncom.CoInitialize()

class Cube:
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

    # -------- Cube creation --------
    def create(self, edge_mm):
        edge = edge_mm / 1000.0  # mm → meters
        half = edge / 2

        sk = self.model.SketchManager
        fm = self.model.FeatureManager

        # Center rectangle
        sk.CreateCenterRectangle(0, 0, 0, half, half, 0)
        print(f"Created center rectangle with half-edge: {half} m")

        # Extrude
        fm.FeatureExtrusion2(
            True, False, False,
            0, 0,
            edge, 0,
            False, False, False, False,
            0, 0,
            False, False, False, False,
            True, True, True,
            0, 0, False
        )
        print(f"Extruded cube to depth: {edge} m (={edge_mm} mm)")


class CubeBuilder:
    def build(self, model, data):
        # 1. Initialize logic
        cube = Cube(model)
        # 2. Plane selection
        cube.Plane(data["plane"])
        # 3. Create cube
        cube.create(data["edge_mm"])