from pyexpat import model
import win32com.client
import pythoncom
import os

pythoncom.CoInitialize()

class Cuboid:
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

    # -------- Cuboid creation --------
    def create(self, length_mm,breadth_mm,height_mm):
        length = length_mm / 1000.0  # mm → meters
        breadth = breadth_mm / 1000.0
        height = height_mm/1000.0
        half_l = length / 2
        half_b = breadth / 2
        

        sk = self.model.SketchManager
        fm = self.model.FeatureManager

        # Center rectangle
        sk.CreateCenterRectangle(0, 0, 0, half_l, half_b, 0)
        print(f"Created center rectangle with half-length , half-breadth respectively : {half_l} m , {half_b} m ")

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
        print(f"Extruded cube to depth: {height} m (={height_mm} mm)")

class CuboidBuilder:
    def build(self, model, data):
        # 1. Initialize logic
        cuboid = Cuboid(model)
        cuboid.Plane(data["plane"])
        cuboid.create(
            data["length_mm"],
            data["breadth_mm"],
            data["height_mm"]
        )
