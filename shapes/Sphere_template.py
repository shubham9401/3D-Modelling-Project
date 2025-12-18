import win32com.client
import pythoncom
import os
import math

pythoncom.CoInitialize()

class Sphere:
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

    # -------- Sphere creation --------
    def create(self, diameter_mm):
        radius = (diameter_mm / 2) / 1000.0  # mm → meters

        sk = self.model.SketchManager
        fm = self.model.FeatureManager

        # Create centerline for revolution axis (vertical through origin)
        axis_seg = sk.CreateCenterLine(0, -radius * 1.5, 0, 0, radius * 1.5, 0)

        # Create a semicircle to the right of the axis (start at top, end at bottom)
        sk.CreateArc(0, 0, 0, 0, radius, 0, 0, -radius, 0, 1)

        print(f"Created semicircle profile with radius: {radius} m (diameter={diameter_mm} mm)")

        # Select the centerline as revolution axis
        self.model.ClearSelection2(True)
        try:
            axis_seg.Select4(False, None)
        except Exception:
            # Fallback: attempt name-based selection if needed
            self.model.Extension.SelectByID2("Line1", "SKETCHSEGMENT", 0, 0, 0, False, 0, self.nothing, 0)

        # Revolve 360 degrees to create the sphere
        fm.FeatureRevolve2(
            True, True, False, False, False, False,
            0, 0,
            2 * math.pi, 0,
            False, False,
            0.0, 0.0,
            0, 0, 0,
            True, True, True
        )

        print(f"Revolved profile to create sphere with diameter: {diameter_mm} mm")
class SphereBuilder:
    def build(self, model, data):
        # 1. Initialize logic
        sphere = Sphere(model)
        # 2. Plane selection
        sphere.Plane(data["plane"])
        # 3. Create sphere
        sphere.create(data["diameter_mm"])
 
