import win32com.client
import pythoncom
import os
import math

pythoncom.CoInitialize()

class Part:
    def __init__(self):
        self.app = win32com.client.Dispatch("SldWorks.Application")
        self.app.Visible = True
        print("Connected to SolidWorks and set visible.")

        self.template = r"C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2021\templates\Part.prtdot"
        if os.path.exists(self.template):
            self.app.NewDocument(self.template, 0, 0, 0)
            print("Opened new Part using template:", self.template)
        else:
            try:
                self.app.NewPart()
                print("Opened new Part using default template.")
            except Exception as e:
                raise Exception("Unable to create new Part. Check SolidWorks templates.") from e
        self.model = self.app.ActiveDoc
        if self.model is None:
            raise Exception("ActiveDoc is None after creating Part.")
        print("Active document ready.")

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

    # -------- Cone creation --------
    def cone(self, base_diameter_mm, height_mm):
        base_radius = (base_diameter_mm / 2) / 1000.0  # mm → meters
        height = height_mm / 1000.0

        sk = self.model.SketchManager
        fm = self.model.FeatureManager

        # Create centerline for revolution axis (vertical through origin)
        axis_seg = sk.CreateCenterLine(0, -height * 0.2, 0, 0, height * 1.2, 0)

        # Create triangle for cone profile (right triangle entirely on right of axis)
        sk.CreateLine(0, 0, 0, base_radius, 0, 0)           # Base (from axis to right)
        sk.CreateLine(base_radius, 0, 0, 0, height, 0)      # Slant up to apex on axis
        sk.CreateLine(0, height, 0, 0, 0, 0)                # Close back down along axis
        
        print(f"Created cone profile: base radius={base_radius} m, height={height} m")

        # Select the centerline as revolution axis
        self.model.ClearSelection2(True)
        try:
            axis_seg.Select4(False, None)
        except Exception:
            # Fallback to name-based selection (first line usually the centerline)
            self.model.Extension.SelectByID2("Line1", "SKETCHSEGMENT", 0, 0, 0, False, 0, self.nothing, 0)

        # Revolve the profile 360 degrees to create solid cone
        fm.FeatureRevolve2(
            True, True, False, False, False, False,
            0, 0,
            2 * math.pi, 0,
            False, False,
            0.0, 0.0,
            0, 0, 0,
            True, True, True
        )
        
        print(f"Revolved profile to create cone: base diameter={base_diameter_mm} mm, height={height_mm} mm")

# Create and use Part
if __name__ == "__main__":
    part = Part()
    part.Plane("Front")  # Select Front plane for revolution
    part.cone(40, 60)    # base diameter=40mm, height=60mm
    print("Finished creating cone (base diameter=40mm, height=60mm).")
