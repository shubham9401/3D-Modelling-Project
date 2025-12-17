
import win32com.client
import pythoncom
import os

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
            # Fallback to default part template if specific path is missing
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

    # -------- Cube creation --------
    def cube(self, edge_mm):
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

# Create and use Part
if __name__ == "__main__":
    part = Part()
    part.Plane("Top")   # Select Top plane
    part.cube(5)        # 3 mm × 3 mm × 3 mm cube
    print("Finished creating 3mm cube.")