import win32com.client
import pythoncom
import math

pythoncom.CoInitialize()

class Cup:
    """
    Revolve-based cup with wall thickness and optional lip fillet.
    Parameters (mm): outer_d_mm, wall_mm, height_mm, lip_fillet_mm (optional)
    """
    def __init__(self, model):
        self.model = model
        self.nothing = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)

    def Plane(self, name):
        plane_map = {"Top": "Top Plane", "Front": "Front Plane", "Right": "Right Plane"}
        if name not in plane_map:
            raise ValueError("Plane must be Top, Front, or Right")
        self.model.Extension.SelectByID2(plane_map[name], "PLANE", 0, 0, 0, False, 0, self.nothing, 0)
        self.model.InsertSketch2(True)
        print(f"Started sketch on: {plane_map[name]}")

    def create(self, outer_d_mm, wall_mm, height_mm, lip_fillet_mm=None):
        R_outer = (outer_d_mm / 2.0) / 1000.0
        t = wall_mm / 1000.0
        H = height_mm / 1000.0
        r_inner = R_outer - t
        if r_inner <= 0:
            raise ValueError("wall thickness too large for given outer diameter")

        sk = self.model.SketchManager
        fm = self.model.FeatureManager

        axis_seg = sk.CreateCenterLine(0, -H * 0.5, 0, 0, H * 1.5, 0)

        sk.CreateLine(R_outer, 0, 0, R_outer, H, 0)

        if lip_fillet_mm and lip_fillet_mm > 0:
            rf = min(lip_fillet_mm / 1000.0, t * 0.9)
            sk.CreateArc(R_outer - rf, H - rf, 0, R_outer, H, 0, r_inner, H - 2 * rf, 0, 1)
            inner_top_y = H - 2 * rf
        else:
            inner_top_y = H

        sk.CreateLine(r_inner, inner_top_y, 0, r_inner, t, 0)
        sk.CreateLine(r_inner, t, 0, 0, t, 0)
        sk.CreateLine(0, t, 0, 0, 0, 0)
        sk.CreateLine(0, 0, 0, R_outer, 0, 0)

        print(f"Cup profile: R_outer={R_outer} m, thickness={t} m, height={H} m")

        self.model.ClearSelection2(True)
        try:
            axis_seg.Select4(False, None)
        except Exception:
            self.model.Extension.SelectByID2("Line1", "SKETCHSEGMENT", 0, 0, 0, False, 0, self.nothing, 0)

        fm.FeatureRevolve2(
            True, True, False, False, False, False,
            0, 0,
            2 * math.pi, 0,
            False, False,
            0.0, 0.0,
            0, 0, 0,
            True, True, True
        )
        print("Revolved profile to create cup.")

class CupBuilder:
    def build(self, model, data):
        cup = Cup(model)
        cup.Plane(data.get("plane", "Front"))
        cup.create(
            outer_d_mm=data["outer_d_mm"],
            wall_mm=data["wall_mm"],
            height_mm=data["height_mm"],
            lip_fillet_mm=data.get("lip_fillet_mm", None),
        )
