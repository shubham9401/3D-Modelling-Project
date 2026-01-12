import re
from typing import Optional, List, Dict

SHAPE_ALIASES = {
    "cube": ["cube", "box", "square prism"],
    "cuboid": ["cuboid", "rectangular prism", "rect box"],
    "cylinder": ["cylinder", "tube (solid)"],
    "sphere": ["sphere", "ball"],
    "cone": ["cone"],
    "wedge": ["wedge", "triangular prism"],
    "cup": ["cup", "tumbler"],
}

DEFAULTS = {
    "cube": {"plane": "Top", "edge_mm": 40},
    "cuboid": {"plane": "Top", "length_mm": 80, "breadth_mm": 50, "height_mm": 30},
    "cylinder": {"plane": "Top", "diameter_mm": 50, "height_mm": 60},
    "sphere": {"plane": "Front", "diameter_mm": 45},
    "cone": {"plane": "Front", "base_diameter_mm": 60, "height_mm": 80},
    "wedge": {"plane": "Front", "base_mm": 40, "height_mm": 30, "depth_mm": 60},
    "cup": {"plane": "Front", "outer_d_mm": 80, "wall_mm": 3, "height_mm": 100},
}

PLANES = {"top": "Top", "front": "Front", "right": "Right"}

def detect_shape(text: str) -> Optional[str]:
    t = text.lower()
    for shape, aliases in SHAPE_ALIASES.items():
        for a in aliases:
            if a in t:
                return shape
    return None

def parse_plane(text: str) -> Optional[str]:
    t = text.lower()
    for k, v in PLANES.items():
        if k in t:
            return v
    return None

def extract_mm_numbers(text: str) -> List[float]:
    nums = []
    for m in re.finditer(r"(\d+(?:\.\d+)?)\s*(mm)?", text.lower()):
        try:
            nums.append(float(m.group(1)))
        except (ValueError, AttributeError):
            pass
    return nums

def parse_prompt(prompt: str) -> Dict:
    shape = detect_shape(prompt)
    if not shape:
        raise ValueError("Could not detect shape in prompt. Try e.g., 'make a cube 40mm on top plane'.")

    data = {"shape": shape}
    plane = parse_plane(prompt)
    if plane:
        data["plane"] = plane

    nums = extract_mm_numbers(prompt)

    if shape == "cube":
        edge = None
        m = re.search(r"edge\s*(\d+(?:\.\d+)?)", prompt.lower())
        if m:
            edge = float(m.group(1))
        elif nums:
            edge = nums[0]
        data["edge_mm"] = edge if edge is not None else DEFAULTS["cube"]["edge_mm"]

    elif shape == "cuboid":
        def find(label):
            m = re.search(label + r"\s*(\d+(?:\.\d+)?)", prompt.lower())
            return float(m.group(1)) if m else None
        L = find("length") or (nums[0] if len(nums) >= 1 else DEFAULTS["cuboid"]["length_mm"])
        B = find("breadth") or find("width") or (nums[1] if len(nums) >= 2 else DEFAULTS["cuboid"]["breadth_mm"])
        H = find("height") or (nums[2] if len(nums) >= 3 else DEFAULTS["cuboid"]["height_mm"])
        data.update({"length_mm": L, "breadth_mm": B, "height_mm": H})

    elif shape == "cylinder":
        D = None
        H = None
        mD = re.search(r"(diameter|dia)\s*(\d+(?:\.\d+)?)", prompt.lower())
        mH = re.search(r"height\s*(\d+(?:\.\d+)?)", prompt.lower())
        if mD: D = float(mD.group(2))
        if mH: H = float(mH.group(1))
        if D is None and nums: D = nums[0]
        if H is None and len(nums) >= 2: H = nums[1]
        data.update({"diameter_mm": D or DEFAULTS["cylinder"]["diameter_mm"], "height_mm": H or DEFAULTS["cylinder"]["height_mm"]})

    elif shape == "sphere":
        D = None
        mD = re.search(r"(diameter|dia)\s*(\d+(?:\.\d+)?)", prompt.lower())
        if mD: D = float(mD.group(2))
        if D is None and nums: D = nums[0]
        data.update({"diameter_mm": D or DEFAULTS["sphere"]["diameter_mm"]})

    elif shape == "cone":
        BD = None
        H = None
        mBD = re.search(r"(base\s*diameter|base\s*dia|diameter)\s*(\d+(?:\.\d+)?)", prompt.lower())
        mH = re.search(r"height\s*(\d+(?:\.\d+)?)", prompt.lower())
        if mBD: BD = float(mBD.group(2))
        if mH: H = float(mH.group(1))
        if BD is None and nums: BD = nums[0]
        if H is None and len(nums) >= 2: H = nums[1]
        data.update({"base_diameter_mm": BD or DEFAULTS["cone"]["base_diameter_mm"], "height_mm": H or DEFAULTS["cone"]["height_mm"]})

    elif shape == "wedge":
        B = None; H = None; D = None
        mB = re.search(r"base\s*(\d+(?:\.\d+)?)", prompt.lower())
        mH = re.search(r"height\s*(\d+(?:\.\d+)?)", prompt.lower())
        mD = re.search(r"(depth|extrude)\s*(\d+(?:\.\d+)?)", prompt.lower())
        if mB: B = float(mB.group(1))
        if mH: H = float(mH.group(1))
        if mD: D = float(mD.group(2))
        if B is None and nums: B = nums[0]
        if H is None and len(nums) >= 2: H = nums[1]
        if D is None and len(nums) >= 3: D = nums[2]
        data.update({"base_mm": B or DEFAULTS["wedge"]["base_mm"], "height_mm": H or DEFAULTS["wedge"]["height_mm"], "depth_mm": D or DEFAULTS["wedge"]["depth_mm"]})

    elif shape == "cup":
        OD = None; WT = None; HT = None
        mOD = re.search(r"(outer\s*dia|outer\s*diameter|diameter)\s*(\d+(?:\.\d+)?)", prompt.lower())
        mWT = re.search(r"(wall|thickness)\s*(\d+(?:\.\d+)?)", prompt.lower())
        mHT = re.search(r"height\s*(\d+(?:\.\d+)?)", prompt.lower())
        if mOD: OD = float(mOD.group(2))
        if mWT: WT = float(mWT.group(2))
        if mHT: HT = float(mHT.group(1))
        nums_iter = iter(nums)
        if OD is None:
            try: OD = next(nums_iter)
            except StopIteration: pass
        if WT is None:
            try: WT = next(nums_iter)
            except StopIteration: pass
        if HT is None:
            try: HT = next(nums_iter)
            except StopIteration: pass
        data.update({
            "outer_d_mm": OD or DEFAULTS["cup"]["outer_d_mm"],
            "wall_mm": WT or DEFAULTS["cup"]["wall_mm"],
            "height_mm": HT or DEFAULTS["cup"]["height_mm"],
        })

    if "plane" not in data:
        data["plane"] = DEFAULTS[shape]["plane"]

    return data

def parse_kv_params(kv: str) -> Dict:
    out = {}
    for item in kv.split(","):
        item = item.strip()
        if not item or "=" not in item:
            continue
        k, v = item.split("=", 1)
        k = k.strip(); v = v.strip()
        try:
            out[k] = float(v)
        except ValueError:
            out[k] = v
    return out
