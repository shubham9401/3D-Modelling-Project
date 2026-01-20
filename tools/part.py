"""
Part module for MCP-style SolidWorks CAD engine.

Responsibilities:
- Part document lifecycle
- Validation and safety guards
- File save operations
- Feature-readiness checks

Does NOT:
- Create sketches
- Create geometry
- Create features
"""

from solidworks_app import get_sw_app, get_active_model

# ============================================================
# INTERNAL STATE
# ============================================================

_PART_ACTIVE = False

# ============================================================
# INTERNAL HELPERS
# ============================================================

def _sw():
    sw = get_sw_app()
    if sw is None:
        raise Exception("SolidWorks application not available")
    return sw

def _model():
    model = get_active_model()
    if model is None:
        raise Exception("No active SolidWorks document")
    return model

def _require_part_active():
    if not _PART_ACTIVE:
        raise Exception("No active part document")

def _is_part(model):
    # swDocPART = 1
    return model.GetType() == 1

# ============================================================
# PART LIFECYCLE
# ============================================================

def create_part():
    """
    Creates a new part document.
    """
    global _PART_ACTIVE

    sw = _sw()
    model = sw.NewDocument("", 0, 0, 0)

    if model is None:
        raise Exception("Failed to create part document")

    _PART_ACTIVE = True
    return "Part document created"


def close_part(save=False):
    """
    Closes the active part document.
    """
    global _PART_ACTIVE

    _require_part_active()
    model = _model()

    if save:
        model.Save3(1, 0, 0)

    title = model.GetTitle()
    _sw().CloseDoc(title)

    _PART_ACTIVE = False
    return "Part document closed"

# ============================================================
# PART VALIDATION
# ============================================================

def validate_part():
    """
    Ensures:
    - Active document exists
    - Document is a PART
    """
    _require_part_active()
    model = _model()

    if not _is_part(model):
        raise Exception("Active document is not a part")

    return "Part validated"

# ============================================================
# FEATURE SAFETY GUARDS
# ============================================================

def require_ready_for_feature():
    """
    Ensures part is ready before feature creation.
    Call this BEFORE extrude / cut / revolve.
    """
    _require_part_active()
    model = _model()

    if not _is_part(model):
        raise Exception("Active document is not a part")

    sketch = model.GetActiveSketch2()
    if sketch is None:
        raise Exception("No sketch available for feature creation")

    if not sketch.IsClosed():
        raise Exception("Sketch is not closed – cannot create feature")

    return "Part ready for feature creation"

# ============================================================
# FILE OPERATIONS
# ============================================================

def save_part(path):
    """
    Saves the active part to disk.
    """
    _require_part_active()
    model = _model()

    success = model.SaveAs3(path, 0, 0)
    if not success:
        raise Exception("Failed to save part")

    return f"Part saved at {path}"


def save_part_incremental(directory, base_name):
    """
    Saves part with auto-incremented filename.
    Example: block_001.sldprt
    """
    _require_part_active()
    model = _model()

    import os

    i = 1
    while True:
        filename = f"{base_name}_{i:03d}.sldprt"
        path = os.path.join(directory, filename)
        if not os.path.exists(path):
            break
        i += 1

    model.SaveAs3(path, 0, 0)
    return f"Part saved as {filename}"

# ============================================================
# PART INFO / DEBUG
# ============================================================

def get_part_info():
    """
    Returns lightweight part metadata (useful for logging).
    """
    _require_part_active()
    model = _model()

    return {
        "name": model.GetTitle(),
        "type": "Part",
        "feature_count": model.GetFeatureCount(False)
    }
