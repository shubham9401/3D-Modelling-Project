"""
Part module for MCP-style SolidWorks CAD engine.
"""

try:
    from .solidworks_app import get_sw_app, get_active_model, create_new_part, get_nothing
except ImportError:
    from solidworks_app import get_sw_app, get_active_model, create_new_part, get_nothing

_PART_ACTIVE = False

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
    val = model.GetType
    if callable(val):
        val = val()
    return val == 1

def create_part():
    """Creates a new part document."""
    global _PART_ACTIVE
    model = create_new_part()
    if model is None:
        raise Exception("Failed to create part document")
    _PART_ACTIVE = True
    return "Part document created"

def close_part(save=False):
    """Closes the active part document."""
    global _PART_ACTIVE
    _require_part_active()
    model = _model()
    if save:
        model.Save3(1, 0, 0)
    title = model.GetTitle()
    _sw().CloseDoc(title)
    _PART_ACTIVE = False
    return "Part document closed"

def validate_part():
    """Validates that active document is a PART."""
    _require_part_active()
    model = _model()
    if not _is_part(model):
        raise Exception("Active document is not a part")
    return "Part validated"

def save_part(path):
    """Saves the active part to disk."""
    _require_part_active()
    model = _model()
    if not path.lower().endswith('.sldprt'):
        path = path + '.sldprt'
    success = model.SaveAs3(path, 0, 0)
    if not success:
        raise Exception("Failed to save part")
    return f"Part saved at {path}"

def get_part_info():
    """Returns part metadata."""
    _require_part_active()
    model = _model()
    return {
        "name": model.GetTitle(),
        "type": "Part",
        "feature_count": model.GetFeatureCount(False)
    }