"""
Assembly module for MCP-style SolidWorks CAD engine.

Supports:
- Assembly lifecycle
- Component insertion
- Fix/float
- Full set of essential mates

Design:
- Semantic mate functions
- No geometry logic here
- Selection assumed to be done before mate call
"""

try:
    from .solidworks_app import get_sw_app, get_active_model, create_new_assembly
except ImportError:
    from solidworks_app import get_sw_app, get_active_model, create_new_assembly

# ============================================================
# INTERNAL STATE
# ============================================================

_ASSEMBLY_ACTIVE = False

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

def _require_assembly_active():
    if not _ASSEMBLY_ACTIVE:
        raise Exception("No active assembly document")

def _is_assembly(model):
    return model.GetType() == 2  # swDocASSEMBLY

def _add_mate(mate_type, value=0):
    model = _model()
    errors = 0
    model.AddMate5(
        mate_type,
        0,
        False,
        value,
        0, 0, 0, 0,
        errors
    )
    if errors != 0:
        raise Exception(f"Mate failed (type={mate_type})")

# ============================================================
# ASSEMBLY LIFECYCLE
# ============================================================

def create_assembly():
    global _ASSEMBLY_ACTIVE
    model = create_new_assembly()
    if model is None:
        raise Exception("Failed to create assembly")
    _ASSEMBLY_ACTIVE = True
    return "Assembly created"

def close_assembly(save=False):
    global _ASSEMBLY_ACTIVE
    _require_assembly_active()
    model = _model()
    if save:
        model.Save3(1, 0, 0)
    _sw().CloseDoc(model.GetTitle())
    _ASSEMBLY_ACTIVE = False
    return "Assembly closed"

def validate_assembly():
    _require_assembly_active()
    if not _is_assembly(_model()):
        raise Exception("Active document is not an assembly")
    return "Assembly validated"

# ============================================================
# COMPONENT OPERATIONS
# ============================================================

def insert_part(path, x=0, y=0, z=0):
    _require_assembly_active()
    comp = _model().AddComponent5(
        path,
        0, "", False, "",
        x / 1000, y / 1000, z / 1000
    )
    if comp is None:
        raise Exception(f"Failed to insert part: {path}")
    return f"Inserted part {path}"

def fix_component():
    _require_assembly_active()
    _model().FixComponent()
    return "Component fixed"

def float_component():
    _require_assembly_active()
    _model().UnfixComponent()
    return "Component floated"

# ============================================================
# ESSENTIAL STANDARD MATES
# ============================================================

def mate_coincident():
    _require_assembly_active()
    _add_mate(0)
    return "Coincident mate applied"

def mate_concentric():
    _require_assembly_active()
    _add_mate(1)
    return "Concentric mate applied"

def mate_parallel():
    _require_assembly_active()
    _add_mate(3)
    return "Parallel mate applied"

def mate_perpendicular():
    _require_assembly_active()
    _add_mate(4)
    return "Perpendicular mate applied"

def mate_distance(distance):
    _require_assembly_active()
    _add_mate(5, distance / 1000)
    return f"Distance mate {distance} mm applied"

def mate_angle(angle_deg):
    _require_assembly_active()
    _add_mate(6, angle_deg * 3.14159 / 180)
    return f"Angle mate {angle_deg}° applied"

def mate_tangent():
    _require_assembly_active()
    _add_mate(2)
    return "Tangent mate applied"

# ============================================================
# MECHANICAL MATES (IMPORTANT)
# ============================================================

def mate_lock():
    _require_assembly_active()
    _add_mate(10)
    return "Lock mate applied"

def mate_width():
    _require_assembly_active()
    _add_mate(11)
    return "Width mate applied"

def mate_hinge():
    """
    Combination mate (concentric + coincident)
    """
    _require_assembly_active()
    _add_mate(1)  # concentric
    _add_mate(0)  # coincident
    return "Hinge mate applied"

# ============================================================
# ASSEMBLY INFO
# ============================================================

def get_component_count():
    _require_assembly_active()
    return _model().GetComponentCount(False)
