"""
Tools package - Exposes all SolidWorks automation modules.
"""

try:
    from .solidworks_app import get_sw_app, get_active_model, get_nothing, create_new_part
    from . import part, sketch, feature, assembly
except ImportError:
    from solidworks_app import get_sw_app, get_active_model, get_nothing, create_new_part
    import part, sketch, feature, assembly
