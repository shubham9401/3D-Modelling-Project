"""
solidworks_app.py
Handles the connection between Python and SolidWorks.
"""
import win32com.client
import pythoncom
import os

# Template path for new parts (adjust if your SW version differs)
PART_TEMPLATE = r"C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2021\templates\Part.prtdot"
ASSEMBLY_TEMPLATE = r"C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2021\templates\Assembly.asmdot"

# Cached app reference
_SW_APP = None

def get_nothing():
    """Returns a VT_DISPATCH None variant for API calls that require it."""
    return win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)


def get_sw_app():
    """
    Connects to SolidWorks application.
    Uses Dispatch to connect to running instance or start new one.
    """
    global _SW_APP
    
    if _SW_APP is not None:
        return _SW_APP
    
    try:
        # Try to connect to existing SolidWorks instance
        _SW_APP = win32com.client.Dispatch("SldWorks.Application")
        _SW_APP.Visible = True
        return _SW_APP
    except Exception as e:
        print(f"Error connecting to SolidWorks: {e}")
        return None


def get_active_model():
    """Gets the currently open document."""
    app = get_sw_app()
    if not app:
        return None
    return app.ActiveDoc


def create_new_part():
    """
    Creates a new part document using the template.
    Returns the model object.
    """
    app = get_sw_app()
    if not app:
        raise Exception("SolidWorks is not available")
    
    # Check if template exists
    if not os.path.exists(PART_TEMPLATE):
        # Fallback: try to create without template
        print(f"Warning: Template not found at {PART_TEMPLATE}")
        print("Attempting to create part without template...")
        model = app.NewDocument("", 0, 0, 0)
    else:
        # Use template
        model = app.NewDocument(PART_TEMPLATE, 0, 0, 0)
    
    if model is None:
        raise Exception("Failed to create new part document")
    
    return model


def create_new_assembly():
    """
    Creates a new assembly document using the template.
    Returns the model object.
    """
    app = get_sw_app()
    if not app:
        raise Exception("SolidWorks is not available")
    
    if not os.path.exists(ASSEMBLY_TEMPLATE):
        model = app.NewDocument("", 2, 0, 0)  # 2 = swDocASSEMBLY
    else:
        model = app.NewDocument(ASSEMBLY_TEMPLATE, 0, 0, 0)
    
    if model is None:
        raise Exception("Failed to create new assembly document")
    
    return model