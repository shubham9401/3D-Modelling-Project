"""
solidworks_app.py
Handles the connection between Python and SolidWorks.
"""
import win32com.client
import pythoncom
import os

# Template paths
PART_TEMPLATE = r"C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2021\templates\Part.prtdot"
ASSEMBLY_TEMPLATE = r"C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2021\templates\Assembly.asmdot"

# Cached app reference
_SW_APP = None

def get_nothing():
    """Returns a VT_DISPATCH None variant for API calls."""
    return win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)

def get_sw_app():
    """Connects to SolidWorks application."""
    global _SW_APP
    
    if _SW_APP is not None:
        return _SW_APP
    
    try:
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
    """Creates a new part document."""
    app = get_sw_app()
    if not app:
        raise Exception("SolidWorks is not available")
    
    if not os.path.exists(PART_TEMPLATE):
        print(f"Warning: Template not found at {PART_TEMPLATE}")
        model = app.NewDocument("", 0, 0, 0)
    else:
        model = app.NewDocument(PART_TEMPLATE, 0, 0, 0)
    
    if model is None:
        raise Exception("Failed to create new part document")
    
    return model

def create_new_assembly():
    """Creates a new assembly document."""
    app = get_sw_app()
    if not app:
        raise Exception("SolidWorks is not available")
    
    if not os.path.exists(ASSEMBLY_TEMPLATE):
        model = app.NewDocument("", 2, 0, 0)
    else:
        model = app.NewDocument(ASSEMBLY_TEMPLATE, 0, 0, 0)
    
    if model is None:
        raise Exception("Failed to create new assembly document")
    
    return model