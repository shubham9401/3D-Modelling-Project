"""
solidworks_app.py
Handles the connection between Python and SolidWorks.
"""
import win32com.client
import pythoncom

def get_sw_app():
    """Connects to the running SolidWorks instance."""
    try:
        # Connect to existing SolidWorks
        return win32com.client.GetActiveObject("SldWorks.Application")
    except Exception as e:
        # If SW is not open, this fails
        return None

def get_active_model():
    """Gets the currently open document."""
    app = get_sw_app()
    if not app:
        return None
    return app.ActiveDoc