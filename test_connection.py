# sw_connect_test.py
import win32com.client

try:
    sw = win32com.client.Dispatch("SldWorks.Application")
    sw.Visible = True
    print("Connected to SolidWorks successfully.")
except Exception as e:
    print("Failed to connect to SolidWorks.")
    print("Error:", e)