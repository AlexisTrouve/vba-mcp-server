"""
Create sample Excel file with VBA macros for testing.

Requires: Windows + Microsoft Excel installed
"""

import sys
from pathlib import Path

try:
    import win32com.client
except ImportError:
    print("Error: pywin32 required. Install with: pip install pywin32")
    sys.exit(1)


SAMPLE_VBA_CODE = '''
Option Explicit

Public Const APP_VERSION As String = "1.0.0"

Public Function HelloWorld() As String
    HelloWorld = "Hello from VBA!"
End Function

Public Function SumArray(numbers() As Double) As Double
    Dim total As Double
    Dim i As Long

    total = 0
    For i = LBound(numbers) To UBound(numbers)
        total = total + numbers(i)
    Next i

    SumArray = total
End Function

Public Sub LogMessage(message As String)
    Debug.Print "[" & Now & "] " & message
End Sub
'''


def create_sample_xlsm():
    """Create sample.xlsm with test VBA code."""
    print("Creating sample Excel file with VBA...")

    # Launch Excel
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        # Create new workbook
        wb = excel.Workbooks.Add()

        # Add VBA module
        vb_module = wb.VBProject.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
        vb_module.Name = "TestModule"
        vb_module.CodeModule.AddFromString(SAMPLE_VBA_CODE)

        # Save as .xlsm
        output_path = Path(__file__).parent / "sample.xlsm"
        wb.SaveAs(
            str(output_path.absolute()),
            FileFormat=52  # xlOpenXMLWorkbookMacroEnabled
        )

        print(f"[OK] Created: {output_path}")

        # Close workbook
        wb.Close(SaveChanges=False)

    except Exception as e:
        print(f"[ERROR] {e}")
        raise
    finally:
        excel.Quit()


if __name__ == "__main__":
    create_sample_xlsm()
