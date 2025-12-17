#!/usr/bin/env python3
"""
Create a test Excel file with VBA macros for testing VBA MCP Pro.
"""

import win32com.client
from pathlib import Path

def create_test_excel():
    """Create test.xlsm with simple VBA macros."""
    output_path = Path(__file__).parent / "test.xlsm"

    print(f"Creating test Excel file: {output_path}")

    # Create Excel application
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        # Create new workbook
        wb = excel.Workbooks.Add()

        # Setup Sheet1 with sample data
        ws = wb.Worksheets(1)
        ws.Name = "Sheet1"

        # Add headers
        ws.Cells(1, 1).Value = "Name"
        ws.Cells(1, 2).Value = "Age"
        ws.Cells(1, 3).Value = "Score"

        # Add sample data
        data = [
            ("Alice", 30, 95),
            ("Bob", 25, 87),
            ("Charlie", 35, 92),
            ("Diana", 28, 88),
        ]

        for i, (name, age, score) in enumerate(data, start=2):
            ws.Cells(i, 1).Value = name
            ws.Cells(i, 2).Value = age
            ws.Cells(i, 3).Value = score

        # Auto-fit columns
        ws.Columns("A:C").AutoFit()

        # Add VBA code
        vb_module = wb.VBProject.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
        vb_module.Name = "TestModule"

        vba_code = '''
' Test Module for VBA MCP Pro
Option Explicit

' Simple function that returns a value
Public Function HelloWorld() As String
    HelloWorld = "Hello from VBA!"
End Function

' Function with parameters
Public Function AddNumbers(a As Long, b As Long) As Long
    AddNumbers = a + b
End Function

' Function that returns double
Public Function MultiplyNumbers(a As Double, b As Double) As Double
    MultiplyNumbers = a * b
End Function

' Sub that displays a message
Public Sub ShowMessage()
    MsgBox "This is a test macro!", vbInformation, "Test"
End Sub

' Function to calculate sum of range
Public Function CalculateSum() As Double
    Dim ws As Worksheet
    Dim total As Double
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets("Sheet1")
    total = 0

    For i = 2 To 5
        total = total + ws.Cells(i, 3).Value
    Next i

    CalculateSum = total
End Function

' Sub to create a summary sheet
Public Sub CreateSummary()
    Dim ws As Worksheet
    Dim summarySheet As Worksheet

    ' Delete if exists
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("Summary").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Create new summary sheet
    Set summarySheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    summarySheet.Name = "Summary"

    ' Add summary data
    summarySheet.Cells(1, 1).Value = "Total Records"
    summarySheet.Cells(1, 2).Value = 4

    summarySheet.Cells(2, 1).Value = "Average Score"
    summarySheet.Cells(2, 2).Value = CalculateSum() / 4

    summarySheet.Columns("A:B").AutoFit()
End Sub
'''

        vb_module.CodeModule.AddFromString(vba_code)

        # Save as macro-enabled workbook
        wb.SaveAs(str(output_path.absolute()), FileFormat=52)  # 52 = xlOpenXMLWorkbookMacroEnabled
        wb.Close()

        print(f"[OK] Test file created: {output_path}")
        print(f"[OK] Contains module: TestModule")
        print(f"[OK] Contains functions: HelloWorld, AddNumbers, MultiplyNumbers, CalculateSum")
        print(f"[OK] Contains subs: ShowMessage, CreateSummary")
        print(f"[OK] Contains data: 4 rows in Sheet1")

    except Exception as e:
        print(f"[ERROR] Error: {e}")
        raise

    finally:
        excel.Quit()

if __name__ == "__main__":
    create_test_excel()
