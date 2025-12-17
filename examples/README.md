# Example Files

This directory contains sample Office files with VBA macros for testing.

## Creating Test Files

### Option 1: Manual Creation (Recommended)

1. Open Excel
2. Create a new workbook
3. Press `Alt+F11` to open VBA editor
4. Insert a new module (Insert > Module)
5. Paste the sample code below
6. Save as `sample.xlsm` (Excel Macro-Enabled Workbook) in this directory

### Sample VBA Code

```vba
' Module: TestModule
' Description: Simple test module for VBA extraction

Option Explicit

Public Const APP_VERSION As String = "1.0.0"

' Simple hello world function
Public Function HelloWorld() As String
    HelloWorld = "Hello from VBA!"
End Function

' Calculate sum of array
Public Function SumArray(numbers() As Double) As Double
    Dim total As Double
    Dim i As Long

    total = 0
    For i = LBound(numbers) To UBound(numbers)
        total = total + numbers(i)
    Next i

    SumArray = total
End Function

' Subroutine example
Public Sub LogMessage(message As String)
    Debug.Print "[" & Now & "] " & message
End Sub
```

### Option 2: Programmatic Creation (Windows Only)

Run the helper script (requires Excel installed):

```bash
python create_sample.py
```

## Test Files

- `sample.xlsm` - Basic Excel file with test macros
- `complex.xlsm` - More complex example with multiple modules
- `empty.xlsm` - Excel file without macros (for error testing)
