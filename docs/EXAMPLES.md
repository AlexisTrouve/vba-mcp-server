# Usage Examples

This document provides practical examples of using the VBA MCP Server with Claude Code.

## Prerequisites

Ensure the MCP server is configured in your Claude Code settings (see README.md).

## Basic Examples

### Example 1: Extract VBA from Excel File

**Scenario:** You have an Excel file with macros and want to see the code.

**User message to Claude Code:**
```
Extract the VBA code from C:/Users/john/Documents/budget.xlsm
```

**What happens:**
1. Claude Code calls `extract_vba` tool
2. Server extracts all modules
3. Claude displays the code organized by module

**Expected output:**
```
I found 3 VBA modules in budget.xlsm:

**Module1** (Standard Module):
```vba
Sub CalculateBudget()
    Dim total As Double
    total = Worksheets("Data").Range("A1").Value
    MsgBox "Total: " & Format(total, "$#,##0.00")
End Sub
```

**Sheet1** (Worksheet Module):
```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Column = 1 Then
        Call CalculateBudget
    End If
End Sub
```

**ThisWorkbook** (Workbook Module):
```vba
Private Sub Workbook_Open()
    MsgBox "Welcome!"
End Sub
```
```

---

### Example 2: List Modules Only

**Scenario:** You want to see what modules exist before extracting.

**User message:**
```
What VBA modules are in report.xlsm?
```

**Claude Code calls:** `list_modules`

**Expected output:**
```
The file report.xlsm contains 4 VBA modules:

1. **Module1** (Standard) - 150 lines, 5 procedures
2. **Module2** (Standard) - 200 lines, 8 procedures
3. **Sheet1** (Worksheet) - 45 lines, 2 procedures
4. **ThisWorkbook** (Workbook) - 30 lines, 3 procedures

Total: 425 lines of VBA code
```

---

### Example 3: Extract Specific Module

**Scenario:** You only want one specific module.

**User message:**
```
Show me just Module1 from budget.xlsm
```

**Claude Code calls:** `extract_vba` with `module_name: "Module1"`

**Expected output:**
```
Here's Module1 from budget.xlsm:

```vba
Option Explicit

Sub CalculateBudget()
    Dim total As Double
    Dim tax As Double

    total = GetRevenue()
    tax = total * 0.2

    MsgBox "Total: " & Format(total + tax, "$#,##0.00")
End Sub

Function GetRevenue() As Double
    GetRevenue = Worksheets("Data").Range("A1").Value
End Function
```
```

---

### Example 4: Analyze Code Structure

**Scenario:** You want to understand the code organization and complexity.

**User message:**
```
Analyze the VBA structure in my Excel file
```

**Claude Code calls:** `analyze_structure`

**Expected output:**
```
Analysis of budget.xlsm:

**Procedures:**
- CalculateBudget (Sub) - Complexity: 3, Calls: GetRevenue, MsgBox
- GetRevenue (Function) - Complexity: 1, Called by: CalculateBudget
- Worksheet_Change (Event) - Complexity: 2, Calls: CalculateBudget

**Dependencies:**
- Sheet1 depends on Module1
- Module1 has no dependencies

**External References:**
- None

**Metrics:**
- Total procedures: 3
- Average complexity: 2 (Good)
- Total lines: 40

**Recommendations:**
✅ Code complexity is low - well structured
✅ No circular dependencies detected
```

---

## Advanced Examples

### Example 5: Refactoring Workflow

**Step 1:** Extract and analyze

```
User: Extract and analyze the code from legacy_report.xlsm
```

Claude extracts + analyzes, finds high complexity (15) in `GenerateReport` procedure.

**Step 2:** Ask for refactoring help

```
User: The GenerateReport function is too complex. Can you help refactor it?
```

Claude uses the extracted code to suggest:
```
I'll break down GenerateReport into smaller functions:

1. **SeparateDataCollection** - Extract data gathering logic
2. **SeparateFormatting** - Extract formatting logic
3. **SeparateOutput** - Extract output logic
4. **KeepMainOrchestration** - Keep GenerateReport as coordinator

Here's the refactored code:
[Claude provides improved code]
```

**Step 3:** (Pro version) Apply changes

```
User: Apply these changes to the file
```

(Pro version would reinject the modified code)

---

### Example 6: Understanding Dependencies

**User message:**
```
Show me all the dependencies between modules in accounting.xlsm
```

**Expected output:**
```
Dependency graph for accounting.xlsm:

Module1 (root)
 └─> Module2
      ├─> Module3
      └─> Module4
           └─> Module1 ⚠️ CIRCULAR DEPENDENCY

**Warning:** Circular dependency detected!
- Module1 calls Module2
- Module2 calls Module4
- Module4 calls Module1

This can cause stack overflow errors. Consider refactoring.
```

---

### Example 7: Find Specific Code

**User message:**
```
Find all procedures that use Worksheets() in my Excel file
```

Claude calls `analyze_structure`, then searches the results:

**Expected output:**
```
Found 3 procedures using Worksheets():

1. **Module1.GetData**
   - Line 15: `Set ws = Worksheets("Data")`

2. **Module2.SaveResults**
   - Line 8: `Worksheets("Output").Range("A1").Value = result`

3. **ThisWorkbook.Workbook_Open**
   - Line 3: `Worksheets(1).Activate`
```

---

### Example 8: Code Quality Check

**User message:**
```
Check the code quality of my VBA project
```

**Expected output:**
```
Code Quality Report for project.xlsm:

**✅ Good Practices:**
- All modules use Option Explicit
- Consistent naming conventions
- No hardcoded values in procedures

**⚠️ Warnings:**
- 2 procedures have complexity > 10
- Module3 has 300+ lines (consider splitting)
- No error handling in 5 procedures

**❌ Issues:**
- GlobalVariable found in Module1 (avoid global state)
- On Error Resume Next in Module2 line 45 (masks errors)

**Recommendations:**
1. Refactor CalculateAll() - complexity 15
2. Add error handling to database procedures
3. Replace global variables with parameters
```

---

## Integration Examples

### Example 9: With Git Workflow

**Scenario:** You want to version control VBA code.

```bash
# 1. Extract VBA to text files
User: "Extract all VBA from project.xlsm to separate .bas files"

# Claude extracts each module to:
# - Module1.bas
# - Module2.bas
# - Sheet1.cls

# 2. Commit to git
git add *.bas *.cls
git commit -m "feat: Add budget calculation logic"

# 3. Later, review changes
git diff Module1.bas
```

---

### Example 10: Documentation Generation

**User message:**
```
Generate documentation for all public procedures in my VBA project
```

**Expected output:**
```
# API Documentation

## Module1

### CalculateBudget()
**Type:** Sub (Public)
**Parameters:** None
**Description:** Calculates total budget including tax
**Calls:** GetRevenue, MsgBox

### GetRevenue() As Double
**Type:** Function (Public)
**Parameters:** None
**Returns:** Double - Revenue from Data sheet
**Called by:** CalculateBudget

## Module2

### SaveToDatabase(data As Variant)
**Type:** Sub (Public)
**Parameters:**
- data (Variant) - Data to save
**Calls:** ADODB.Connection.Execute
```

---

## Error Handling Examples

### Example 11: File Not Found

**User message:**
```
Extract VBA from nonexistent.xlsm
```

**Expected output:**
```
❌ Error: File not found

The file "nonexistent.xlsm" doesn't exist. Please check:
- File path is correct
- File hasn't been moved/deleted
- You have permission to access the file
```

---

### Example 12: No VBA Macros

**User message:**
```
Extract VBA from data.xlsx
```

**Expected output:**
```
ℹ️ This Excel file contains no VBA macros.

The file "data.xlsx" is a standard Excel file (.xlsx) without macros.
If you need to work with macros, save it as .xlsm format.
```

---

### Example 13: Password Protected

**User message:**
```
Extract VBA from protected.xlsm
```

**Expected output:**
```
🔒 This file is password protected.

Please unlock the file first, then try again.
```

---

## Performance Examples

### Example 14: Large File Handling

**Scenario:** Processing a 50MB Excel file with 100+ modules.

```
User: Extract VBA from huge_project.xlsm

Claude: This is a large file (50 MB). Extraction may take 10-30 seconds...

[Progress indicator]

✅ Extraction complete! Found 127 modules with 15,000+ lines of code.

Would you like me to:
1. Show a summary of all modules
2. Extract a specific module
3. Analyze code structure
```

---

## Tips & Tricks

### Tip 1: Batch Processing

```
User: I have 5 Excel files. Extract VBA from all of them and compare the code.

Claude calls extract_vba for each file, then analyzes differences.
```

### Tip 2: Search Across Files

```
User: Find all VBA code that connects to databases across my project files

Claude extracts from all .xlsm files, searches for ADODB/DAO references
```

### Tip 3: Code Migration

```
User: I need to migrate this Access VBA to Excel. What changes are needed?

Claude compares Access-specific vs Excel-specific APIs
```

---

## Next Steps

- See [API.md](API.md) for technical details
- See [ARCHITECTURE.md](ARCHITECTURE.md) for system design
- For pro features (modification/reinjection), contact for licensing
