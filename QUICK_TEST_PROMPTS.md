# Quick Test Prompts for VBA MCP Pro

Copy-paste these prompts into Claude Desktop to test the VBA MCP server.

---

## 🔍 Basic Tests (Read-Only)

### 1. Check Available Tools
```
What VBA MCP tools do you have available?
```

### 2. List Modules
```
List all VBA modules in C:\Users\alexi\Documents\projects\vba-mcp-monorepo\test.xlsm
```

### 3. Extract Code
```
Extract the VBA code from TestModule in test.xlsm
```

### 4. Analyze Structure
```
Analyze the structure of the VBA code in test.xlsm
```

---

## 🔍 VBA Validation Tests (NEW!)

### 5. Validate VBA Code (Valid)
```
Validate this VBA code:
Sub HelloWorld()
    MsgBox "Hello from VBA!"
End Sub
```
**Expected:** Code Valid with successful compilation message

### 6. Validate VBA Code (Invalid Syntax)
```
Validate this VBA code:
Sub Test()
    If x = 1
    MsgBox "Test"
End Sub
```
**Expected:** Code Invalid with error about missing End If

### 7. Validate VBA Code (Unicode Characters)
```
Validate this VBA code:
Sub Test()
    MsgBox "✓ Success"
End Sub
```
**Expected:** Code Invalid with ASCII replacement suggestions

### 8. List All Macros
```
List all the macros in C:\Users\alexi\Documents\projects\vba-mcp-monorepo\test.xlsm
```
**Expected:** Shows all public Subs and Functions with signatures, grouped by module

---

## 📊 Excel Table Operations Tests (NEW in v0.3.0!)

### 8A. List All Tables
```
List all Excel tables in C:\Users\alexi\Documents\projects\vba-mcp-monorepo\test.xlsm
```
**Expected:** Shows all tables with names, sheets, dimensions, headers

### 8B. Create Table from Range
```
In test.xlsm, convert range A1:C10 on Sheet1 to an Excel table named TestTable
```
**Expected:** Creates formatted table with default style

### 8C. Insert Rows in Table
```
Insert 3 rows at position 5 in TestTable in test.xlsm, Sheet1
```
**Expected:** Adds 3 blank rows to table, shifts existing data down

### 8D. Delete Rows from Table
```
Delete rows 8-10 from TestTable in test.xlsm
```
**Expected:** Removes specified rows from table

### 8E. Insert Column in Table
```
Insert a column named "Grade" after column C in TestTable, in test.xlsm
```
**Expected:** New column added to table with header "Grade"

### 8F. Delete Columns by Header Name
```
Delete columns ["Grade", "TempData"] from TestTable in test.xlsm
```
**Expected:** Removes specified columns from table

### 8G. Read Table Data with Column Selection
```
Get data from TestTable in test.xlsm, only columns ["Name", "Score"]
```
**Expected:** Returns only specified columns as 2D array

### 8H. Read Entire Table
```
Get all data from table TestTable in test.xlsm, Sheet1, include headers
```
**Expected:** Returns full table including headers

---

## 🚀 Office Automation Tests

### 9. Open File
```
Open C:\Users\alexi\Documents\projects\vba-mcp-monorepo\test.xlsm in Excel
```
**Expected:** Excel launches with file visible

### 10. List Open Files
```
List all currently open Office files
```

### 11. Run Simple Macro
```
Run the HelloWorld function in test.xlsm
```
**Expected:** Returns "Hello from VBA!"

### 12. Run Macro with Parameters
```
Run the AddNumbers function in test.xlsm with arguments 123 and 456
```
**Expected:** Returns 579

### 13. Run Macro with Improved Error Messages (NEW!)
```
Run a macro called NonExistentMacro in test.xlsm
```
**Expected:** Error message lists all available macros with their signatures

### 14. Read Worksheet Data
```
Get the data from Sheet1 in test.xlsm, range A1:C5
```
**Expected:** JSON array with names, ages, scores

### 15. Write Data to Sheet
```
Write the following data to Sheet2 in test.xlsm starting at A1:
[[1, 2, 3], [4, 5, 6], [7, 8, 9]]
```
**Expected:** Data written, can verify in Excel

### 16. Run Complex Macro
```
Run the CalculateSum function in test.xlsm
```
**Expected:** Returns sum of scores (95 + 87 + 92 + 88 = 362)

### 17. Run Macro that Creates Sheet
```
Run the CreateSummary sub in test.xlsm
```
**Expected:** Creates Summary sheet with statistics

### 18. Close File
```
Close test.xlsm and save changes
```
**Expected:** Excel closes, file saved

---

## 🔧 Advanced Workflows

### Validation Workflow (NEW! - Recommended Best Practice)

```
I want to add a new VBA function to test.xlsm. Follow this workflow:

1. First, validate this code:
   Function MultiplyNumbers(a As Double, b As Double) As Double
       MultiplyNumbers = a * b
   End Function

2. If validation passes, extract the current code from TestModule

3. Add the new function to the existing code

4. Validate the complete updated code

5. If validation passes, inject the updated code into test.xlsm

6. List all macros to confirm the new function is there

7. Run MultiplyNumbers with arguments 15 and 4 to test it
```
**Expected:** Validation catches errors before injection, ensuring clean workflow

### Excel Table Processing Workflow (NEW!)

```
I need to work with the sales data in test.xlsm. Follow this workflow:

1. Open test.xlsm in Excel

2. Create an Excel table from range A1:D20 on Sheet1, name it "SalesData"

3. List all tables to confirm it was created

4. Get the data from columns ["Product", "Revenue"] in SalesData table

5. Insert a new column named "Profit" after column D in the table

6. Calculate profit (Revenue * 0.3) and write it to the Profit column

7. Insert 5 new rows at the end of the table for new entries

8. Delete any rows where Revenue is 0

9. Close and save the file
```
**Expected:** Complete table manipulation workflow demonstrating CRUD operations

### Extract, Modify, Inject

```
1. Extract the VBA code from TestModule in test.xlsm
2. Add a new function called SquareNumber that takes a number and returns its square
3. Validate the modified code first
4. Inject the modified code back into test.xlsm
5. Run SquareNumber with argument 12
```

### Data Processing Pipeline

```
1. Open test.xlsm
2. Get all data from Sheet1
3. Calculate the average age and average score
4. Write a summary to Sheet2 with the calculated averages
5. Close and save the file
```

### Refactoring Suggestions

```
Analyze the VBA code in test.xlsm and suggest refactoring improvements
```

---

## 🎯 Real-World Scenarios

### Scenario 1: Legacy Code Analysis

```
I have an old Excel file with VBA macros. Can you:
1. List all the modules
2. Extract and analyze the code
3. Identify potential issues or improvements
4. Suggest modern alternatives

File: C:\path\to\your\file.xlsm
```

### Scenario 2: Automated Data Processing

```
I need to:
1. Open budget.xlsm
2. Run the CalculateMonthlyTotals macro
3. Get the results from Summary sheet
4. Export that data to a new file

Can you help?
```

### Scenario 3: Macro Debugging

```
I have a macro in invoices.xlsm called GenerateReport that's not working.
Can you:
1. Extract the code
2. Analyze it for errors
3. Suggest fixes
4. Test it with sample data
```

---

## 📊 File Paths Reference

**Test file:**
```
C:\Users\alexi\Documents\projects\vba-mcp-monorepo\test.xlsm
```

**Demo files (if created):**
```
C:\Users\alexi\Documents\projects\vba-mcp-demo\examples\budget-calculator.xlsm
C:\Users\alexi\Documents\projects\vba-mcp-demo\examples\data-importer.xlsm
C:\Users\alexi\Documents\projects\vba-mcp-demo\examples\report-generator.xlsm
```

---

## ⚡ Pro Tips

1. **Keep files open:** After opening a file, you can run multiple operations without reopening (faster!)

2. **Session timeout:** Files auto-close after 1 hour of inactivity

3. **Backups:** Injection creates automatic backups with timestamp

4. **Large data:** For data > 10,000 cells, specify exact range for better performance

5. **Formulas:** Use `include_formulas: true` to get formulas instead of values

6. **Excel Tables:** Use table names for structured data - easier column selection and automatic range updates

7. **Column references:** Use letters (A, B) or numbers (1, 2) for columns - tables also support header names

8. **Bulk operations:** Delete multiple columns at once by passing a list of header names to tables

---

## 🐛 If Something Goes Wrong

### Quick Fixes:

**"Cannot run macro"**
→ Check Excel Trust Settings (see SETUP_INSTRUCTIONS.md Step 1)

**"File is locked"**
→ Close the file in Excel manually

**"No MCP servers connected"**
→ Restart Claude Desktop, check config file

**Macro doesn't return value**
→ It might be a Sub (no return) instead of Function (has return)

---

**Ready to test?** Start with prompt #1 to verify everything is working!
