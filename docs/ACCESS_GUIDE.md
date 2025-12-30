# Microsoft Access Guide - VBA MCP Pro

Complete guide for using VBA MCP Server with Microsoft Access databases (.accdb).

## Table of Contents

- [Overview](#overview)
- [Supported Features](#supported-features)
- [MCP Tools Reference](#mcp-tools-reference)
- [Examples](#examples)
- [Differences from Excel](#differences-from-excel)
- [Known Limitations](#known-limitations)
- [Troubleshooting](#troubleshooting)

---

## Overview

VBA MCP Pro v0.6.0 provides full support for Microsoft Access databases, including:

- **VBA Code**: Extract, inject, and validate VBA code
- **Data Operations**: Read and write table data with SQL support
- **Schema Discovery**: List tables, fields, types, and record counts
- **Query Execution**: Run saved queries or custom SQL

---

## Supported Features

| Feature | Status | Tool |
|---------|--------|------|
| List tables with schema | 100% | `list_access_tables` |
| List saved queries | 100% | `list_access_queries` |
| SELECT queries | 100% | `run_access_query` |
| INSERT queries | 100% | `run_access_query` |
| UPDATE queries | 100% | `run_access_query` |
| DELETE queries | 100% | `run_access_query` |
| Run saved queries | 100% | `run_access_query` |
| Read data with filters | 100% | `get_worksheet_data` |
| Write data (append) | 100% | `set_worksheet_data` |
| Write data (replace) | 100% | `set_worksheet_data` |
| Extract VBA | 100% | `extract_vba` |
| Inject VBA | 100% | `inject_vba` |
| Validate VBA | 100% | `validate_vba` |
| Backup/Restore | 100% | `create_backup`, `restore_backup` |

---

## MCP Tools Reference

### list_access_tables

List all tables in an Access database with full schema information.

**Parameters:**
- `file_path` (string): Path to .accdb file

**Returns:**
- Table names
- Field names, types, and sizes
- Record counts

**Example:**
```
List all tables in C:\data\customers.accdb
```

---

### list_access_queries

List all saved queries (QueryDefs) in the database.

**Parameters:**
- `file_path` (string): Path to .accdb file

**Returns:**
- Query names
- Query types (SELECT, INSERT, UPDATE, DELETE, etc.)
- SQL preview

**Example:**
```
Show me all queries in the database
```

---

### run_access_query

Execute SQL queries or saved queries.

**Parameters:**
- `file_path` (string): Path to .accdb file
- `query_name` (optional): Name of saved query to run
- `sql` (optional): Custom SQL to execute

**Supported SQL:**
- `SELECT` - Returns data as JSON
- `INSERT` - Returns rows affected
- `UPDATE` - Returns rows affected
- `DELETE` - Returns rows affected

**Examples:**
```sql
-- Run saved query
run_access_query(file, query_name="ActiveCustomers")

-- Custom SELECT
run_access_query(file, sql="SELECT * FROM Customers WHERE Country = 'France'")

-- INSERT
run_access_query(file, sql="INSERT INTO Logs (Message) VALUES ('Test')")

-- UPDATE
run_access_query(file, sql="UPDATE Products SET Price = Price * 1.1 WHERE Category = 'Electronics'")

-- DELETE
run_access_query(file, sql="DELETE FROM TempData WHERE CreatedAt < #2024-01-01#")
```

---

### get_worksheet_data

Read data from an Access table with optional filtering.

**Parameters:**
- `file_path` (string): Path to .accdb file
- `sheet_name` (string): Table name
- `where_clause` (optional): SQL WHERE condition
- `order_by` (optional): SQL ORDER BY clause
- `limit` (optional): Maximum rows to return

**Examples:**
```
Read the Customers table from database.accdb

Read top 100 orders sorted by date descending

Read employees where Department = 'IT'
```

---

### set_worksheet_data

Write data to an Access table.

**Parameters:**
- `file_path` (string): Path to .accdb file
- `sheet_name` (string): Table name
- `data` (array): 2D array of values
- `columns` (optional): Column names for the data
- `mode` (string): "append" or "replace"

**Modes:**
- `append` - Add new records to existing data
- `replace` - Delete all existing records, then insert new ones

**Example:**
```python
# Append new customers
set_worksheet_data(
    file_path="database.accdb",
    sheet_name="Customers",
    data=[
        ["John", "Doe", "john@email.com"],
        ["Jane", "Smith", "jane@email.com"]
    ],
    columns=["FirstName", "LastName", "Email"],
    mode="append"
)
```

---

### inject_vba

Inject VBA code into an Access database module.

**Parameters:**
- `file_path` (string): Path to .accdb file
- `module_name` (string): Name for the VBA module
- `code` (string): VBA code to inject

**Example:**
```vba
' Inject this code into MyModule
Sub ProcessData()
    Dim db As Database
    Set db = CurrentDb()
    db.Execute "UPDATE Stats SET LastRun = Now()"
End Sub
```

---

### validate_vba

Validate VBA syntax before injection.

**Parameters:**
- `code` (string): VBA code to validate
- `file_type` (string): "access"

**Returns:**
- Validation result (valid/invalid)
- Error messages if any

---

## Examples

### Example 1: Analyze Database Schema

```
1. List all tables: list_access_tables("database.accdb")
2. List all queries: list_access_queries("database.accdb")
3. Preview data: get_worksheet_data("database.accdb", "Customers", limit=10)
```

### Example 2: Data Migration

```
1. Read source data: get_worksheet_data("source.accdb", "Products")
2. Write to destination: set_worksheet_data("dest.accdb", "Products", data, mode="replace")
```

### Example 3: Add VBA Automation

```
1. Validate code: validate_vba(code, file_type="access")
2. Create backup: create_backup("database.accdb")
3. Inject VBA: inject_vba("database.accdb", "AutomationModule", code)
```

---

## Differences from Excel

| Aspect | Excel | Access |
|--------|-------|--------|
| File extension | .xlsm, .xlsb | .accdb |
| Data structure | Worksheets/Cells | Tables/Records |
| Query support | Limited | Full SQL |
| VBA Project access | `Workbook.VBProject` | `VBE.ActiveVBProject` |
| Save behavior | Manual | Auto-save |
| run_macro | Works | Limited (known limitation) |

### Key Differences in COM Automation

```python
# Excel
app = Dispatch("Excel.Application")
workbook = app.Workbooks.Open(path)
vb_project = workbook.VBProject
workbook.Save()
workbook.Close()

# Access
app = Dispatch("Access.Application")
app.OpenCurrentDatabase(path)
vb_project = app.VBE.ActiveVBProject
# No explicit save needed - Access auto-saves
app.CloseCurrentDatabase()
```

---

## Known Limitations

### 1. oletools doesn't support .accdb

The `list_modules` tool uses oletools which doesn't support Access database files.

**Workaround:** Use COM via session manager for VBA extraction. The `extract_vba` and `inject_vba` tools work correctly.

### 2. run_macro behaves differently

Access `Application.Run` doesn't support VBA procedures the same way as Excel. Access macros (created via UI) are different from VBA code.

**Workaround:** Use `inject_vba` to add callable code, then execute via other means.

### 3. DisplayAlerts property

Access doesn't have the `DisplayAlerts` property like Excel. This is non-blocking and just produces a warning.

### 4. Concurrent Access

Access databases have file-level locking. Only one process can write at a time.

**Best Practice:** Close the database in Access UI before using MCP tools.

---

## Troubleshooting

### "File is locked" Error

**Cause:** Database is open in Access or another process.

**Solution:**
```bash
# Kill Access process
taskkill /IM msaccess.exe /F
```

### "Trust access to VBA project" Error

**Cause:** Access security settings block VBA manipulation.

**Solution:**
1. Open Access
2. File > Options > Trust Center > Trust Center Settings
3. Macro Settings > Check "Trust access to the VBA project object model"

### SQL Syntax Errors

**Access SQL differences:**
- Use `#date#` for dates, not `'date'`
- Use `"string"` or `'string'` for text
- Wildcards: `*` instead of `%`, `?` instead of `_`

**Examples:**
```sql
-- Date comparison
SELECT * FROM Orders WHERE OrderDate > #2024-01-01#

-- Wildcard search
SELECT * FROM Products WHERE Name LIKE '*widget*'
```

---

## Requirements

- Windows OS (COM automation required)
- Microsoft Access installed
- pywin32 package
- "Trust access to VBA project object model" enabled

---

**Version:** 0.6.0
**Last Updated:** 2025-12-30
