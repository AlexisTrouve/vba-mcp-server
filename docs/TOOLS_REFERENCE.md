# VBA MCP Server - Tools Reference

Complete reference for all 24 MCP tools available in VBA MCP Server.

## Quick Reference

| Tool | Package | Description |
|------|---------|-------------|
| `extract_vba` | Lite | Extract VBA code from modules |
| `list_modules` | Lite | List all VBA modules |
| `analyze_structure` | Lite | Analyze code structure |
| `inject_vba` | Pro | Inject VBA code into modules |
| `validate_vba` | Pro | Validate VBA syntax |
| `run_macro` | Pro | Execute VBA macros |
| `open_in_office` | Pro | Open file in Office application |
| `get_worksheet_data` | Pro | Read data from Excel/Access |
| `set_worksheet_data` | Pro | Write data to Excel/Access |
| `list_excel_tables` | Pro | List Excel Tables |
| `create_excel_table` | Pro | Create new Excel Table |
| `insert_rows` | Pro | Insert rows in worksheet/table |
| `delete_rows` | Pro | Delete rows |
| `insert_columns` | Pro | Insert columns |
| `delete_columns` | Pro | Delete columns |
| `create_backup` | Pro | Create file backup |
| `restore_backup` | Pro | Restore from backup |
| `list_backups` | Pro | List available backups |
| `list_access_tables` | Pro | List Access tables with schema |
| `list_access_queries` | Pro | List saved queries |
| `run_access_query` | Pro | Execute SQL/saved queries |

---

## Lite Tools (Free - MIT License)

### extract_vba

Extract VBA source code from a specific module.

**Parameters:**
- `file_path` (string, required): Path to Office file
- `module_name` (string, required): Name of module to extract

**Returns:** VBA source code as text

**Supported files:** .xlsm, .xlsb, .docm, .accdb

---

### list_modules

List all VBA modules in a file.

**Parameters:**
- `file_path` (string, required): Path to Office file

**Returns:** List of module names with types (Standard, Class, Form, Document)

---

### analyze_structure

Analyze VBA code structure and complexity.

**Parameters:**
- `file_path` (string, required): Path to Office file
- `module_name` (string, optional): Specific module to analyze

**Returns:**
- Procedure list (Sub, Function)
- Line counts
- Complexity metrics
- Dependencies

---

## Pro Tools (Commercial License)

### VBA Manipulation

#### inject_vba

Inject VBA code into a module (creates if doesn't exist).

**Parameters:**
- `file_path` (string): Path to Office file
- `module_name` (string): Target module name
- `code` (string): VBA code to inject

**Features:**
- Automatic backup before injection
- Syntax validation
- Rollback on error

---

#### validate_vba

Validate VBA syntax without injecting.

**Parameters:**
- `code` (string): VBA code to validate
- `file_type` (string): "excel", "word", or "access"

**Returns:**
- Valid/Invalid status
- Error details if invalid
- Line numbers for errors

---

#### run_macro

Execute a VBA macro (Sub or Function).

**Parameters:**
- `file_path` (string): Path to Office file
- `macro_name` (string): Name of macro to run
- `args` (array, optional): Arguments to pass

**Returns:** Function return value or execution status

---

### Office Automation

#### open_in_office

Open a file in its Office application (visible).

**Parameters:**
- `file_path` (string): Path to Office file
- `read_only` (boolean, optional): Open read-only

**Returns:** Session ID for subsequent operations

---

### Excel Data Operations

#### get_worksheet_data

Read data from Excel worksheet or Access table.

**Parameters:**
- `file_path` (string): Path to file
- `sheet_name` (string): Worksheet or table name
- `range` (string, optional): Cell range (Excel only)
- `where_clause` (string, optional): Filter condition (Access)
- `order_by` (string, optional): Sort order (Access)
- `limit` (integer, optional): Max rows

**Returns:** Data as JSON array

---

#### set_worksheet_data

Write data to Excel worksheet or Access table.

**Parameters:**
- `file_path` (string): Path to file
- `sheet_name` (string): Worksheet or table name
- `data` (array): 2D array of values
- `start_cell` (string, optional): Starting cell (Excel, default "A1")
- `columns` (array, optional): Column names (Access)
- `mode` (string, optional): "append" or "replace" (Access)

---

### Excel Tables

#### list_excel_tables

List all Excel Tables (ListObjects) in a workbook.

**Parameters:**
- `file_path` (string): Path to Excel file

**Returns:**
- Table names
- Worksheet locations
- Row/column counts
- Header names

---

#### create_excel_table

Create a new Excel Table from a range.

**Parameters:**
- `file_path` (string): Path to Excel file
- `sheet_name` (string): Worksheet name
- `range` (string): Data range (e.g., "A1:D10")
- `table_name` (string): Name for new table
- `has_headers` (boolean, optional): First row is headers

---

#### insert_rows

Insert rows in a worksheet or table.

**Parameters:**
- `file_path` (string): Path to Excel file
- `sheet_name` (string): Worksheet name
- `position` (integer): Row number to insert at
- `count` (integer, optional): Number of rows (default 1)
- `table_name` (string, optional): Insert in specific table

---

#### delete_rows

Delete rows from worksheet or table.

**Parameters:**
- `file_path` (string): Path to Excel file
- `sheet_name` (string): Worksheet name
- `start_row` (integer): First row to delete
- `count` (integer, optional): Number of rows
- `table_name` (string, optional): Delete from specific table

---

#### insert_columns

Insert columns in worksheet or table.

**Parameters:**
- `file_path` (string): Path to Excel file
- `sheet_name` (string): Worksheet name
- `position` (string/integer): Column letter or number
- `count` (integer, optional): Number of columns

---

#### delete_columns

Delete columns from worksheet or table.

**Parameters:**
- `file_path` (string): Path to Excel file
- `sheet_name` (string): Worksheet name
- `start_column` (string/integer): Column letter or number
- `count` (integer, optional): Number of columns

---

### Backup System

#### create_backup

Create a backup of an Office file.

**Parameters:**
- `file_path` (string): Path to file to backup

**Returns:** Backup file path with timestamp

---

#### restore_backup

Restore a file from backup.

**Parameters:**
- `file_path` (string): Original file path
- `backup_path` (string, optional): Specific backup to restore

**Returns:** Restoration status

---

#### list_backups

List all available backups for a file.

**Parameters:**
- `file_path` (string): Original file path

**Returns:** List of backups with timestamps and sizes

---

### Access-Specific Tools

#### list_access_tables

List all tables in Access database with schema.

**Parameters:**
- `file_path` (string): Path to .accdb file

**Returns:**
- Table names
- Field names, types, sizes
- Record counts

---

#### list_access_queries

List saved queries (QueryDefs).

**Parameters:**
- `file_path` (string): Path to .accdb file

**Returns:**
- Query names
- Query types
- SQL preview

---

#### run_access_query

Execute SQL or saved query.

**Parameters:**
- `file_path` (string): Path to .accdb file
- `query_name` (string, optional): Saved query name
- `sql` (string, optional): Custom SQL

**Supports:** SELECT, INSERT, UPDATE, DELETE

---

## Error Handling

All tools return structured error messages:

```json
{
  "error": true,
  "message": "Description of error",
  "details": "Technical details"
}
```

Common errors:
- File not found
- File locked by another process
- VBA project access denied
- Syntax validation failed
- COM automation error

---

## Best Practices

1. **Always backup before modifications**
   ```
   create_backup -> inject_vba -> verify
   ```

2. **Validate before injecting**
   ```
   validate_vba -> inject_vba
   ```

3. **Close files in Office before operations**
   ```
   File locked? -> Close in Office -> Retry
   ```

4. **Use read_only when just reading**
   ```
   get_worksheet_data (automatically read-only)
   ```

---

**Version:** 0.6.0
**Last Updated:** 2025-12-30
