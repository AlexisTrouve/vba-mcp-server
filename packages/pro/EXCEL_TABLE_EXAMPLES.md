# Excel Table Operations - Examples

## New Features Added (Phase 1)

The `get_worksheet_data_tool` and `set_worksheet_data_tool` have been enhanced to support Excel Tables (ListObjects).

---

## Helper Functions Added

### 1. `_column_letter_to_number(letter: str) -> int`
Converts Excel column letters to numbers:
- "A" → 1
- "Z" → 26
- "AA" → 27
- "AB" → 28

### 2. `_find_column_indices(header_range, column_names: List[str]) -> List[int]`
Finds column indices by name from a header range.

**Raises:**
- `ValueError` if any column name is not found (lists available columns)

### 3. `_extract_columns(data_range, col_indices: List[int]) -> List[List]`
Extracts specific columns from a data range.

---

## Enhanced: `get_worksheet_data_tool`

### New Parameters:
- `table_name` (Optional[str]): Excel Table name (e.g., "BudgetTable")
- `columns` (Optional[List[str]]): List of column names to extract
- `include_headers` (bool): Include header row in output (default: True)

### Usage Examples:

#### Example 1: Read entire table
```python
get_worksheet_data(
    file_path="/path/to/budget.xlsm",
    sheet_name="Sheet1",
    table_name="BudgetTable"
)
```

**Output:**
```json
[
  ["Category", "Amount", "Date"],
  ["Groceries", 150.00, "2025-01-01"],
  ["Transport", 75.50, "2025-01-02"],
  ["Entertainment", 100.00, "2025-01-03"]
]
```

#### Example 2: Read specific columns from table
```python
get_worksheet_data(
    file_path="/path/to/budget.xlsm",
    sheet_name="Sheet1",
    table_name="BudgetTable",
    columns=["Category", "Amount"]
)
```

**Output:**
```json
[
  ["Category", "Amount"],
  ["Groceries", 150.00],
  ["Transport", 75.50],
  ["Entertainment", 100.00]
]
```

#### Example 3: Read table without headers
```python
get_worksheet_data(
    file_path="/path/to/budget.xlsm",
    sheet_name="Sheet1",
    table_name="BudgetTable",
    include_headers=False
)
```

**Output:**
```json
[
  ["Groceries", 150.00, "2025-01-01"],
  ["Transport", 75.50, "2025-01-02"],
  ["Entertainment", 100.00, "2025-01-03"]
]
```

#### Example 4: Traditional range mode (still works)
```python
get_worksheet_data(
    file_path="/path/to/budget.xlsm",
    sheet_name="Sheet1",
    range="A1:C10"
)
```

---

## Enhanced: `set_worksheet_data_tool`

### New Parameters:
- `table_name` (Optional[str]): Excel Table name to write to
- `column_mapping` (Optional[dict]): Map data columns to table columns (e.g., `{"Name": 0, "Age": 1}`)
- `append` (bool): Append rows to end of table (requires `table_name`, default: False)

### Usage Examples:

#### Example 1: Append rows to table
```python
set_worksheet_data(
    file_path="/path/to/budget.xlsm",
    sheet_name="Sheet1",
    data=[
        ["Utilities", 200.00, "2025-01-04"],
        ["Healthcare", 350.00, "2025-01-05"]
    ],
    table_name="BudgetTable",
    append=True
)
```

**Result:**
- Adds 2 new rows to the end of "BudgetTable"
- Table automatically expands
- Returns new table size

#### Example 2: Update specific columns with column mapping
```python
set_worksheet_data(
    file_path="/path/to/employees.xlsm",
    sheet_name="Staff",
    data=[
        ["John Doe", 35],
        ["Jane Smith", 28],
        ["Bob Johnson", 42]
    ],
    table_name="EmployeeTable",
    column_mapping={
        "Name": 0,      # data[i][0] → Name column
        "Age": 1        # data[i][1] → Age column
    }
)
```

**Result:**
- Updates only "Name" and "Age" columns
- Other columns (e.g., "Salary", "Department") remain unchanged
- Creates new rows if needed

#### Example 3: Replace entire table data
```python
set_worksheet_data(
    file_path="/path/to/budget.xlsm",
    sheet_name="Sheet1",
    data=[
        ["Food", 100.00, "2025-01-10"],
        ["Gas", 50.00, "2025-01-11"]
    ],
    table_name="BudgetTable"
)
```

**Result:**
- Clears all existing data in table
- Writes new data
- Table structure (columns) preserved

#### Example 4: Traditional range mode (still works)
```python
set_worksheet_data(
    file_path="/path/to/budget.xlsm",
    sheet_name="Sheet1",
    data=[["A", "B"], ["C", "D"]],
    start_cell="A1"
)
```

---

## Error Handling

### Table Not Found
```python
ValueError: Table 'InvalidTable' not found in sheet 'Sheet1'
Available tables: ['BudgetTable', 'ExpensesTable']
```

### Column Not Found
```python
ValueError: Column 'InvalidColumn' not found.
Available columns: ['Category', 'Amount', 'Date']
```

### Column Mapping Error
```python
ValueError: Column 'Salary' not found in table.
Available columns: ['Name', 'Age', 'Department']
```

---

## Compatibility

### Backward Compatible
All existing code using `get_worksheet_data_tool` and `set_worksheet_data_tool` continues to work:
- Range-based operations unchanged
- Default parameters preserve existing behavior
- No breaking changes

### New Features Available
- Use `table_name` to work with Excel Tables
- Use `columns` to filter specific columns by name
- Use `append=True` to add rows incrementally
- Use `column_mapping` for partial updates

---

## Performance Notes

1. **Table operations are efficient:**
   - Direct access to ListObjects
   - No need to calculate ranges manually

2. **Column filtering happens in Python:**
   - Extracts only requested columns
   - Reduces data transfer

3. **Append mode is optimized:**
   - Uses `ListRows.Add()` for each row
   - Automatic table expansion

4. **Calculation mode disabled during writes:**
   - Manual calculation for performance
   - Recalculates once after write completes

---

## Next Steps (Phase 2)

The plan includes creating a new file `excel_tables.py` with additional tools:
- `list_tables_tool` - List all tables in a workbook
- `insert_rows_tool` - Insert rows into tables
- `delete_rows_tool` - Delete rows from tables
- `insert_columns_tool` - Insert columns into tables
- `delete_columns_tool` - Delete columns from tables
- `create_table_tool` - Convert range to Excel Table

---

## Testing

To test these features, you'll need:
1. An Excel file (.xlsm or .xlsx) with a defined Table (ListObject)
2. The table should have headers
3. Windows environment with Microsoft Excel installed

Example test:
```python
# 1. Open file
open_in_office("C:/path/to/budget.xlsm")

# 2. Read entire table
data = get_worksheet_data(
    "C:/path/to/budget.xlsm",
    "Sheet1",
    table_name="BudgetTable"
)

# 3. Append new row
set_worksheet_data(
    "C:/path/to/budget.xlsm",
    "Sheet1",
    data=[["New Item", 99.99, "2025-01-15"]],
    table_name="BudgetTable",
    append=True
)

# 4. Close file
close_office_file("C:/path/to/budget.xlsm")
```
