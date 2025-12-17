# Plan: Amélioration Opérations Tableaux Excel

**Date:** 2025-12-14
**Objectif:** Gérer les tableaux Excel structurés (Tables) et opérations colonnes/lignes
**Version:** 0.3.0

---

## Vue d'Ensemble

Améliorer les outils Excel pour travailler avec des **tableaux structurés** (Excel Tables) au lieu de simples ranges.

### Fonctionnalités à Ajouter

1. **Support des Tables Excel** (ListObjects)
   - Lire/écrire par nom de table au lieu de range
   - Accès par nom de colonne
   - Filtrage et tri

2. **Opérations sur Lignes**
   - Insérer lignes
   - Supprimer lignes
   - Copier/déplacer lignes

3. **Opérations sur Colonnes**
   - Insérer colonnes
   - Supprimer colonnes
   - Renommer colonnes

4. **Gestion Tables**
   - Lister tables disponibles
   - Créer nouveau tableau
   - Convertir range en table

---

## Fichiers à Modifier/Créer

### Fichiers à Modifier
1. `packages/pro/src/vba_mcp_pro/tools/office_automation.py`
   - Améliorer `get_worksheet_data_tool`
   - Améliorer `set_worksheet_data_tool`

### Fichiers à Créer
2. `packages/pro/src/vba_mcp_pro/tools/excel_tables.py` (NOUVEAU)
   - Tous les nouveaux outils Excel Tables

3. `packages/pro/tests/test_excel_tables.py` (NOUVEAU)
   - Tests pour les nouvelles fonctionnalités

---

## Phase 1: Améliorer Outils Existants (2h)

### 1.1 Améliorer `get_worksheet_data_tool`

**Nouvelles capacités:**

```python
async def get_worksheet_data_tool(
    file_path: str,
    sheet_name: str,
    range: Optional[str] = None,          # Existant: "A1:C10"
    table_name: Optional[str] = None,     # NOUVEAU: "BudgetTable"
    columns: Optional[List[str]] = None,  # NOUVEAU: ["Name", "Age"]
    include_headers: bool = True,         # NOUVEAU
    include_formulas: bool = False        # Existant
) -> str:
    """
    Read data from Excel worksheet, table, or specific columns.

    Examples:
        # By range (existing)
        get_worksheet_data(file, "Sheet1", range="A1:C10")

        # By table name (NEW)
        get_worksheet_data(file, "Sheet1", table_name="BudgetTable")

        # Specific columns from table (NEW)
        get_worksheet_data(file, "Sheet1", table_name="BudgetTable",
                          columns=["Name", "Total"])
    """
```

**Implémentation:**

```python
# Logique de détection
if table_name:
    # Utiliser ListObjects
    table = ws.ListObjects(table_name)
    data_range = table.DataBodyRange

    if columns:
        # Filtrer colonnes spécifiques
        header_range = table.HeaderRowRange
        col_indices = _find_column_indices(header_range, columns)
        data = _extract_columns(data_range, col_indices)
    else:
        data = data_range.Value

    if include_headers:
        headers = table.HeaderRowRange.Value
        data = [list(headers[0])] + data

elif range:
    # Comportement existant
    data = ws.Range(range).Value
else:
    # UsedRange (existant)
    data = ws.UsedRange.Value
```

**Helpers nécessaires:**

```python
def _find_column_indices(header_range, column_names: List[str]) -> List[int]:
    """Find column indices by name."""
    headers = [str(h).strip() for h in header_range.Value[0]]
    indices = []
    for col_name in column_names:
        try:
            idx = headers.index(col_name)
            indices.append(idx)
        except ValueError:
            raise ValueError(f"Column '{col_name}' not found. Available: {headers}")
    return indices

def _extract_columns(data_range, col_indices: List[int]) -> List[List]:
    """Extract specific columns from range."""
    data = data_range.Value
    if not isinstance(data, tuple):
        data = [[data]]

    result = []
    for row in data:
        result.append([row[i] for i in col_indices])
    return result
```

---

### 1.2 Améliorer `set_worksheet_data_tool`

**Nouvelles capacités:**

```python
async def set_worksheet_data_tool(
    file_path: str,
    sheet_name: str,
    data: List[List[Any]],
    start_cell: str = "A1",               # Existant
    table_name: Optional[str] = None,     # NOUVEAU
    column_mapping: Optional[Dict] = None, # NOUVEAU
    append: bool = False,                 # NOUVEAU
    clear_existing: bool = False          # Existant
) -> str:
    """
    Write data to Excel worksheet or table.

    Examples:
        # By range (existing)
        set_worksheet_data(file, "Sheet1", data, start_cell="A1")

        # Append to table (NEW)
        set_worksheet_data(file, "Sheet1", data, table_name="BudgetTable",
                          append=True)

        # Update specific columns (NEW)
        set_worksheet_data(file, "Sheet1", data, table_name="BudgetTable",
                          column_mapping={"Name": 0, "Age": 1})
    """
```

**Implémentation:**

```python
if table_name:
    table = ws.ListObjects(table_name)

    if append:
        # Ajouter à la fin du tableau
        last_row = table.ListRows.Count
        table.ListRows.Add()
        target_range = table.ListRows(last_row + 1).Range
        target_range.Value = tuple(data[0])

    elif column_mapping:
        # Mapper colonnes spécifiques
        headers = [str(h) for h in table.HeaderRowRange.Value[0]]
        for row_idx, row_data in enumerate(data, start=1):
            for col_name, data_idx in column_mapping.items():
                col_idx = headers.index(col_name)
                cell = table.DataBodyRange.Cells(row_idx, col_idx + 1)
                cell.Value = row_data[data_idx]
    else:
        # Remplacer tout le tableau
        table.DataBodyRange.Value = tuple(tuple(row) for row in data)

else:
    # Comportement existant (par range)
    # ... code existant ...
```

---

## Phase 2: Nouveaux Outils Tables (3h)

### Fichier: `packages/pro/src/vba_mcp_pro/tools/excel_tables.py`

### 2.1 `list_tables_tool`

```python
async def list_tables_tool(
    file_path: str,
    sheet_name: Optional[str] = None
) -> str:
    """
    List all Excel Tables (ListObjects) in file or specific sheet.

    Returns:
        JSON with table info: name, sheet, rows, columns, range
    """
    manager = OfficeSessionManager.get_instance()
    session = await manager.get_or_create_session(Path(file_path).resolve())

    tables_info = []

    if sheet_name:
        sheets = [session.file_obj.Worksheets(sheet_name)]
    else:
        sheets = session.file_obj.Worksheets

    for ws in sheets:
        for table in ws.ListObjects:
            info = {
                "name": table.Name,
                "sheet": ws.Name,
                "rows": table.ListRows.Count,
                "columns": table.ListColumns.Count,
                "headers": [str(h) for h in table.HeaderRowRange.Value[0]],
                "range": table.Range.Address,
                "total_row": table.ShowTotals
            }
            tables_info.append(info)

    # Format output
    output = f"### Excel Tables in {Path(file_path).name}\n\n"

    if not tables_info:
        output += "No tables found.\n"
        return output

    output += f"**Total:** {len(tables_info)} table(s)\n\n"

    for table in tables_info:
        output += f"#### {table['name']}\n"
        output += f"- **Sheet:** {table['sheet']}\n"
        output += f"- **Size:** {table['rows']} rows × {table['columns']} columns\n"
        output += f"- **Columns:** {', '.join(table['headers'])}\n"
        output += f"- **Range:** {table['range']}\n\n"

    return output
```

---

### 2.2 `insert_rows_tool`

```python
async def insert_rows_tool(
    file_path: str,
    sheet_name: str,
    position: int,
    count: int = 1,
    table_name: Optional[str] = None
) -> str:
    """
    Insert row(s) in worksheet or table.

    Args:
        position: Row number (1-based) or relative position in table
        count: Number of rows to insert
        table_name: If specified, insert in table context
    """
    manager = OfficeSessionManager.get_instance()
    session = await manager.get_or_create_session(Path(file_path).resolve())
    ws = session.file_obj.Worksheets(sheet_name)

    if table_name:
        # Insert in table
        table = ws.ListObjects(table_name)
        for i in range(count):
            table.ListRows.Add(Position=position + i)

        output = f"### Rows Inserted in Table\n\n"
        output += f"**Table:** {table_name}\n"
        output += f"**Position:** {position}\n"
        output += f"**Count:** {count}\n"
        output += f"**New size:** {table.ListRows.Count} rows\n"

    else:
        # Insert in worksheet
        for i in range(count):
            ws.Rows(position + i).Insert()

        output = f"### Rows Inserted in Worksheet\n\n"
        output += f"**Sheet:** {sheet_name}\n"
        output += f"**Position:** Row {position}\n"
        output += f"**Count:** {count}\n"

    return output
```

---

### 2.3 `delete_rows_tool`

```python
async def delete_rows_tool(
    file_path: str,
    sheet_name: str,
    start_row: int,
    end_row: Optional[int] = None,
    table_name: Optional[str] = None
) -> str:
    """
    Delete row(s) from worksheet or table.

    Args:
        start_row: First row to delete (1-based)
        end_row: Last row to delete (inclusive). If None, delete only start_row
        table_name: If specified, delete from table
    """
    if end_row is None:
        end_row = start_row

    count = end_row - start_row + 1

    manager = OfficeSessionManager.get_instance()
    session = await manager.get_or_create_session(Path(file_path).resolve())
    ws = session.file_obj.Worksheets(sheet_name)

    if table_name:
        table = ws.ListObjects(table_name)
        # Delete in reverse to avoid index shifting
        for i in range(end_row, start_row - 1, -1):
            table.ListRows(i).Delete()

        output = f"### Rows Deleted from Table\n\n"
        output += f"**Table:** {table_name}\n"
        output += f"**Rows:** {start_row} to {end_row} ({count} row(s))\n"
        output += f"**New size:** {table.ListRows.Count} rows\n"
    else:
        # Delete from worksheet
        range_to_delete = ws.Rows(f"{start_row}:{end_row}")
        range_to_delete.Delete()

        output = f"### Rows Deleted from Worksheet\n\n"
        output += f"**Sheet:** {sheet_name}\n"
        output += f"**Rows:** {start_row} to {end_row} ({count} row(s))\n"

    return output
```

---

### 2.4 `insert_columns_tool`

```python
async def insert_columns_tool(
    file_path: str,
    sheet_name: str,
    position: Union[int, str],  # 1 ou "A" ou "B"
    count: int = 1,
    table_name: Optional[str] = None,
    header_name: Optional[str] = None
) -> str:
    """
    Insert column(s) in worksheet or table.

    Args:
        position: Column number (1-based) or letter ("A", "B", etc.)
        count: Number of columns to insert
        table_name: If specified, insert in table
        header_name: New column header (for tables)
    """
    manager = OfficeSessionManager.get_instance()
    session = await manager.get_or_create_session(Path(file_path).resolve())
    ws = session.file_obj.Worksheets(sheet_name)

    # Convert column letter to number if needed
    if isinstance(position, str):
        position = _column_letter_to_number(position)

    if table_name:
        table = ws.ListObjects(table_name)

        for i in range(count):
            col = table.ListColumns.Add(Position=position + i)
            if header_name:
                col.Name = f"{header_name}_{i+1}" if count > 1 else header_name

        output = f"### Columns Inserted in Table\n\n"
        output += f"**Table:** {table_name}\n"
        output += f"**Position:** {position}\n"
        output += f"**Count:** {count}\n"
        if header_name:
            output += f"**Header:** {header_name}\n"
        output += f"**New size:** {table.ListColumns.Count} columns\n"
    else:
        # Insert in worksheet
        for i in range(count):
            ws.Columns(position + i).Insert()

        output = f"### Columns Inserted in Worksheet\n\n"
        output += f"**Sheet:** {sheet_name}\n"
        output += f"**Position:** Column {position}\n"
        output += f"**Count:** {count}\n"

    return output


def _column_letter_to_number(letter: str) -> int:
    """Convert column letter to number (A=1, B=2, ... Z=26, AA=27)."""
    num = 0
    for char in letter.upper():
        num = num * 26 + (ord(char) - ord('A') + 1)
    return num
```

---

### 2.5 `delete_columns_tool`

```python
async def delete_columns_tool(
    file_path: str,
    sheet_name: str,
    column: Union[int, str, List[str]],  # 1 ou "A" ou ["Name", "Age"]
    table_name: Optional[str] = None
) -> str:
    """
    Delete column(s) from worksheet or table.

    Args:
        column: Column number, letter, or list of column names (for tables)
        table_name: If specified, delete from table
    """
    manager = OfficeSessionManager.get_instance()
    session = await manager.get_or_create_session(Path(file_path).resolve())
    ws = session.file_obj.Worksheets(sheet_name)

    if table_name:
        table = ws.ListObjects(table_name)

        if isinstance(column, list):
            # Delete by column names
            for col_name in column:
                table.ListColumns(col_name).Delete()

            output = f"### Columns Deleted from Table\n\n"
            output += f"**Table:** {table_name}\n"
            output += f"**Columns:** {', '.join(column)} ({len(column)} column(s))\n"
        else:
            # Delete by position
            if isinstance(column, str):
                column = _column_letter_to_number(column)
            table.ListColumns(column).Delete()

            output = f"### Column Deleted from Table\n\n"
            output += f"**Table:** {table_name}\n"
            output += f"**Position:** {column}\n"

        output += f"**New size:** {table.ListColumns.Count} columns\n"
    else:
        # Delete from worksheet
        if isinstance(column, str) and not column.isdigit():
            column = _column_letter_to_number(column)

        ws.Columns(column).Delete()

        output = f"### Column Deleted from Worksheet\n\n"
        output += f"**Sheet:** {sheet_name}\n"
        output += f"**Column:** {column}\n"

    return output
```

---

### 2.6 `create_table_tool`

```python
async def create_table_tool(
    file_path: str,
    sheet_name: str,
    range: str,
    table_name: str,
    has_headers: bool = True,
    style: str = "TableStyleMedium2"
) -> str:
    """
    Convert a range to an Excel Table (ListObject).

    Args:
        range: Range to convert (e.g., "A1:D10")
        table_name: Name for the new table
        has_headers: First row contains headers
        style: Excel table style name
    """
    manager = OfficeSessionManager.get_instance()
    session = await manager.get_or_create_session(Path(file_path).resolve())
    ws = session.file_obj.Worksheets(sheet_name)

    # Check if table name already exists
    for existing_table in ws.ListObjects:
        if existing_table.Name == table_name:
            raise ValueError(f"Table '{table_name}' already exists in {sheet_name}")

    # Create table
    source_range = ws.Range(range)
    table = ws.ListObjects.Add(
        SourceType=1,  # xlSrcRange
        Source=source_range,
        XlListObjectHasHeaders=1 if has_headers else 2
    )
    table.Name = table_name
    table.TableStyle = style

    output = f"### Excel Table Created\n\n"
    output += f"**Name:** {table_name}\n"
    output += f"**Sheet:** {sheet_name}\n"
    output += f"**Range:** {range}\n"
    output += f"**Headers:** {'Yes' if has_headers else 'No'}\n"
    output += f"**Size:** {table.ListRows.Count} rows × {table.ListColumns.Count} columns\n"

    if has_headers:
        headers = [str(h) for h in table.HeaderRowRange.Value[0]]
        output += f"**Columns:** {', '.join(headers)}\n"

    return output
```

---

## Phase 3: Enregistrement MCP (1h)

### Modifier `server.py`

Ajouter 8 nouveaux outils:

```python
from .tools.excel_tables import (
    list_tables_tool,
    insert_rows_tool,
    delete_rows_tool,
    insert_columns_tool,
    delete_columns_tool,
    create_table_tool
)

# Dans list_tools():
Tool(
    name="list_tables",
    description="[PRO] List all Excel Tables (ListObjects) in a file.",
    inputSchema={...}
),
Tool(
    name="insert_rows",
    description="[PRO] Insert row(s) in worksheet or Excel table.",
    inputSchema={...}
),
Tool(
    name="delete_rows",
    description="[PRO] Delete row(s) from worksheet or Excel table.",
    inputSchema={...}
),
Tool(
    name="insert_columns",
    description="[PRO] Insert column(s) in worksheet or Excel table.",
    inputSchema={...}
),
Tool(
    name="delete_columns",
    description="[PRO] Delete column(s) from worksheet or Excel table.",
    inputSchema={...}
),
Tool(
    name="create_table",
    description="[PRO] Convert a range to an Excel Table (ListObject).",
    inputSchema={...}
)
```

**Total outils après Phase 3:** 14 → 20 outils (+6)

---

## Phase 4: Tests (2h)

### Fichier: `packages/pro/tests/test_excel_tables.py`

**Classes de tests:**

1. **TestTableSupport** - get/set avec tables
2. **TestListTables** - Lister tables
3. **TestRowOperations** - Insert/delete rows
4. **TestColumnOperations** - Insert/delete columns
5. **TestCreateTable** - Créer tables
6. **TestColumnMapping** - Mapper colonnes par nom

**~30 tests au total**

---

## Phase 5: Documentation (1h)

**Fichiers à créer/modifier:**

1. `packages/pro/CHANGELOG.md` - Version 0.3.0
2. `packages/pro/README.md` - Nouvelles features
3. `QUICK_TEST_PROMPTS.md` - Tests tables Excel
4. `../vba-mcp-demo/PROMPTS_READY_TO_USE.md` - Workflows tables

---

## Exemples d'Usage

### Scénario 1: Travailler avec Tables

```python
# 1. Lister les tables disponibles
list_tables("budget.xlsm")
# → BudgetTable, ExpensesTable

# 2. Lire toute la table
get_worksheet_data("budget.xlsm", "Sheet1", table_name="BudgetTable")

# 3. Lire seulement certaines colonnes
get_worksheet_data("budget.xlsm", "Sheet1",
                  table_name="BudgetTable",
                  columns=["Category", "Total"])

# 4. Ajouter une ligne à la table
set_worksheet_data("budget.xlsm", "Sheet1",
                  data=[["New Item", 150.00]],
                  table_name="BudgetTable",
                  append=True)

# 5. Insérer une colonne
insert_columns("budget.xlsm", "Sheet1",
              position=3,
              table_name="BudgetTable",
              header_name="Notes")
```

### Scénario 2: Créer et Gérer Tables

```python
# 1. Créer une nouvelle table
create_table("data.xlsm", "Sheet1",
            range="A1:D100",
            table_name="SalesData",
            has_headers=True)

# 2. Insérer des lignes
insert_rows("data.xlsm", "Sheet1",
           position=5,
           count=3,
           table_name="SalesData")

# 3. Supprimer colonnes par nom
delete_columns("data.xlsm", "Sheet1",
              column=["TempCol1", "TempCol2"],
              table_name="SalesData")
```

---

## Temps Total Estimé

| Phase | Durée |
|-------|-------|
| Phase 1: Améliorer outils existants | 2h |
| Phase 2: Nouveaux outils | 3h |
| Phase 3: Enregistrement MCP | 1h |
| Phase 4: Tests | 2h |
| Phase 5: Documentation | 1h |
| **TOTAL** | **9h** (~1.5 jour)

---

## Checklist d'Implémentation

### Phase 1
- [ ] Améliorer get_worksheet_data_tool (table_name, columns)
- [ ] Améliorer set_worksheet_data_tool (append, column_mapping)
- [ ] Ajouter helpers (_find_column_indices, _extract_columns)

### Phase 2
- [ ] Créer excel_tables.py
- [ ] Implémenter list_tables_tool
- [ ] Implémenter insert_rows_tool
- [ ] Implémenter delete_rows_tool
- [ ] Implémenter insert_columns_tool
- [ ] Implémenter delete_columns_tool
- [ ] Implémenter create_table_tool

### Phase 3
- [ ] Mettre à jour __init__.py (exports)
- [ ] Mettre à jour server.py (6 nouveaux tools)
- [ ] Vérifier tous les schemas MCP

### Phase 4
- [ ] Créer test_excel_tables.py
- [ ] Tests pour table support
- [ ] Tests pour row operations
- [ ] Tests pour column operations
- [ ] Tests pour create table

### Phase 5
- [ ] Mettre à jour CHANGELOG.md (v0.3.0)
- [ ] Mettre à jour README.md
- [ ] Ajouter tests dans QUICK_TEST_PROMPTS.md
- [ ] Ajouter workflows dans PROMPTS_READY_TO_USE.md

---

**Prêt pour implémentation!**
