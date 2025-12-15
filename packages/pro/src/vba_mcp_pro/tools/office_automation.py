"""
Office Automation Tools (PRO)

Interactive Office automation tools that work with visible Office applications.
Files remain open between operations for faster interaction.
"""

import json
import logging
from pathlib import Path
from typing import Any, List, Optional
import platform

from ..session_manager import OfficeSessionManager


# Configure logging
logger = logging.getLogger(__name__)


class AutomationSecurityContext:
    """
    Context manager for temporarily lowering Office AutomationSecurity.
    Allows macro execution without user prompts.
    ALWAYS restores original security level on exit.
    """

    def __init__(self, app, target_level: int = 1):
        """
        Initialize AutomationSecurity context manager.

        Args:
            app: Office Application COM object
            target_level: Security level to set during execution
                1 = msoAutomationSecurityLow (macros enabled)
                2 = msoAutomationSecurityByUI (user setting)
                3 = msoAutomationSecurityForceDisable (disabled)
        """
        self.app = app
        self.target_level = target_level
        self.original_level = None

    def __enter__(self):
        """Lower AutomationSecurity to target level."""
        try:
            self.original_level = self.app.AutomationSecurity
            logger.info(f"AutomationSecurity: {self.original_level} → {self.target_level}")
            self.app.AutomationSecurity = self.target_level
        except Exception as e:
            logger.warning(f"Cannot modify AutomationSecurity: {e}")
            # Continue anyway - some applications may not support this property
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Restore original AutomationSecurity level."""
        if self.original_level is not None:
            try:
                self.app.AutomationSecurity = self.original_level
                logger.info(f"AutomationSecurity restored: {self.original_level}")
            except Exception as e:
                logger.warning(f"Cannot restore AutomationSecurity: {e}")
        # Don't suppress exceptions
        return False


# Maximum cells to read/write to prevent memory issues
MAX_CELLS = 1_000_000


async def open_in_office_tool(
    file_path: str,
    read_only: bool = False
) -> str:
    """
    Open an Office file interactively with visible UI.

    The file remains open for further operations (run macros, read/write data).
    Sessions automatically close after 1 hour of inactivity or when server stops.

    Args:
        file_path: Absolute path to Office file (.xlsm, .docm, .accdb, etc.)
        read_only: Open in read-only mode (default: false)

    Returns:
        Success message with file and application details

    Raises:
        FileNotFoundError: If file doesn't exist
        RuntimeError: If not on Windows or missing dependencies
        PermissionError: If file is locked
    """
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    # Platform check
    if platform.system() != "Windows":
        raise RuntimeError(
            "Office automation is only supported on Windows. "
            "Install this package on a Windows machine with Microsoft Office."
        )

    # Get session manager
    manager = OfficeSessionManager.get_instance()

    # Create/get session (this opens the file visibly)
    session = await manager.get_or_create_session(path, read_only=read_only)

    # Format result
    return "\n".join([
        "**Office File Opened**",
        "",
        f"File: {path.name}",
        f"Path: {path}",
        f"Application: {session.app_type}",
        f"Mode: {'Read-only' if read_only else 'Editable'}",
        f"Window: Visible",
        "",
        "The file remains open for interactive use.",
        "You can now:",
        "- Run macros with run_macro",
        "- Read data with get_worksheet_data",
        "- Write data with set_worksheet_data",
        "- Close with close_office_file (or wait for auto-close after 1h idle)",
    ])


async def close_office_file_tool(
    file_path: str,
    save_changes: bool = True
) -> str:
    """
    Close an open Office file session.

    Args:
        file_path: Absolute path to Office file
        save_changes: Save changes before closing (default: true)

    Returns:
        Success message

    Raises:
        ValueError: If no session exists for this file
    """
    path = Path(file_path)
    manager = OfficeSessionManager.get_instance()

    await manager.close_session(path, save=save_changes)

    return "\n".join([
        "**Office File Closed**",
        "",
        f"File: {path.name}",
        f"Changes: {'Saved' if save_changes else 'Not saved'}",
        "",
        "Session closed. Application quit."
    ])


async def list_open_files_tool() -> str:
    """
    List all currently open Office file sessions.

    Returns:
        List of open files with session details
    """
    manager = OfficeSessionManager.get_instance()
    sessions = manager.list_sessions()

    if not sessions:
        return "**No Office Files Currently Open**"

    lines = [
        "**Open Office File Sessions**",
        "",
        f"Total sessions: {len(sessions)}",
        ""
    ]

    for session_info in sessions:
        age_min = session_info["age_seconds"] // 60
        idle_min = session_info["last_accessed_seconds"] // 60

        lines.append(
            f"- **{session_info['file_name']}** ({session_info['app_type']})\n"
            f"  - Mode: {'Read-only' if session_info['read_only'] else 'Editable'}\n"
            f"  - Open for: {age_min} minutes\n"
            f"  - Idle for: {idle_min} minutes\n"
            f"  - Path: {session_info['file_path']}"
        )

    return "\n".join(lines)


async def run_macro_tool(
    file_path: str,
    macro_name: str,
    arguments: Optional[List[Any]] = None,
    enable_macros: bool = True
) -> str:
    """
    Execute a VBA macro in an Office file.

    The file will be opened if not already open.
    Supports both Sub procedures (no return) and Functions (with return value).

    Args:
        file_path: Absolute path to Office file
        macro_name: Macro name in format 'ModuleName.ProcedureName' or 'ProcedureName'
        arguments: Arguments to pass to the macro (optional)
        enable_macros: Enable macros by temporarily lowering AutomationSecurity (default: true)

    Returns:
        Execution result with return value (if Function)

    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If macro not found or argument mismatch
        RuntimeError: If execution fails
    """
    path = Path(file_path)
    manager = OfficeSessionManager.get_instance()

    # Get or create session
    session = await manager.get_or_create_session(path)
    session.refresh_last_accessed()

    # Import COM error handling
    try:
        import pythoncom
    except ImportError:
        raise RuntimeError("pywin32 required for macro execution")

    # Prepare arguments
    args = arguments or []

    # Parse macro name
    parts = macro_name.split('.')
    if len(parts) == 2:
        module_name, proc_name = parts
    else:
        module_name = None
        proc_name = macro_name

    # Try different formats
    formats_to_try = []

    if session.app_type == "Excel":
        workbook_name = session.file_obj.Name
        if module_name:
            formats_to_try = [
                f"{module_name}.{proc_name}",  # Module.Macro
                f"'{workbook_name}'!{module_name}.{proc_name}",  # 'Book.xlsm'!Module.Macro
                proc_name,  # Just macro name
            ]
        else:
            formats_to_try = [
                proc_name,  # Just macro name
                f"'{workbook_name}'!{proc_name}",  # 'Book.xlsm'!Macro
            ]
    elif session.app_type in ["Word", "Access"]:
        # Word/Access typically just use macro name
        formats_to_try = [
            proc_name,
            f"{module_name}.{proc_name}" if module_name else proc_name
        ]

    # Try each format
    last_error = None

    # Create security context manager
    if enable_macros:
        security_context = AutomationSecurityContext(session.app, target_level=1)
    else:
        from contextlib import nullcontext
        security_context = nullcontext()

    # Execute macro with appropriate security context
    with security_context:
        for format_name in formats_to_try:
            try:
                result = session.app.Run(format_name, *args)

                # Build success message
                output = "**Macro Executed Successfully**\n\n"
                output += f"File: {path.name}\n"
                output += f"Macro: {macro_name}\n"
                output += f"Format used: {format_name}\n"
                output += f"Arguments: {args}\n"

                if result is not None:
                    output += f"Return value: {result}\n"
                else:
                    output += "Type: Sub (no return value)\n"

                output += f"\nExecution completed in {session.app_type}."

                return output

            except pythoncom.com_error as e:
                last_error = str(e)
                continue  # Try next format

    # All formats failed - list available macros
    available_macros = None
    try:
        available_macros = _list_available_macros(session)
    except Exception as list_error:
        logger.warning(f"Failed to list available macros: {list_error}")

    # Build error message
    if available_macros:
        raise ValueError(
            f"Macro '{macro_name}' not found in {path.name}\n\n"
            f"Available macros:\n{available_macros}\n\n"
            f"Formats tried: {', '.join(formats_to_try)}\n"
            f"Last error: {last_error}"
        )
    else:
        raise ValueError(
            f"Macro '{macro_name}' not found in {path.name}\n\n"
            f"Formats tried: {', '.join(formats_to_try)}\n"
            f"Last error: {last_error}"
        )


def _list_available_macros(session) -> str:
    """List all public macros in the VBA project."""
    macros = []

    try:
        vb_project = session.vb_project

        for component in vb_project.VBComponents:
            module_name = component.Name
            code_module = component.CodeModule

            # Iterate through all lines to find Public Sub/Function
            line_count = code_module.CountOfLines
            for line_num in range(1, line_count + 1):
                line = code_module.Lines(line_num, 1).strip()

                # Check for Public Sub or Public Function
                if line.startswith("Public Sub ") or line.startswith("Sub "):
                    # Extract procedure name
                    if "(" in line:
                        proc_name = line.split("Sub ")[1].split("(")[0].strip()
                        macros.append(f"  - {module_name}.{proc_name} (Sub)")

                elif line.startswith("Public Function ") or line.startswith("Function "):
                    if "(" in line:
                        proc_name = line.split("Function ")[1].split("(")[0].strip()
                        macros.append(f"  - {module_name}.{proc_name} (Function)")

        if macros:
            return "\n".join(macros)
        else:
            return "  (No public macros found)"

    except Exception as e:
        return f"  (Error listing macros: {e})"


async def get_worksheet_data_tool(
    file_path: str,
    sheet_name: str,
    range: Optional[str] = None,
    table_name: Optional[str] = None,
    columns: Optional[List[str]] = None,
    include_headers: bool = True,
    include_formulas: bool = False
) -> str:
    """
    Read data from Excel worksheet, table, or Access table.

    Args:
        file_path: Absolute path to Office file
        sheet_name: Worksheet name (Excel) or table name (Access)
        range: Cell range (e.g., 'A1:D10') or None for entire used range
        table_name: Excel Table name (e.g., 'BudgetTable') - NEW
        columns: List of column names to extract (e.g., ['Name', 'Age']) - NEW
        include_headers: Include header row in output (default: true) - NEW
        include_formulas: Return formulas instead of values (Excel only, default: false)

    Returns:
        JSON formatted data

    Examples:
        # Read by range (existing)
        get_worksheet_data(file, "Sheet1", range="A1:C10")

        # Read entire table (NEW)
        get_worksheet_data(file, "Sheet1", table_name="BudgetTable")

        # Read specific columns from table (NEW)
        get_worksheet_data(file, "Sheet1", table_name="BudgetTable",
                          columns=["Name", "Total"])

    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If sheet/table not found or range too large
        RuntimeError: If read operation fails
    """
    path = Path(file_path)
    manager = OfficeSessionManager.get_instance()

    # Get or create session
    session = await manager.get_or_create_session(path, read_only=True)
    session.refresh_last_accessed()

    # Route to appropriate handler
    if session.app_type == "Excel":
        return await _get_excel_data(
            session, sheet_name, range, table_name, columns,
            include_headers, include_formulas
        )
    elif session.app_type == "Access":
        return await _get_access_data(session, sheet_name)
    else:
        raise ValueError(
            f"{session.app_type} does not support get_worksheet_data. "
            f"Only Excel and Access are supported."
        )


async def _get_excel_data(
    session,
    sheet_name: str,
    range_str: Optional[str],
    table_name: Optional[str],
    columns: Optional[List[str]],
    include_headers: bool,
    include_formulas: bool
) -> str:
    """Extract data from Excel worksheet or table."""
    try:
        ws = session.file_obj.Worksheets(sheet_name)
    except Exception:
        raise ValueError(
            f"Worksheet '{sheet_name}' not found in {session.file_path.name}\n"
            f"Available sheets: {[ws.Name for ws in session.file_obj.Worksheets]}"
        )

    # Mode 1: Extract from Excel Table (ListObject)
    if table_name:
        try:
            table = ws.ListObjects(table_name)
        except Exception:
            available_tables = [t.Name for t in ws.ListObjects]
            raise ValueError(
                f"Table '{table_name}' not found in sheet '{sheet_name}'\n"
                f"Available tables: {available_tables if available_tables else '(none)'}"
            )

        # Get data range (body only, no headers yet)
        data_range = table.DataBodyRange
        if data_range is None:
            # Empty table
            data_array = []
        else:
            # Check if we need to filter columns
            if columns:
                header_range = table.HeaderRowRange
                col_indices = _find_column_indices(header_range, columns)
                data_array = _extract_columns(data_range, col_indices)
            else:
                # Get all data
                data = data_range.Value
                rows = data_range.Rows.Count
                cols = data_range.Columns.Count
                data_array = _normalize_com_data(data, rows, cols)

        # Add headers if requested
        if include_headers:
            header_data = table.HeaderRowRange.Value
            if not isinstance(header_data, tuple):
                headers = [str(header_data)]
            else:
                headers = [str(h) for h in header_data[0]]

            # Filter headers if we filtered columns
            if columns:
                col_indices = _find_column_indices(table.HeaderRowRange, columns)
                headers = [headers[i] for i in col_indices]

            data_array = [headers] + data_array

        # Format as JSON
        result_json = json.dumps(data_array, indent=2, default=str)

        return "\n".join([
            "**Data Retrieved from Excel Table**",
            "",
            f"File: {session.file_path.name}",
            f"Sheet: {sheet_name}",
            f"Table: {table_name}",
            f"Columns: {', '.join(columns) if columns else 'All'}",
            f"Dimensions: {len(data_array)} rows × {len(data_array[0]) if data_array else 0} columns",
            f"Headers: {'Included' if include_headers else 'Excluded'}",
            "",
            "```json",
            result_json,
            "```"
        ])

    # Mode 2: Extract from range (existing behavior)
    else:
        # Determine range
        if range_str:
            range_obj = ws.Range(range_str)
        else:
            range_obj = ws.UsedRange

        # Check size
        cell_count = range_obj.Rows.Count * range_obj.Columns.Count
        if cell_count > MAX_CELLS:
            raise ValueError(
                f"Range too large: {cell_count:,} cells (limit: {MAX_CELLS:,})\n"
                f"Specify a smaller range."
            )

        # Get data
        if include_formulas:
            # Get formulas (returns None for non-formula cells)
            data = range_obj.Formula
        else:
            # Get values
            data = range_obj.Value

        # Convert COM data to 2D array
        data_array = _normalize_com_data(data, range_obj.Rows.Count, range_obj.Columns.Count)

        # Format as JSON
        result_json = json.dumps(data_array, indent=2, default=str)

        return "\n".join([
            "**Data Retrieved from Excel**",
            "",
            f"File: {session.file_path.name}",
            f"Sheet: {sheet_name}",
            f"Range: {range_str or 'Entire used range'}",
            f"Dimensions: {len(data_array)} rows × {len(data_array[0]) if data_array else 0} columns",
            f"Type: {'Formulas' if include_formulas else 'Values'}",
            "",
            "```json",
            result_json,
            "```"
        ])


async def _get_access_data(session, table_name: str) -> str:
    """Extract data from Access table."""
    try:
        db = session.app.CurrentDb()
        rs = db.OpenRecordset(table_name)

        # Get field names
        fields = [field.Name for field in rs.Fields]

        # Read all records
        data = []
        if not rs.EOF:
            rs.MoveFirst()
            while not rs.EOF:
                row = [rs.Fields(field).Value for field in fields]
                data.append(row)
                rs.MoveNext()

        rs.Close()

        # Format with headers
        result = {
            "headers": fields,
            "rows": data
        }

        return "\n".join([
            "**Data Retrieved from Access**",
            "",
            f"Database: {session.file_path.name}",
            f"Table: {table_name}",
            f"Records: {len(data)}",
            f"Fields: {len(fields)}",
            "",
            "```json",
            json.dumps(result, indent=2, default=str),
            "```"
        ])

    except Exception as e:
        raise ValueError(
            f"Failed to read table '{table_name}': {str(e)}\n"
            f"Make sure the table name is correct."
        )


def _normalize_com_data(data: Any, rows: int, cols: int) -> List[List[Any]]:
    """
    Normalize COM Range.Value data to 2D array.

    Excel COM returns different structures depending on range size:
    - Single cell: direct value
    - Single row/column: tuple
    - 2D range: tuple of tuples
    """
    if data is None:
        # Empty range
        return [[None] * cols for _ in range(rows)]

    if not isinstance(data, tuple):
        # Single cell
        return [[data]]

    if rows == 1 and cols == 1:
        # Single cell (shouldn't happen but handle it)
        return [[data]]

    if rows == 1:
        # Single row
        return [list(data)]

    if cols == 1:
        # Single column
        return [[item] for item in data]

    # 2D range
    return [list(row) for row in data]


# Helper functions for Excel Table support

def _column_letter_to_number(letter: str) -> int:
    """
    Convert column letter to number (A=1, B=2, ... Z=26, AA=27).

    Args:
        letter: Column letter (e.g., 'A', 'Z', 'AA')

    Returns:
        Column number (1-indexed)
    """
    num = 0
    for char in letter.upper():
        num = num * 26 + (ord(char) - ord('A') + 1)
    return num


def _find_column_indices(header_range, column_names: List[str]) -> List[int]:
    """
    Find column indices by name.

    Args:
        header_range: Excel range containing headers
        column_names: List of column names to find

    Returns:
        List of column indices (0-indexed)

    Raises:
        ValueError: If any column name not found
    """
    # Get header values
    header_data = header_range.Value
    if not isinstance(header_data, tuple):
        headers = [str(header_data).strip()]
    else:
        headers = [str(h).strip() for h in header_data[0]]

    indices = []
    for col_name in column_names:
        try:
            idx = headers.index(col_name)
            indices.append(idx)
        except ValueError:
            raise ValueError(f"Column '{col_name}' not found. Available columns: {headers}")

    return indices


def _extract_columns(data_range, col_indices: List[int]) -> List[List]:
    """
    Extract specific columns from range.

    Args:
        data_range: Excel range containing data
        col_indices: List of column indices to extract (0-indexed)

    Returns:
        2D array with only the specified columns
    """
    data = data_range.Value

    # Handle single cell
    if not isinstance(data, tuple):
        data = [[data]]
    # Handle single row
    elif not isinstance(data[0], tuple):
        data = [data]
    else:
        data = [list(row) for row in data]

    # Extract columns
    result = []
    for row in data:
        result.append([row[i] for i in col_indices])

    return result


async def set_worksheet_data_tool(
    file_path: str,
    sheet_name: str,
    data: List[List[Any]],
    start_cell: str = "A1",
    table_name: Optional[str] = None,
    column_mapping: Optional[dict] = None,
    append: bool = False,
    clear_existing: bool = False
) -> str:
    """
    Write data to Excel worksheet or table.

    Args:
        file_path: Absolute path to Excel file
        sheet_name: Worksheet name (will be created if doesn't exist)
        data: 2D array of values [[row1], [row2], ...]
        start_cell: Top-left cell to start writing (default: "A1")
        table_name: Excel Table name to write to - NEW
        column_mapping: Map data columns to table columns (e.g., {"Name": 0, "Age": 1}) - NEW
        append: Append rows to end of table (requires table_name) - NEW
        clear_existing: Clear all existing data in sheet first (default: false)

    Returns:
        Success message with range written

    Examples:
        # Write to range (existing)
        set_worksheet_data(file, "Sheet1", data, start_cell="A1")

        # Append to table (NEW)
        set_worksheet_data(file, "Sheet1", data, table_name="BudgetTable", append=True)

        # Update specific columns (NEW)
        set_worksheet_data(file, "Sheet1", data, table_name="BudgetTable",
                          column_mapping={"Name": 0, "Total": 1})

    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If data format invalid or range too large
        RuntimeError: If write operation fails
    """
    path = Path(file_path)
    manager = OfficeSessionManager.get_instance()

    # Get or create session (not read-only!)
    session = await manager.get_or_create_session(path, read_only=False)
    session.refresh_last_accessed()

    # Only Excel supports set_worksheet_data
    if session.app_type != "Excel":
        raise ValueError(
            f"{session.app_type} does not support set_worksheet_data. "
            f"Only Excel is supported."
        )

    # Validate data
    _validate_data(data)

    # Get or create worksheet
    try:
        ws = session.file_obj.Worksheets(sheet_name)
    except Exception:
        # Worksheet doesn't exist - create it
        ws = session.file_obj.Worksheets.Add()
        ws.Name = sheet_name
        logger.info(f"Created new worksheet: {sheet_name}")

    # Mode 1: Write to Excel Table
    if table_name:
        try:
            table = ws.ListObjects(table_name)
        except Exception:
            available_tables = [t.Name for t in ws.ListObjects]
            raise ValueError(
                f"Table '{table_name}' not found in sheet '{sheet_name}'\n"
                f"Available tables: {available_tables if available_tables else '(none)'}"
            )

        # Disable calculation for performance
        calc_mode = session.app.Calculation
        session.app.Calculation = -4135  # xlCalculationManual

        try:
            if append:
                # Append rows to end of table
                rows_added = 0
                for row_data in data:
                    new_row = table.ListRows.Add()
                    new_row.Range.Value = tuple(row_data)
                    rows_added += 1

                return "\n".join([
                    "**Data Appended to Excel Table**",
                    "",
                    f"File: {session.file_path.name}",
                    f"Sheet: {sheet_name}",
                    f"Table: {table_name}",
                    f"Rows added: {rows_added}",
                    f"New table size: {table.ListRows.Count} rows",
                    "",
                    "Data appended successfully.",
                    "File remains open. Changes will be saved when you close the session."
                ])

            elif column_mapping:
                # Update specific columns by mapping
                headers = [str(h) for h in table.HeaderRowRange.Value[0]]

                # Validate column names
                for col_name in column_mapping.keys():
                    if col_name not in headers:
                        raise ValueError(
                            f"Column '{col_name}' not found in table.\n"
                            f"Available columns: {headers}"
                        )

                rows_updated = 0
                for row_idx, row_data in enumerate(data, start=1):
                    if row_idx > table.ListRows.Count:
                        # Add new row if needed
                        table.ListRows.Add()

                    for col_name, data_idx in column_mapping.items():
                        col_idx = headers.index(col_name) + 1  # 1-indexed
                        cell = table.DataBodyRange.Cells(row_idx, col_idx)
                        cell.Value = row_data[data_idx]

                    rows_updated += 1

                return "\n".join([
                    "**Data Written to Excel Table (Column Mapping)**",
                    "",
                    f"File: {session.file_path.name}",
                    f"Sheet: {sheet_name}",
                    f"Table: {table_name}",
                    f"Columns updated: {', '.join(column_mapping.keys())}",
                    f"Rows updated: {rows_updated}",
                    "",
                    "Data written successfully.",
                    "File remains open. Changes will be saved when you close the session."
                ])

            else:
                # Replace entire table data
                if table.DataBodyRange is not None:
                    table.DataBodyRange.ClearContents()

                # Add rows
                for row_data in data:
                    new_row = table.ListRows.Add()
                    new_row.Range.Value = tuple(row_data)

                return "\n".join([
                    "**Data Written to Excel Table**",
                    "",
                    f"File: {session.file_path.name}",
                    f"Sheet: {sheet_name}",
                    f"Table: {table_name}",
                    f"Rows written: {len(data)}",
                    f"Columns: {len(data[0])}",
                    "",
                    "Data written successfully.",
                    "File remains open. Changes will be saved when you close the session."
                ])

        finally:
            # Restore calculation and recalculate
            session.app.Calculation = calc_mode
            session.app.Calculate()

    # Mode 2: Write to range (existing behavior)
    else:
        # Clear existing data if requested
        if clear_existing:
            ws.Cells.Clear()

        # Calculate dimensions
        rows = len(data)
        cols = len(data[0])

        # Check size
        cell_count = rows * cols
        if cell_count > MAX_CELLS:
            raise ValueError(
                f"Data too large: {cell_count:,} cells (limit: {MAX_CELLS:,})"
            )

        # Parse start cell
        start_range = ws.Range(start_cell)
        start_row = start_range.Row
        start_col = start_range.Column

        # Calculate end cell
        end_row = start_row + rows - 1
        end_col = start_col + cols - 1

        # Get target range
        target_range = ws.Range(
            ws.Cells(start_row, start_col),
            ws.Cells(end_row, end_col)
        )

        # Disable calculation for performance
        calc_mode = session.app.Calculation
        session.app.Calculation = -4135  # xlCalculationManual

        try:
            # Convert data to tuple of tuples (COM requires this)
            data_tuples = tuple(tuple(row) for row in data)

            # Write data in one operation (FAST!)
            target_range.Value = data_tuples

        finally:
            # Restore calculation and recalculate
            session.app.Calculation = calc_mode
            session.app.Calculate()

        # Format end cell address
        end_cell = _get_cell_address(end_row, end_col)

        return "\n".join([
            "**Data Written to Excel**",
            "",
            f"File: {session.file_path.name}",
            f"Sheet: {sheet_name}",
            f"Range: {start_cell}:{end_cell}",
            f"Dimensions: {rows} rows × {cols} columns",
            f"Total cells: {cell_count:,}",
            "",
            "Data written successfully.",
            "File remains open. Changes will be saved when you close the session."
        ])


def _validate_data(data: List[List[Any]]) -> None:
    """Validate data structure for set_worksheet_data."""
    if not isinstance(data, list):
        raise ValueError("data must be a 2D array (list of lists)")

    if not data:
        raise ValueError("data cannot be empty")

    if not all(isinstance(row, list) for row in data):
        raise ValueError("data must be a 2D array (each row must be a list)")

    # Check rectangular shape
    row_lengths = [len(row) for row in data]
    if len(set(row_lengths)) > 1:
        raise ValueError(
            f"All rows must have the same length (rectangular array).\n"
            f"Row lengths: {row_lengths}"
        )

    if row_lengths[0] == 0:
        raise ValueError("Rows cannot be empty")


def _get_cell_address(row: int, col: int) -> str:
    """
    Convert row/column numbers to Excel address (e.g., 1, 1 → 'A1').

    Args:
        row: Row number (1-indexed)
        col: Column number (1-indexed)

    Returns:
        Excel cell address (e.g., 'A1', 'Z10', 'AA100')
    """
    col_letter = ""
    while col > 0:
        col -= 1
        col_letter = chr(65 + (col % 26)) + col_letter
        col //= 26
    return f"{col_letter}{row}"


async def list_macros_tool(file_path: str) -> str:
    """
    List all public macros (Subs and Functions) in an Office file.

    Args:
        file_path: Path to Office file

    Returns:
        Formatted list of macros with signatures
    """

    manager = OfficeSessionManager.get_instance()
    path = Path(file_path).resolve()

    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    # Get or create session
    session = await manager.get_or_create_session(path, read_only=True)
    session.refresh_last_accessed()

    try:
        macros_by_module = {}
        vb_project = session.vb_project

        for component in vb_project.VBComponents:
            module_name = component.Name
            code_module = component.CodeModule
            module_macros = []

            line_count = code_module.CountOfLines

            for line_num in range(1, line_count + 1):
                line = code_module.Lines(line_num, 1).strip()

                # Public Sub
                if line.startswith("Public Sub ") or line.startswith("Sub "):
                    signature = line
                    # Extract name and parameters
                    if "(" in signature:
                        full_sig = signature.replace("Public ", "").replace("Sub ", "")
                        module_macros.append({
                            "type": "Sub",
                            "signature": full_sig.split("'")[0].strip()  # Remove comments
                        })

                # Public Function
                elif line.startswith("Public Function ") or line.startswith("Function "):
                    signature = line
                    if "(" in signature:
                        full_sig = signature.replace("Public ", "").replace("Function ", "")
                        # Extract return type
                        if " As " in full_sig:
                            name_part, return_type = full_sig.split(" As ", 1)
                            return_type = return_type.split("'")[0].strip()
                        else:
                            name_part = full_sig
                            return_type = "Variant"

                        module_macros.append({
                            "type": "Function",
                            "signature": name_part.split("'")[0].strip(),
                            "returns": return_type
                        })

            if module_macros:
                macros_by_module[module_name] = module_macros

        # Format output
        output = f"### Macros in {path.name}\n\n"

        if not macros_by_module:
            output += "No public macros found.\n"
            return output

        total_macros = sum(len(macros) for macros in macros_by_module.values())
        output += f"**Total:** {total_macros} public macros in {len(macros_by_module)} modules\n\n"

        for module_name, macros in sorted(macros_by_module.items()):
            output += f"#### {module_name}\n\n"

            for macro in macros:
                if macro["type"] == "Sub":
                    output += f"- `{macro['signature']}` (Sub)\n"
                else:
                    output += f"- `{macro['signature']}` → {macro['returns']} (Function)\n"

            output += "\n"

        output += f"**Usage:** `run_macro(file, \"MacroName\")` or `run_macro(file, \"Module.MacroName\")`\n"

        return output

    except Exception as e:
        raise RuntimeError(f"Error listing macros: {str(e)}")
