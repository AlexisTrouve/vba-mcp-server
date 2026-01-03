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
    elif session.app_type == "Access":
        # Access Application.Run via COM wants JUST the procedure name, NOT Module.Procedure
        # The module prefix causes "cannot find procedure" error
        formats_to_try = [proc_name]  # Always try just the procedure name first
    elif session.app_type == "Word":
        # Word typically just uses macro name
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
    include_formulas: bool = False,
    # Access-specific parameters
    sql_query: Optional[str] = None,
    where_clause: Optional[str] = None,
    order_by: Optional[str] = None,
    limit: Optional[int] = None
) -> str:
    """
    Read data from Excel worksheet, table, or Access table.

    Args:
        file_path: Absolute path to Office file
        sheet_name: Worksheet name (Excel) or table name (Access)
        range: Cell range (e.g., 'A1:D10') or None for entire used range
        table_name: Excel Table name (e.g., 'BudgetTable')
        columns: List of column names to extract (e.g., ['Name', 'Age'])
        include_headers: Include header row in output (default: true)
        include_formulas: Return formulas instead of values (Excel only, default: false)
        sql_query: Custom SQL query (Access only, overrides table_name)
        where_clause: SQL WHERE clause without 'WHERE' keyword (Access only)
        order_by: SQL ORDER BY clause without 'ORDER BY' keyword (Access only)
        limit: Maximum number of records to return (Access only)

    Returns:
        JSON formatted data

    Examples:
        # Excel - Read by range
        get_worksheet_data(file, "Sheet1", range="A1:C10")

        # Excel - Read entire table
        get_worksheet_data(file, "Sheet1", table_name="BudgetTable")

        # Excel - Read specific columns from table
        get_worksheet_data(file, "Sheet1", table_name="BudgetTable",
                          columns=["Name", "Total"])

        # Access - Read table
        get_worksheet_data(file, "Clients")

        # Access - Read with filter
        get_worksheet_data(file, "Clients", where_clause="DateCreation > #2024-01-01#")

        # Access - Custom SQL
        get_worksheet_data(file, "", sql_query="SELECT * FROM Clients WHERE Actif = True")

        # Access - With limit and order
        get_worksheet_data(file, "Commandes", order_by="Date DESC", limit=100)

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
        return await _get_access_data(
            session, sheet_name, sql_query, where_clause,
            order_by, limit, columns
        )
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


async def _get_access_data(
    session,
    table_name: str,
    sql_query: Optional[str] = None,
    where_clause: Optional[str] = None,
    order_by: Optional[str] = None,
    limit: Optional[int] = None,
    columns: Optional[List[str]] = None
) -> str:
    """
    Extract data from Access table or query.

    Args:
        session: Active Access session
        table_name: Table name to query
        sql_query: Custom SQL query (overrides table_name if provided)
        where_clause: SQL WHERE clause (without 'WHERE' keyword)
        order_by: SQL ORDER BY clause (without 'ORDER BY' keyword)
        limit: Maximum number of records to return
        columns: List of specific columns to retrieve

    Returns:
        Formatted JSON data with headers and rows
    """
    try:
        db = session.app.CurrentDb()

        # Build SQL query
        if sql_query:
            # Use custom SQL directly
            sql = sql_query
            query_description = f"Custom SQL: {sql[:100]}{'...' if len(sql) > 100 else ''}"
        else:
            # Build SELECT statement
            if columns:
                cols = ", ".join([f"[{col}]" for col in columns])
            else:
                cols = "*"

            sql = f"SELECT {cols} FROM [{table_name}]"

            if where_clause:
                sql += f" WHERE {where_clause}"

            if order_by:
                sql += f" ORDER BY {order_by}"

            query_description = f"Table: {table_name}"
            if where_clause:
                query_description += f" | WHERE: {where_clause}"
            if order_by:
                query_description += f" | ORDER BY: {order_by}"

        # Execute query
        try:
            rs = db.OpenRecordset(sql)
        except Exception as sql_error:
            # Try to provide helpful error message
            raise ValueError(
                f"SQL Error: {str(sql_error)}\n"
                f"Query: {sql}\n"
                f"Check table name, column names, and SQL syntax."
            )

        # Get field names from result
        fields = [field.Name for field in rs.Fields]

        # Read records with optional limit
        data = []
        count = 0

        if not rs.EOF:
            rs.MoveFirst()
            while not rs.EOF:
                # Check limit
                if limit is not None and count >= limit:
                    break

                # Read row
                row = []
                for i in range(rs.Fields.Count):
                    value = rs.Fields(i).Value
                    row.append(value)
                data.append(row)

                rs.MoveNext()
                count += 1

        rs.Close()

        # Check if more records exist (for limit info)
        limited = limit is not None and count >= limit

        # Format with headers
        result = {
            "headers": fields,
            "rows": data
        }

        # Build output message
        output_lines = [
            "**Data Retrieved from Access**",
            "",
            f"Database: {session.file_path.name}",
            f"Query: {query_description}",
            f"Records: {len(data)}" + (" (limited)" if limited else ""),
            f"Fields: {len(fields)}",
        ]

        if columns:
            output_lines.append(f"Columns selected: {', '.join(columns)}")

        if limit:
            output_lines.append(f"Limit: {limit}")

        output_lines.extend([
            "",
            "```json",
            json.dumps(result, indent=2, default=str),
            "```"
        ])

        return "\n".join(output_lines)

    except ValueError:
        # Re-raise ValueError (SQL errors with helpful messages)
        raise
    except Exception as e:
        # Try to list available tables for helpful error
        try:
            db = session.app.CurrentDb()
            tables = []
            for td in db.TableDefs:
                if not td.Name.startswith("MSys"):  # Skip system tables
                    tables.append(td.Name)
            tables_list = ", ".join(tables) if tables else "(no tables found)"
        except Exception:
            tables_list = "(unable to list tables)"

        raise ValueError(
            f"Failed to read from Access: {str(e)}\n"
            f"Table/Query: {table_name or sql_query}\n"
            f"Available tables: {tables_list}"
        )


async def _set_access_data(
    session,
    table_name: str,
    data: List[List[Any]],
    columns: Optional[List[str]] = None,
    mode: str = "append"
) -> str:
    """
    Write data to Access table.

    Args:
        session: Active Access session
        table_name: Target table name
        data: 2D array of values [[row1], [row2], ...]
        columns: Column names (if None, uses table field order)
        mode: "append" (add records) or "replace" (delete all then insert)

    Returns:
        Success message with record count
    """
    try:
        db = session.app.CurrentDb()

        # Validate mode
        if mode not in ("append", "replace"):
            raise ValueError(f"Invalid mode '{mode}'. Use 'append' or 'replace'.")

        # Validate data
        if not data:
            raise ValueError("Data cannot be empty")

        if not all(isinstance(row, list) for row in data):
            raise ValueError("Data must be a 2D array (list of lists)")

        # Replace mode: delete all existing records first
        records_deleted = 0
        if mode == "replace":
            try:
                # Count before delete
                count_rs = db.OpenRecordset(f"SELECT COUNT(*) FROM [{table_name}]")
                records_deleted = count_rs.Fields(0).Value or 0
                count_rs.Close()

                # Delete all records
                db.Execute(f"DELETE FROM [{table_name}]")
            except Exception as del_error:
                raise ValueError(
                    f"Failed to delete existing records: {str(del_error)}\n"
                    f"Table: {table_name}"
                )

        # Open recordset for adding records
        try:
            rs = db.OpenRecordset(table_name)
        except Exception as open_error:
            # Try to list available tables
            tables = []
            try:
                for td in db.TableDefs:
                    if not td.Name.startswith("MSys"):
                        tables.append(td.Name)
            except Exception:
                pass

            raise ValueError(
                f"Table '{table_name}' not found or cannot be opened: {str(open_error)}\n"
                f"Available tables: {', '.join(tables) if tables else '(none)'}"
            )

        # Get table field names if columns not specified
        if columns is None:
            columns = []
            for field in rs.Fields:
                # Skip AutoNumber fields (they auto-increment)
                if field.Attributes & 16:  # dbAutoIncrField
                    continue
                columns.append(field.Name)

        # Validate column count matches data
        expected_cols = len(columns)
        for i, row in enumerate(data):
            if len(row) != expected_cols:
                raise ValueError(
                    f"Row {i+1} has {len(row)} values, expected {expected_cols} "
                    f"(columns: {', '.join(columns)})"
                )

        # Insert records
        inserted = 0
        errors = []

        for row_idx, row in enumerate(data):
            try:
                rs.AddNew()

                for col_idx, col_name in enumerate(columns):
                    try:
                        rs.Fields(col_name).Value = row[col_idx]
                    except Exception as field_error:
                        errors.append(f"Row {row_idx+1}, Column '{col_name}': {str(field_error)}")

                rs.Update()
                inserted += 1

            except Exception as row_error:
                errors.append(f"Row {row_idx+1}: {str(row_error)}")

        rs.Close()

        # Build result message
        output_lines = [
            "**Data Written to Access**",
            "",
            f"Database: {session.file_path.name}",
            f"Table: {table_name}",
            f"Mode: {mode}",
        ]

        if mode == "replace":
            output_lines.append(f"Records deleted: {records_deleted}")

        output_lines.extend([
            f"Records inserted: {inserted}",
            f"Columns: {', '.join(columns)}",
        ])

        if errors:
            output_lines.extend([
                "",
                f"Warnings ({len(errors)}):",
            ])
            for error in errors[:5]:  # Show first 5 errors
                output_lines.append(f"  - {error}")
            if len(errors) > 5:
                output_lines.append(f"  ... and {len(errors) - 5} more")

        output_lines.extend([
            "",
            "Data written successfully." if not errors else "Data written with some errors.",
            "Changes are auto-saved in Access."
        ])

        return "\n".join(output_lines)

    except ValueError:
        raise
    except Exception as e:
        raise ValueError(
            f"Failed to write to Access table '{table_name}': {str(e)}"
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
    clear_existing: bool = False,
    # Access-specific parameters
    columns: Optional[List[str]] = None,
    mode: str = "append"
) -> str:
    """
    Write data to Excel worksheet, table, or Access table.

    Args:
        file_path: Absolute path to Office file
        sheet_name: Worksheet name (Excel) or table name (Access)
        data: 2D array of values [[row1], [row2], ...]
        start_cell: Top-left cell to start writing (Excel only, default: "A1")
        table_name: Excel Table name to write to
        column_mapping: Map data columns to table columns (e.g., {"Name": 0, "Age": 1})
        append: Append rows to end of table (Excel tables only)
        clear_existing: Clear all existing data in sheet first (Excel only, default: false)
        columns: Column names for data (Access only, if None uses table order)
        mode: Write mode for Access - "append" (add records), "replace" (delete all then insert)

    Returns:
        Success message with details

    Examples:
        # Excel - Write to range
        set_worksheet_data(file, "Sheet1", data, start_cell="A1")

        # Excel - Append to table
        set_worksheet_data(file, "Sheet1", data, table_name="BudgetTable", append=True)

        # Access - Append records
        set_worksheet_data(file, "Clients", [["Jean", "jean@email.com"]],
                          columns=["Nom", "Email"])

        # Access - Replace all data
        set_worksheet_data(file, "TempData", data, mode="replace")

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

    # Route to appropriate handler
    if session.app_type == "Access":
        return await _set_access_data(session, sheet_name, data, columns, mode)
    elif session.app_type != "Excel":
        raise ValueError(
            f"{session.app_type} does not support set_worksheet_data. "
            f"Only Excel and Access are supported."
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


# =============================================================================
# Access-specific tools
# =============================================================================

def _get_query_type_name(query_type: int) -> str:
    """Convert Access query type constant to readable name."""
    query_types = {
        0: "Select",
        1: "Crosstab",
        2: "Delete",
        3: "Update",
        4: "Append",
        5: "Make-Table",
        6: "DDL",
        7: "SQL Pass-Through",
        8: "Union",
    }
    return query_types.get(query_type, f"Unknown ({query_type})")


async def list_access_queries_tool(file_path: str) -> str:
    """
    List all queries (QueryDefs) in an Access database.

    Args:
        file_path: Path to Access database (.accdb or .mdb)

    Returns:
        Formatted list of queries with name, type, and SQL preview

    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If file is not an Access database
    """
    path = Path(file_path).resolve()

    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    manager = OfficeSessionManager.get_instance()
    session = await manager.get_or_create_session(path, read_only=True)
    session.refresh_last_accessed()

    if session.app_type != "Access":
        raise ValueError(
            f"list_access_queries only works with Access databases. "
            f"Got: {session.app_type}"
        )

    try:
        db = session.app.CurrentDb()
        queries = []

        for qd in db.QueryDefs:
            # Skip system queries (start with ~)
            if qd.Name.startswith("~"):
                continue

            # Truncate long SQL for preview
            sql_preview = qd.SQL.strip()
            if len(sql_preview) > 150:
                sql_preview = sql_preview[:150] + "..."

            queries.append({
                "name": qd.Name,
                "type": _get_query_type_name(qd.Type),
                "sql": sql_preview
            })

        # Format output
        output_lines = [
            f"### Queries in {path.name}",
            "",
            f"**Total:** {len(queries)} queries",
            ""
        ]

        if not queries:
            output_lines.append("No queries found in this database.")
        else:
            for query in sorted(queries, key=lambda q: q["name"]):
                output_lines.extend([
                    f"#### {query['name']}",
                    f"**Type:** {query['type']}",
                    "```sql",
                    query["sql"],
                    "```",
                    ""
                ])

        output_lines.append(
            "**Usage:** `run_access_query(file, query_name=\"QueryName\")` "
            "or `run_access_query(file, sql=\"SELECT ...\")`"
        )

        return "\n".join(output_lines)

    except Exception as e:
        raise RuntimeError(f"Error listing queries: {str(e)}")


def _is_action_query(sql: str) -> bool:
    """Check if SQL is an action query (DELETE, UPDATE, INSERT, DROP, etc.)."""
    sql_upper = sql.strip().upper()
    action_keywords = ('DELETE', 'UPDATE', 'INSERT', 'DROP', 'ALTER', 'CREATE', 'TRUNCATE')
    return sql_upper.startswith(action_keywords)


async def run_access_query_tool(
    file_path: str,
    query_name: Optional[str] = None,
    sql: Optional[str] = None,
    limit: Optional[int] = None
) -> str:
    """
    Execute an Access query and return results.

    Supports both SELECT queries (returns data) and action queries
    (DELETE, UPDATE, INSERT - returns affected row count).

    Args:
        file_path: Path to Access database (.accdb or .mdb)
        query_name: Name of saved query to execute
        sql: Direct SQL to execute (overrides query_name)
        limit: Maximum number of records to return (SELECT only)

    Returns:
        JSON formatted query results (SELECT) or action result message

    Examples:
        # Run saved query
        run_access_query(file, query_name="ClientsActifs")

        # Run direct SQL SELECT
        run_access_query(file, sql="SELECT * FROM Clients WHERE Ville = 'Paris'")

        # Run action query (DELETE/UPDATE/INSERT)
        run_access_query(file, sql="DELETE FROM Clients WHERE Inactive = True")
        run_access_query(file, sql="UPDATE Clients SET Status = 'Active' WHERE ID = 5")

        # With limit
        run_access_query(file, query_name="AllOrders", limit=100)

    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If neither query_name nor sql provided, or query not found
    """
    if not query_name and not sql:
        raise ValueError("Either query_name or sql must be provided")

    path = Path(file_path).resolve()

    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    manager = OfficeSessionManager.get_instance()

    # Action queries need write access
    is_action = sql and _is_action_query(sql)
    session = await manager.get_or_create_session(path, read_only=not is_action)
    session.refresh_last_accessed()

    if session.app_type != "Access":
        raise ValueError(
            f"run_access_query only works with Access databases. "
            f"Got: {session.app_type}"
        )

    try:
        db = session.app.CurrentDb()

        # Determine SQL to execute
        if sql:
            query_sql = sql
            query_description = f"Custom SQL"
        else:
            # Find saved query
            try:
                qd = db.QueryDefs(query_name)
                query_sql = qd.SQL
                query_description = f"Query: {query_name}"
                # Check if saved query is an action query
                is_action = _is_action_query(query_sql)
            except Exception:
                # List available queries for error message
                available = []
                for qd in db.QueryDefs:
                    if not qd.Name.startswith("~"):
                        available.append(qd.Name)

                raise ValueError(
                    f"Query '{query_name}' not found.\n"
                    f"Available queries: {', '.join(available) if available else '(none)'}"
                )

        # Execute action query (DELETE, UPDATE, INSERT, etc.)
        if is_action:
            try:
                db.Execute(query_sql)
                records_affected = db.RecordsAffected
            except Exception as exec_error:
                raise ValueError(
                    f"Action query failed: {str(exec_error)}\n"
                    f"SQL: {query_sql[:200]}{'...' if len(query_sql) > 200 else ''}"
                )

            # Determine action type for message
            sql_upper = query_sql.strip().upper()
            if sql_upper.startswith('DELETE'):
                action_type = "deleted"
            elif sql_upper.startswith('UPDATE'):
                action_type = "updated"
            elif sql_upper.startswith('INSERT'):
                action_type = "inserted"
            else:
                action_type = "affected"

            output_lines = [
                "**Action Query Executed**",
                "",
                f"Database: {path.name}",
                f"Query: {query_description}",
                f"Records {action_type}: {records_affected}",
                "",
                f"SQL: `{query_sql[:100]}{'...' if len(query_sql) > 100 else ''}`"
            ]

            return "\n".join(output_lines)

        # Execute SELECT query
        try:
            rs = db.OpenRecordset(query_sql)
        except Exception as exec_error:
            raise ValueError(
                f"Query execution failed: {str(exec_error)}\n"
                f"SQL: {query_sql[:200]}{'...' if len(query_sql) > 200 else ''}"
            )

        # Get field names
        fields = [field.Name for field in rs.Fields]

        # Read records
        data = []
        count = 0

        if not rs.EOF:
            rs.MoveFirst()
            while not rs.EOF:
                if limit is not None and count >= limit:
                    break

                row = []
                for i in range(rs.Fields.Count):
                    row.append(rs.Fields(i).Value)
                data.append(row)

                rs.MoveNext()
                count += 1

        rs.Close()

        # Check if limited
        limited = limit is not None and count >= limit

        # Format result
        result = {
            "headers": fields,
            "rows": data
        }

        output_lines = [
            "**Query Results**",
            "",
            f"Database: {path.name}",
            f"Query: {query_description}",
            f"Records: {len(data)}" + (" (limited)" if limited else ""),
            f"Fields: {len(fields)}",
        ]

        if limit:
            output_lines.append(f"Limit: {limit}")

        output_lines.extend([
            "",
            "```json",
            json.dumps(result, indent=2, default=str),
            "```"
        ])

        return "\n".join(output_lines)

    except ValueError:
        raise
    except Exception as e:
        raise RuntimeError(f"Error executing query: {str(e)}")


async def list_access_tables_tool(file_path: str) -> str:
    """
    List all tables in an Access database with schema information.

    Args:
        file_path: Path to Access database (.accdb or .mdb)

    Returns:
        Formatted list of tables with fields and types

    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If file is not an Access database
    """
    path = Path(file_path).resolve()

    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    manager = OfficeSessionManager.get_instance()
    session = await manager.get_or_create_session(path, read_only=True)
    session.refresh_last_accessed()

    if session.app_type != "Access":
        raise ValueError(
            f"list_access_tables only works with Access databases. "
            f"Got: {session.app_type}"
        )

    # Field type mapping
    field_types = {
        1: "Boolean",
        2: "Byte",
        3: "Integer",
        4: "Long",
        5: "Currency",
        6: "Single",
        7: "Double",
        8: "Date/Time",
        10: "Text",
        11: "OLE Object",
        12: "Memo",
        15: "GUID",
        16: "Big Integer",
        101: "Attachment",
        102: "Complex",
    }

    try:
        db = session.app.CurrentDb()
        tables = []

        for td in db.TableDefs:
            # Skip system tables
            if td.Name.startswith("MSys") or td.Name.startswith("~"):
                continue

            # Get fields
            fields = []
            for field in td.Fields:
                field_type = field_types.get(field.Type, f"Type {field.Type}")

                field_info = {
                    "name": field.Name,
                    "type": field_type,
                }

                # Add size for text fields
                if field.Type == 10:  # Text
                    field_info["size"] = field.Size

                # Mark primary key / auto-increment
                if field.Attributes & 16:  # dbAutoIncrField
                    field_info["auto_increment"] = True

                fields.append(field_info)

            # Get record count
            try:
                record_count = td.RecordCount
            except Exception:
                record_count = -1

            tables.append({
                "name": td.Name,
                "fields": fields,
                "record_count": record_count
            })

        # Format output
        output_lines = [
            f"### Tables in {path.name}",
            "",
            f"**Total:** {len(tables)} tables",
            ""
        ]

        if not tables:
            output_lines.append("No user tables found in this database.")
        else:
            for table in sorted(tables, key=lambda t: t["name"]):
                records = table["record_count"]
                record_str = f"{records} records" if records >= 0 else "unknown records"

                output_lines.extend([
                    f"#### {table['name']}",
                    f"*{record_str}, {len(table['fields'])} fields*",
                    ""
                ])

                for field in table["fields"]:
                    field_desc = f"- **{field['name']}** ({field['type']}"
                    if "size" in field:
                        field_desc += f", max {field['size']}"
                    if field.get("auto_increment"):
                        field_desc += ", AutoNumber"
                    field_desc += ")"
                    output_lines.append(field_desc)

                output_lines.append("")

        return "\n".join(output_lines)

    except Exception as e:
        raise RuntimeError(f"Error listing tables: {str(e)}")


# =============================================================================
# Access Forms tools
# =============================================================================

async def list_access_forms_tool(file_path: str) -> str:
    """
    [PRO] List all forms in an Access database.

    Args:
        file_path: Path to .accdb or .mdb file

    Returns:
        Formatted list of forms with metadata

    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If file is not an Access database
    """
    path = Path(file_path).resolve()

    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    manager = OfficeSessionManager.get_instance()
    session = await manager.get_or_create_session(path, read_only=True)
    session.refresh_last_accessed()

    if session.app_type != "Access":
        raise ValueError(
            f"list_access_forms only works with Access databases. "
            f"Got: {session.app_type}"
        )

    try:
        app = session.app
        forms = []

        # Iterate through AllForms collection
        for form in app.CurrentProject.AllForms:
            forms.append({
                "name": form.Name,
                "is_loaded": form.IsLoaded,
            })

        # Format output
        output_lines = [
            "## Access Forms",
            "",
            f"**Database:** {path.name}",
            f"**Total forms:** {len(forms)}",
            ""
        ]

        if forms:
            output_lines.append("| Name | Loaded |")
            output_lines.append("|------|--------|")
            for f in sorted(forms, key=lambda x: x["name"]):
                loaded = "Yes" if f["is_loaded"] else "No"
                output_lines.append(f"| {f['name']} | {loaded} |")
        else:
            output_lines.append("_No forms found in this database._")

        output_lines.extend([
            "",
            "**Usage:**",
            "- `export_form_definition` to export form as text",
            "- `create_access_form` to create new forms",
            "- `delete_access_form` to remove forms"
        ])

        return "\n".join(output_lines)

    except Exception as e:
        raise RuntimeError(f"Error listing forms: {str(e)}")


async def create_access_form_tool(
    file_path: str,
    form_name: str,
    record_source: Optional[str] = None,
    form_type: str = "single"
) -> str:
    """
    [PRO] Create a new Access form.

    Args:
        file_path: Path to .accdb or .mdb file
        form_name: Name for the new form
        record_source: Table or query name to bind to (optional)
        form_type: "single" (default), "continuous", or "datasheet"

    Returns:
        Success message with form details

    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If form already exists or invalid parameters

    Examples:
        # Empty form
        create_access_form("db.accdb", "frm_New")

        # Form bound to table
        create_access_form("db.accdb", "frm_Clients", record_source="Clients")

        # Continuous form
        create_access_form("db.accdb", "frm_List", record_source="Items",
                          form_type="continuous")
    """
    path = Path(file_path).resolve()

    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    manager = OfficeSessionManager.get_instance()
    session = await manager.get_or_create_session(path, read_only=False)
    session.refresh_last_accessed()

    if session.app_type != "Access":
        raise ValueError(
            f"create_access_form only works with Access databases. "
            f"Got: {session.app_type}"
        )

    try:
        app = session.app

        # Check if form already exists
        for form in app.CurrentProject.AllForms:
            if form.Name.lower() == form_name.lower():
                raise ValueError(f"Form '{form_name}' already exists")

        # Validate form_type
        view_map = {"single": 0, "continuous": 1, "datasheet": 2}
        if form_type not in view_map:
            raise ValueError(
                f"Invalid form_type '{form_type}'. "
                f"Must be one of: single, continuous, datasheet"
            )

        # Create the form using CreateForm
        # This creates a form in Design view with a temp name like "Form1"
        frm = app.CreateForm()
        temp_name = frm.Name

        # Set record source if provided
        if record_source:
            frm.RecordSource = record_source

        # Set default view
        frm.DefaultView = view_map[form_type]

        # Save the form first (required before closing)
        app.DoCmd.Save(2, temp_name)  # acForm = 2

        # Close the form
        app.DoCmd.Close(2, temp_name, 1)  # acForm=2, acSaveYes=1

        # Rename to the desired name if different
        if temp_name.lower() != form_name.lower():
            app.DoCmd.Rename(form_name, 2, temp_name)  # acForm = 2

        return f"""## Form Created

**Name:** {form_name}
**Database:** {path.name}
**Record Source:** {record_source or "None (unbound)"}
**Type:** {form_type}

Form created successfully.

**Next steps:**
- Use `export_form_definition` to view/edit the form structure as text
- Use `import_form_definition` to import modified form definitions
- Open the form in Access to add controls visually
"""

    except ValueError:
        raise
    except Exception as e:
        raise RuntimeError(f"Error creating form: {str(e)}")


async def delete_access_form_tool(
    file_path: str,
    form_name: str,
    backup_first: bool = True
) -> str:
    """
    [PRO] Delete an Access form.

    Args:
        file_path: Path to .accdb or .mdb file
        form_name: Name of form to delete
        backup_first: If True, export to temp folder before deleting (default: True)

    Returns:
        Confirmation message with backup path if created

    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If form not found
    """
    from datetime import datetime

    path = Path(file_path).resolve()

    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    manager = OfficeSessionManager.get_instance()
    session = await manager.get_or_create_session(path, read_only=False)
    session.refresh_last_accessed()

    if session.app_type != "Access":
        raise ValueError(
            f"delete_access_form only works with Access databases. "
            f"Got: {session.app_type}"
        )

    try:
        app = session.app

        # Verify form exists and get exact name
        actual_name = None
        for form in app.CurrentProject.AllForms:
            if form.Name.lower() == form_name.lower():
                actual_name = form.Name
                break

        if not actual_name:
            # List available forms for helpful error
            available = [f.Name for f in app.CurrentProject.AllForms]
            raise ValueError(
                f"Form '{form_name}' not found.\n"
                f"Available forms: {', '.join(available) if available else '(none)'}"
            )

        # Backup if requested
        backup_path = None
        if backup_first:
            backup_dir = path.parent / ".form_backups"
            backup_dir.mkdir(exist_ok=True)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_path = backup_dir / f"{actual_name}_{timestamp}.txt"
            # acForm = 2
            app.SaveAsText(2, actual_name, str(backup_path))

        # Delete the form
        # acForm = 2
        app.DoCmd.DeleteObject(2, actual_name)

        # Build result message
        output_lines = [
            "## Form Deleted",
            "",
            f"**Name:** {actual_name}",
            f"**Database:** {path.name}",
        ]

        if backup_path:
            output_lines.extend([
                "",
                f"**Backup saved:** {backup_path}",
                "",
                "To restore, use `import_form_definition` with the backup file."
            ])

        return "\n".join(output_lines)

    except ValueError:
        raise
    except Exception as e:
        raise RuntimeError(f"Error deleting form: {str(e)}")


async def export_form_definition_tool(
    file_path: str,
    form_name: str,
    output_path: Optional[str] = None
) -> str:
    """
    [PRO] Export Access form definition to text file (SaveAsText).

    This exports the complete form definition including:
    - All controls and their properties
    - Layout and positioning
    - VBA code behind the form
    - Event bindings

    The exported file can be:
    - Viewed to understand form structure
    - Modified with any text editor
    - Re-imported with import_form_definition
    - Version controlled with Git

    Args:
        file_path: Path to .accdb or .mdb file
        form_name: Name of form to export
        output_path: Where to save .txt file (default: same folder as db)

    Returns:
        Path to exported file with content preview

    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If form not found
    """
    path = Path(file_path).resolve()

    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    manager = OfficeSessionManager.get_instance()
    session = await manager.get_or_create_session(path, read_only=True)
    session.refresh_last_accessed()

    if session.app_type != "Access":
        raise ValueError(
            f"export_form_definition only works with Access databases. "
            f"Got: {session.app_type}"
        )

    try:
        app = session.app

        # Verify form exists and get exact name
        actual_name = None
        for form in app.CurrentProject.AllForms:
            if form.Name.lower() == form_name.lower():
                actual_name = form.Name
                break

        if not actual_name:
            available = [f.Name for f in app.CurrentProject.AllForms]
            raise ValueError(
                f"Form '{form_name}' not found.\n"
                f"Available forms: {', '.join(available) if available else '(none)'}"
            )

        # Determine output path
        if output_path:
            export_path = Path(output_path).resolve()
        else:
            export_path = path.parent / f"{actual_name}.txt"

        # Ensure parent directory exists
        export_path.parent.mkdir(parents=True, exist_ok=True)

        # Export using SaveAsText
        # acForm = 2
        app.SaveAsText(2, actual_name, str(export_path))

        # Read content for preview
        try:
            with open(export_path, 'r', encoding='utf-8', errors='replace') as f:
                content = f.read()
                preview = content[:2000]
                if len(content) > 2000:
                    preview += "\n... [truncated]"
        except Exception as read_error:
            content = ""
            preview = f"(Could not read file: {read_error})"

        return f"""## Form Exported

**Form:** {actual_name}
**Database:** {path.name}
**Output:** {export_path}
**Size:** {len(content)} characters

### Preview

```
{preview}
```

**Next steps:**
1. Read the full file to see complete structure
2. Modify the content as needed
3. Re-import with `import_form_definition`
"""

    except ValueError:
        raise
    except Exception as e:
        raise RuntimeError(f"Error exporting form: {str(e)}")


async def import_form_definition_tool(
    file_path: str,
    form_name: str,
    definition_path: str,
    overwrite: bool = False
) -> str:
    """
    [PRO] Import Access form from text definition file (LoadFromText).

    This imports a form definition previously exported with SaveAsText
    or manually created/modified.

    Args:
        file_path: Path to .accdb or .mdb file
        form_name: Name for the imported form
        definition_path: Path to .txt definition file
        overwrite: If True, delete existing form first (default: False)

    Returns:
        Success message

    Raises:
        FileNotFoundError: If database or definition file doesn't exist
        ValueError: If form already exists and overwrite=False

    Workflow example:
        1. export_form_definition("db.accdb", "frm_Old")
        2. Read and modify the .txt file
        3. import_form_definition("db.accdb", "frm_New", "modified.txt")
    """
    path = Path(file_path).resolve()
    def_path = Path(definition_path).resolve()

    if not path.exists():
        raise FileNotFoundError(f"Database not found: {file_path}")

    if not def_path.exists():
        raise FileNotFoundError(f"Definition file not found: {definition_path}")

    manager = OfficeSessionManager.get_instance()
    session = await manager.get_or_create_session(path, read_only=False)
    session.refresh_last_accessed()

    if session.app_type != "Access":
        raise ValueError(
            f"import_form_definition only works with Access databases. "
            f"Got: {session.app_type}"
        )

    try:
        app = session.app

        # Check if form already exists
        existing_form = None
        for form in app.CurrentProject.AllForms:
            if form.Name.lower() == form_name.lower():
                existing_form = form.Name
                break

        if existing_form:
            if overwrite:
                # Delete existing form first
                # acForm = 2
                app.DoCmd.DeleteObject(2, existing_form)
            else:
                raise ValueError(
                    f"Form '{form_name}' already exists. "
                    f"Use overwrite=True to replace it."
                )

        # Import using LoadFromText
        # acForm = 2
        app.LoadFromText(2, form_name, str(def_path))

        return f"""## Form Imported

**Form:** {form_name}
**Database:** {path.name}
**Source:** {def_path}
**Overwrite:** {overwrite}

Form imported successfully.

**Next steps:**
- Open the form in Access to verify
- Use `list_access_forms` to confirm it's listed
- Use `export_form_definition` to re-export if needed
"""

    except ValueError:
        raise
    except Exception as e:
        raise RuntimeError(f"Error importing form: {str(e)}")
