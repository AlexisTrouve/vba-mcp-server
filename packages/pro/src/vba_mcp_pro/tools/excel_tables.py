"""
Excel Table Management Tools (PRO)

Tools for managing Excel Tables (ListObjects) and row/column operations.
Provides interactive table management with visible Excel sessions.
"""

import logging
from pathlib import Path
from typing import Any, List, Optional, Union

from ..session_manager import OfficeSessionManager


# Configure logging
logger = logging.getLogger(__name__)


async def list_tables_tool(
    file_path: str,
    sheet_name: Optional[str] = None
) -> str:
    """
    List all Excel Tables (ListObjects) in file or specific sheet.

    Args:
        file_path: Absolute path to Excel file
        sheet_name: Optional sheet name to filter tables

    Returns:
        Formatted markdown string with table information

    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If sheet doesn't exist
    """
    manager = OfficeSessionManager.get_instance()
    session = await manager.get_or_create_session(Path(file_path).resolve())

    tables_info = []

    if sheet_name:
        try:
            sheets = [session.file_obj.Worksheets(sheet_name)]
        except Exception as e:
            raise ValueError(f"Sheet '{sheet_name}' not found: {e}")
    else:
        sheets = session.file_obj.Worksheets

    for ws in sheets:
        for table in ws.ListObjects:
            # Get headers safely
            try:
                headers_raw = table.HeaderRowRange.Value
                if headers_raw:
                    if isinstance(headers_raw[0], tuple):
                        headers = [str(h) for h in headers_raw[0]]
                    else:
                        headers = [str(headers_raw[0])]
                else:
                    headers = []
            except Exception:
                headers = []

            info = {
                "name": table.Name,
                "sheet": ws.Name,
                "rows": table.ListRows.Count,
                "columns": table.ListColumns.Count,
                "headers": headers,
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
        if table['headers']:
            output += f"- **Columns:** {', '.join(table['headers'])}\n"
        output += f"- **Range:** {table['range']}\n"
        output += f"- **Total Row:** {'Enabled' if table['total_row'] else 'Disabled'}\n\n"

    return output


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
        file_path: Absolute path to Excel file
        sheet_name: Name of the worksheet
        position: Row number (1-based) or relative position in table
        count: Number of rows to insert (default: 1)
        table_name: If specified, insert in table context

    Returns:
        Success message with operation details

    Raises:
        ValueError: If sheet or table doesn't exist
    """
    manager = OfficeSessionManager.get_instance()
    session = await manager.get_or_create_session(Path(file_path).resolve())

    try:
        ws = session.file_obj.Worksheets(sheet_name)
    except Exception as e:
        raise ValueError(f"Sheet '{sheet_name}' not found: {e}")

    if table_name:
        # Insert in table
        try:
            table = ws.ListObjects(table_name)
        except Exception as e:
            raise ValueError(f"Table '{table_name}' not found in sheet '{sheet_name}': {e}")

        for i in range(count):
            table.ListRows.Add(Position=position + i)

        output = f"### Rows Inserted in Table\n\n"
        output += f"**Table:** {table_name}\n"
        output += f"**Sheet:** {sheet_name}\n"
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
        file_path: Absolute path to Excel file
        sheet_name: Name of the worksheet
        start_row: First row to delete (1-based)
        end_row: Last row to delete (inclusive). If None, delete only start_row
        table_name: If specified, delete from table

    Returns:
        Success message with operation details

    Raises:
        ValueError: If sheet or table doesn't exist
    """
    if end_row is None:
        end_row = start_row

    count = end_row - start_row + 1

    manager = OfficeSessionManager.get_instance()
    session = await manager.get_or_create_session(Path(file_path).resolve())

    try:
        ws = session.file_obj.Worksheets(sheet_name)
    except Exception as e:
        raise ValueError(f"Sheet '{sheet_name}' not found: {e}")

    if table_name:
        try:
            table = ws.ListObjects(table_name)
        except Exception as e:
            raise ValueError(f"Table '{table_name}' not found in sheet '{sheet_name}': {e}")

        # Delete in reverse to avoid index shifting
        for i in range(end_row, start_row - 1, -1):
            try:
                table.ListRows(i).Delete()
            except Exception as e:
                logger.warning(f"Failed to delete row {i}: {e}")

        output = f"### Rows Deleted from Table\n\n"
        output += f"**Table:** {table_name}\n"
        output += f"**Sheet:** {sheet_name}\n"
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


def _column_letter_to_number(letter: str) -> int:
    """
    Convert column letter to number (A=1, B=2, ... Z=26, AA=27).

    Args:
        letter: Column letter (e.g., "A", "AB", "ZZ")

    Returns:
        Column number (1-based)
    """
    num = 0
    for char in letter.upper():
        num = num * 26 + (ord(char) - ord('A') + 1)
    return num


async def insert_columns_tool(
    file_path: str,
    sheet_name: str,
    position: Union[int, str],
    count: int = 1,
    table_name: Optional[str] = None,
    header_name: Optional[str] = None
) -> str:
    """
    Insert column(s) in worksheet or table.

    Args:
        file_path: Absolute path to Excel file
        sheet_name: Name of the worksheet
        position: Column number (1-based) or letter ("A", "B", etc.)
        count: Number of columns to insert (default: 1)
        table_name: If specified, insert in table
        header_name: New column header (for tables)

    Returns:
        Success message with operation details

    Raises:
        ValueError: If sheet or table doesn't exist
    """
    manager = OfficeSessionManager.get_instance()
    session = await manager.get_or_create_session(Path(file_path).resolve())

    try:
        ws = session.file_obj.Worksheets(sheet_name)
    except Exception as e:
        raise ValueError(f"Sheet '{sheet_name}' not found: {e}")

    # Convert column letter to number if needed
    if isinstance(position, str):
        position = _column_letter_to_number(position)

    if table_name:
        try:
            table = ws.ListObjects(table_name)
        except Exception as e:
            raise ValueError(f"Table '{table_name}' not found in sheet '{sheet_name}': {e}")

        for i in range(count):
            col = table.ListColumns.Add(Position=position + i)
            if header_name:
                col.Name = f"{header_name}_{i+1}" if count > 1 else header_name

        output = f"### Columns Inserted in Table\n\n"
        output += f"**Table:** {table_name}\n"
        output += f"**Sheet:** {sheet_name}\n"
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


async def delete_columns_tool(
    file_path: str,
    sheet_name: str,
    column: Union[int, str, List[str]],
    table_name: Optional[str] = None
) -> str:
    """
    Delete column(s) from worksheet or table.

    Args:
        file_path: Absolute path to Excel file
        sheet_name: Name of the worksheet
        column: Column number, letter, or list of column names (for tables)
        table_name: If specified, delete from table

    Returns:
        Success message with operation details

    Raises:
        ValueError: If sheet or table doesn't exist
    """
    manager = OfficeSessionManager.get_instance()
    session = await manager.get_or_create_session(Path(file_path).resolve())

    try:
        ws = session.file_obj.Worksheets(sheet_name)
    except Exception as e:
        raise ValueError(f"Sheet '{sheet_name}' not found: {e}")

    if table_name:
        try:
            table = ws.ListObjects(table_name)
        except Exception as e:
            raise ValueError(f"Table '{table_name}' not found in sheet '{sheet_name}': {e}")

        if isinstance(column, list):
            # Delete by column names
            for col_name in column:
                try:
                    table.ListColumns(col_name).Delete()
                except Exception as e:
                    logger.warning(f"Failed to delete column '{col_name}': {e}")

            output = f"### Columns Deleted from Table\n\n"
            output += f"**Table:** {table_name}\n"
            output += f"**Sheet:** {sheet_name}\n"
            output += f"**Columns:** {', '.join(column)} ({len(column)} column(s))\n"
        else:
            # Delete by position
            if isinstance(column, str):
                column = _column_letter_to_number(column)
            table.ListColumns(column).Delete()

            output = f"### Column Deleted from Table\n\n"
            output += f"**Table:** {table_name}\n"
            output += f"**Sheet:** {sheet_name}\n"
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
        file_path: Absolute path to Excel file
        sheet_name: Name of the worksheet
        range: Range to convert (e.g., "A1:D10")
        table_name: Name for the new table
        has_headers: First row contains headers (default: True)
        style: Excel table style name (default: "TableStyleMedium2")

    Returns:
        Success message with table details

    Raises:
        ValueError: If sheet doesn't exist or table name is duplicate
    """
    manager = OfficeSessionManager.get_instance()
    session = await manager.get_or_create_session(Path(file_path).resolve())

    try:
        ws = session.file_obj.Worksheets(sheet_name)
    except Exception as e:
        raise ValueError(f"Sheet '{sheet_name}' not found: {e}")

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

    # Apply style
    try:
        table.TableStyle = style
    except Exception as e:
        logger.warning(f"Failed to apply style '{style}': {e}")

    output = f"### Excel Table Created\n\n"
    output += f"**Name:** {table_name}\n"
    output += f"**Sheet:** {sheet_name}\n"
    output += f"**Range:** {range}\n"
    output += f"**Headers:** {'Yes' if has_headers else 'No'}\n"
    output += f"**Size:** {table.ListRows.Count} rows × {table.ListColumns.Count} columns\n"
    output += f"**Style:** {style}\n"

    if has_headers:
        try:
            headers_raw = table.HeaderRowRange.Value
            if headers_raw:
                if isinstance(headers_raw[0], tuple):
                    headers = [str(h) for h in headers_raw[0]]
                else:
                    headers = [str(headers_raw[0])]
                output += f"**Columns:** {', '.join(headers)}\n"
        except Exception as e:
            logger.warning(f"Failed to read headers: {e}")

    return output
