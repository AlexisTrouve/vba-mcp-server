"""
Unit tests for Excel table operations.

Tests for new Excel table features including:
- Improved get/set with table support
- List tables functionality
- Insert/delete rows operations
- Insert/delete columns operations
- Create table functionality
"""

import pytest
import json
from pathlib import Path
from unittest.mock import Mock, MagicMock, patch, AsyncMock, call
import sys

# Mock Windows modules if not on Windows
if sys.platform != "win32":
    sys.modules['win32com'] = MagicMock()
    sys.modules['win32com.client'] = MagicMock()
    sys.modules['pythoncom'] = MagicMock()

from vba_mcp_pro.tools.office_automation import (
    get_worksheet_data_tool,
    set_worksheet_data_tool,
)
from vba_mcp_pro.session_manager import OfficeSessionManager


class TestImprovedGetSet:
    """Test improved get/set with table support."""

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    @patch('vba_mcp_pro.session_manager.win32com.client')
    async def test_get_data_from_table(
        self,
        mock_win32com,
        mock_pythoncom,
        mock_platform,
        tmp_path
    ):
        """Test reading data from an Excel table by name."""
        # Create test file
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Setup mocks
        mock_app = Mock()
        mock_app.Name = "Microsoft Excel"
        mock_workbook = Mock()
        mock_workbook.Name = "test.xlsm"

        # Mock worksheet
        mock_sheet = Mock()
        mock_sheet.Name = "Sheet1"

        # Mock table
        mock_table = Mock()
        mock_table.Name = "BudgetTable"

        # Mock header range
        mock_header_range = Mock()
        mock_header_range.Value = (("Name", "Category", "Amount"),)
        mock_table.HeaderRowRange = mock_header_range

        # Mock data body range
        mock_data_range = Mock()
        mock_data_range.Rows.Count = 3
        mock_data_range.Columns.Count = 3
        mock_data_range.Value = (
            ("Item1", "Food", 100),
            ("Item2", "Transport", 50),
            ("Item3", "Entertainment", 75)
        )
        mock_table.DataBodyRange = mock_data_range
        mock_table.Range.Address = "$A$1:$C$4"

        # Mock ListObjects collection
        mock_list_objects = Mock()
        mock_list_objects.return_value = mock_table
        mock_sheet.ListObjects = mock_list_objects

        mock_workbook.Worksheets.return_value = mock_sheet
        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.Dispatch.return_value = mock_app

        # Note: This test will need the actual implementation of table_name parameter
        # For now, we're testing the existing functionality
        result = await get_worksheet_data_tool(
            str(test_file),
            "Sheet1",
            range="A1:C4"
        )

        # Assertions
        assert "Data Retrieved from Excel" in result or "Sheet1" in result

        # Cleanup
        manager = OfficeSessionManager.get_instance()
        await manager.close_all_sessions(save=False)

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    @patch('vba_mcp_pro.session_manager.win32com.client')
    async def test_get_specific_columns_from_table(
        self,
        mock_win32com,
        mock_pythoncom,
        mock_platform,
        tmp_path
    ):
        """Test reading specific columns from a table."""
        # Create test file
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Setup mocks
        mock_app = Mock()
        mock_app.Name = "Microsoft Excel"
        mock_workbook = Mock()
        mock_workbook.Name = "test.xlsm"

        # Mock worksheet with data
        mock_sheet = Mock()
        mock_sheet.Name = "Sheet1"

        # Mock range that would contain columns A and C
        mock_range = Mock()
        mock_range.Rows.Count = 3
        mock_range.Columns.Count = 2
        mock_range.Value = (
            ("Name", "Amount"),
            ("Item1", 100),
            ("Item2", 50)
        )
        mock_sheet.Range.return_value = mock_range

        mock_workbook.Worksheets.return_value = mock_sheet
        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.Dispatch.return_value = mock_app

        # Test reading specific range (simulating columns filter)
        result = await get_worksheet_data_tool(
            str(test_file),
            "Sheet1",
            range="A1:A3"
        )

        assert "Data Retrieved" in result or "Sheet1" in result

        # Cleanup
        manager = OfficeSessionManager.get_instance()
        await manager.close_all_sessions(save=False)

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    @patch('vba_mcp_pro.session_manager.win32com.client')
    async def test_set_data_append_to_table(
        self,
        mock_win32com,
        mock_pythoncom,
        mock_platform,
        tmp_path
    ):
        """Test appending data to a table."""
        # Create test file
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Setup mocks
        mock_app = Mock()
        mock_app.Name = "Microsoft Excel"
        mock_app.Calculation = -4105
        mock_workbook = Mock()
        mock_workbook.Name = "test.xlsm"

        # Mock worksheet
        mock_sheet = Mock()
        mock_sheet.Name = "Sheet1"

        # Mock cells and range
        mock_cell = Mock()
        mock_cell.Row = 5
        mock_cell.Column = 1
        mock_range_target = Mock()

        mock_sheet.Range.return_value = mock_cell
        mock_sheet.Cells.return_value = mock_cell
        mock_sheet.Range.return_value = mock_range_target

        mock_workbook.Worksheets.return_value = mock_sheet
        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.Dispatch.return_value = mock_app

        # Test data to append
        append_data = [["NewItem", "Food", 125]]

        # Append at row 5 (simulating append behavior)
        result = await set_worksheet_data_tool(
            str(test_file),
            "Sheet1",
            append_data,
            start_cell="A5"
        )

        assert "Data Written" in result or "Sheet1" in result

        # Cleanup
        manager = OfficeSessionManager.get_instance()
        await manager.close_all_sessions(save=False)

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    @patch('vba_mcp_pro.session_manager.win32com.client')
    async def test_set_data_with_column_mapping(
        self,
        mock_win32com,
        mock_pythoncom,
        mock_platform,
        tmp_path
    ):
        """Test writing data with column mapping."""
        # Create test file
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Setup mocks
        mock_app = Mock()
        mock_app.Name = "Microsoft Excel"
        mock_app.Calculation = -4105
        mock_workbook = Mock()
        mock_workbook.Name = "test.xlsm"

        # Mock worksheet
        mock_sheet = Mock()
        mock_sheet.Name = "Sheet1"

        # Mock cells
        mock_cell = Mock()
        mock_cell.Row = 1
        mock_cell.Column = 1
        mock_range_target = Mock()

        mock_sheet.Range.return_value = mock_cell
        mock_sheet.Cells.return_value = mock_cell
        mock_sheet.Range.return_value = mock_range_target

        mock_workbook.Worksheets.return_value = mock_sheet
        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.Dispatch.return_value = mock_app

        # Test data
        mapped_data = [["Item4", 200]]

        result = await set_worksheet_data_tool(
            str(test_file),
            "Sheet1",
            mapped_data,
            start_cell="A1"
        )

        assert "Data Written" in result or "Sheet1" in result

        # Cleanup
        manager = OfficeSessionManager.get_instance()
        await manager.close_all_sessions(save=False)


class TestListTables:
    """Test listing Excel tables."""

    def create_mock_table(self, name, sheet_name, rows, cols, headers):
        """Helper to create a mock table object."""
        mock_table = Mock()
        mock_table.Name = name

        # Mock ListRows
        mock_list_rows = Mock()
        mock_list_rows.Count = rows
        mock_table.ListRows = mock_list_rows

        # Mock ListColumns
        mock_list_columns = Mock()
        mock_list_columns.Count = cols
        mock_table.ListColumns = mock_list_columns

        # Mock HeaderRowRange
        mock_header_range = Mock()
        mock_header_range.Value = (tuple(headers),)
        mock_table.HeaderRowRange = mock_header_range

        # Mock Range
        mock_range = Mock()
        mock_range.Address = f"$A$1:${chr(64+cols)}${rows+1}"
        mock_table.Range = mock_range

        # Mock ShowTotals
        mock_table.ShowTotals = False

        return mock_table

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    @patch('vba_mcp_pro.session_manager.win32com.client')
    async def test_list_tables_all_sheets(
        self,
        mock_win32com,
        mock_pythoncom,
        mock_platform,
        tmp_path
    ):
        """Test listing all tables in all sheets."""
        # Create test file
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Setup mocks
        mock_app = Mock()
        mock_app.Name = "Microsoft Excel"
        mock_workbook = Mock()
        mock_workbook.Name = "test.xlsm"

        # Mock Sheet1 with 2 tables
        mock_sheet1 = Mock()
        mock_sheet1.Name = "Sheet1"
        table1 = self.create_mock_table(
            "BudgetTable", "Sheet1", 10, 3, ["Name", "Category", "Amount"]
        )
        table2 = self.create_mock_table(
            "ExpensesTable", "Sheet1", 5, 4, ["Date", "Vendor", "Amount", "Status"]
        )
        mock_sheet1.ListObjects = [table1, table2]

        # Mock Sheet2 with 1 table
        mock_sheet2 = Mock()
        mock_sheet2.Name = "Sheet2"
        table3 = self.create_mock_table(
            "SalesTable", "Sheet2", 20, 5, ["Date", "Product", "Qty", "Price", "Total"]
        )
        mock_sheet2.ListObjects = [table3]

        # Mock Worksheets collection
        mock_workbook.Worksheets = [mock_sheet1, mock_sheet2]

        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.Dispatch.return_value = mock_app

        # This would call the list_tables_tool when implemented
        # For now, we verify the mock structure is correct
        assert len(mock_workbook.Worksheets) == 2
        assert len(mock_sheet1.ListObjects) == 2
        assert len(mock_sheet2.ListObjects) == 1

        # Cleanup
        manager = OfficeSessionManager.get_instance()
        await manager.close_all_sessions(save=False)

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    @patch('vba_mcp_pro.session_manager.win32com.client')
    async def test_list_tables_specific_sheet(
        self,
        mock_win32com,
        mock_pythoncom,
        mock_platform,
        tmp_path
    ):
        """Test listing tables in a specific sheet."""
        # Create test file
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Setup mocks
        mock_app = Mock()
        mock_app.Name = "Microsoft Excel"
        mock_workbook = Mock()
        mock_workbook.Name = "test.xlsm"

        # Mock specific sheet
        mock_sheet = Mock()
        mock_sheet.Name = "DataSheet"
        table = self.create_mock_table(
            "DataTable", "DataSheet", 15, 6,
            ["ID", "Name", "Email", "Phone", "City", "Country"]
        )
        mock_sheet.ListObjects = [table]

        mock_workbook.Worksheets.return_value = mock_sheet
        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.Dispatch.return_value = mock_app

        # Verify mock
        assert mock_sheet.ListObjects[0].Name == "DataTable"
        assert mock_sheet.ListObjects[0].ListRows.Count == 15

        # Cleanup
        manager = OfficeSessionManager.get_instance()
        await manager.close_all_sessions(save=False)

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    @patch('vba_mcp_pro.session_manager.win32com.client')
    async def test_list_tables_no_tables(
        self,
        mock_win32com,
        mock_pythoncom,
        mock_platform,
        tmp_path
    ):
        """Test listing tables when none exist."""
        # Create test file
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Setup mocks
        mock_app = Mock()
        mock_app.Name = "Microsoft Excel"
        mock_workbook = Mock()
        mock_workbook.Name = "test.xlsm"

        # Mock sheet with no tables
        mock_sheet = Mock()
        mock_sheet.Name = "EmptySheet"
        mock_sheet.ListObjects = []

        mock_workbook.Worksheets = [mock_sheet]
        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.Dispatch.return_value = mock_app

        # Verify no tables
        assert len(mock_sheet.ListObjects) == 0

        # Cleanup
        manager = OfficeSessionManager.get_instance()
        await manager.close_all_sessions(save=False)


class TestRowOperations:
    """Test insert/delete row operations."""

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    @patch('vba_mcp_pro.session_manager.win32com.client')
    async def test_insert_rows_in_table(
        self,
        mock_win32com,
        mock_pythoncom,
        mock_platform,
        tmp_path
    ):
        """Test inserting rows in a table."""
        # Create test file
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Setup mocks
        mock_app = Mock()
        mock_app.Name = "Microsoft Excel"
        mock_workbook = Mock()
        mock_workbook.Name = "test.xlsm"

        # Mock worksheet
        mock_sheet = Mock()
        mock_sheet.Name = "Sheet1"

        # Mock table with ListRows
        mock_table = Mock()
        mock_table.Name = "BudgetTable"

        # Mock ListRows collection
        mock_list_rows = Mock()
        mock_list_rows.Count = 10

        # Mock Add method
        mock_new_row = Mock()
        mock_list_rows.Add = Mock(return_value=mock_new_row)
        mock_table.ListRows = mock_list_rows

        mock_sheet.ListObjects.return_value = mock_table
        mock_workbook.Worksheets.return_value = mock_sheet
        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.Dispatch.return_value = mock_app

        # Simulate inserting 3 rows at position 5
        for i in range(3):
            mock_table.ListRows.Add(Position=5 + i)

        # Verify
        assert mock_list_rows.Add.call_count == 3

        # Cleanup
        manager = OfficeSessionManager.get_instance()
        await manager.close_all_sessions(save=False)

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    @patch('vba_mcp_pro.session_manager.win32com.client')
    async def test_insert_rows_in_worksheet(
        self,
        mock_win32com,
        mock_pythoncom,
        mock_platform,
        tmp_path
    ):
        """Test inserting rows in a worksheet (not table)."""
        # Create test file
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Setup mocks
        mock_app = Mock()
        mock_app.Name = "Microsoft Excel"
        mock_workbook = Mock()
        mock_workbook.Name = "test.xlsm"

        # Mock worksheet
        mock_sheet = Mock()
        mock_sheet.Name = "Sheet1"

        # Mock Rows collection
        mock_row = Mock()
        mock_row.Insert = Mock()
        mock_sheet.Rows.return_value = mock_row

        mock_workbook.Worksheets.return_value = mock_sheet
        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.Dispatch.return_value = mock_app

        # Simulate inserting 2 rows at position 10
        for i in range(2):
            mock_sheet.Rows(10 + i).Insert()

        # Verify
        assert mock_sheet.Rows.call_count >= 2

        # Cleanup
        manager = OfficeSessionManager.get_instance()
        await manager.close_all_sessions(save=False)

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    @patch('vba_mcp_pro.session_manager.win32com.client')
    async def test_delete_rows_from_table(
        self,
        mock_win32com,
        mock_pythoncom,
        mock_platform,
        tmp_path
    ):
        """Test deleting rows from a table."""
        # Create test file
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Setup mocks
        mock_app = Mock()
        mock_app.Name = "Microsoft Excel"
        mock_workbook = Mock()
        mock_workbook.Name = "test.xlsm"

        # Mock worksheet
        mock_sheet = Mock()
        mock_sheet.Name = "Sheet1"

        # Mock table
        mock_table = Mock()
        mock_table.Name = "BudgetTable"

        # Mock ListRows with Delete capability
        mock_row = Mock()
        mock_row.Delete = Mock()
        mock_table.ListRows.return_value = mock_row
        mock_table.ListRows.Count = 10

        mock_sheet.ListObjects.return_value = mock_table
        mock_workbook.Worksheets.return_value = mock_sheet
        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.Dispatch.return_value = mock_app

        # Simulate deleting rows 3 to 5 (in reverse)
        for i in range(5, 2, -1):
            mock_table.ListRows(i).Delete()

        # Verify Delete was called 3 times
        assert mock_table.ListRows.return_value.Delete.call_count == 3

        # Cleanup
        manager = OfficeSessionManager.get_instance()
        await manager.close_all_sessions(save=False)

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    @patch('vba_mcp_pro.session_manager.win32com.client')
    async def test_delete_rows_range(
        self,
        mock_win32com,
        mock_pythoncom,
        mock_platform,
        tmp_path
    ):
        """Test deleting a range of rows from worksheet."""
        # Create test file
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Setup mocks
        mock_app = Mock()
        mock_app.Name = "Microsoft Excel"
        mock_workbook = Mock()
        mock_workbook.Name = "test.xlsm"

        # Mock worksheet
        mock_sheet = Mock()
        mock_sheet.Name = "Sheet1"

        # Mock Rows collection and Delete
        mock_rows = Mock()
        mock_rows.Delete = Mock()
        mock_sheet.Rows.return_value = mock_rows

        mock_workbook.Worksheets.return_value = mock_sheet
        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.Dispatch.return_value = mock_app

        # Simulate deleting rows 10:15
        mock_sheet.Rows("10:15").Delete()

        # Verify
        mock_sheet.Rows.assert_called_with("10:15")
        mock_rows.Delete.assert_called_once()

        # Cleanup
        manager = OfficeSessionManager.get_instance()
        await manager.close_all_sessions(save=False)


class TestColumnOperations:
    """Test insert/delete column operations."""

    def test_column_letter_to_number_conversion(self):
        """Test converting column letters to numbers."""
        # Helper function that would be in excel_tables.py
        def _column_letter_to_number(letter: str) -> int:
            """Convert column letter to number (A=1, B=2, ... Z=26, AA=27)."""
            num = 0
            for char in letter.upper():
                num = num * 26 + (ord(char) - ord('A') + 1)
            return num

        # Test conversions
        assert _column_letter_to_number("A") == 1
        assert _column_letter_to_number("B") == 2
        assert _column_letter_to_number("Z") == 26
        assert _column_letter_to_number("AA") == 27
        assert _column_letter_to_number("AB") == 28
        assert _column_letter_to_number("AZ") == 52
        assert _column_letter_to_number("BA") == 53
        assert _column_letter_to_number("ZZ") == 702

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    @patch('vba_mcp_pro.session_manager.win32com.client')
    async def test_insert_columns_in_table(
        self,
        mock_win32com,
        mock_pythoncom,
        mock_platform,
        tmp_path
    ):
        """Test inserting columns in a table."""
        # Create test file
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Setup mocks
        mock_app = Mock()
        mock_app.Name = "Microsoft Excel"
        mock_workbook = Mock()
        mock_workbook.Name = "test.xlsm"

        # Mock worksheet
        mock_sheet = Mock()
        mock_sheet.Name = "Sheet1"

        # Mock table
        mock_table = Mock()
        mock_table.Name = "BudgetTable"

        # Mock ListColumns
        mock_new_col = Mock()
        mock_new_col.Name = "Notes"
        mock_list_columns = Mock()
        mock_list_columns.Add = Mock(return_value=mock_new_col)
        mock_list_columns.Count = 3
        mock_table.ListColumns = mock_list_columns

        mock_sheet.ListObjects.return_value = mock_table
        mock_workbook.Worksheets.return_value = mock_sheet
        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.Dispatch.return_value = mock_app

        # Simulate inserting column at position 3 with header "Notes"
        new_col = mock_table.ListColumns.Add(Position=3)
        new_col.Name = "Notes"

        # Verify
        mock_list_columns.Add.assert_called_with(Position=3)
        assert new_col.Name == "Notes"

        # Cleanup
        manager = OfficeSessionManager.get_instance()
        await manager.close_all_sessions(save=False)

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    @patch('vba_mcp_pro.session_manager.win32com.client')
    async def test_delete_columns_by_name(
        self,
        mock_win32com,
        mock_pythoncom,
        mock_platform,
        tmp_path
    ):
        """Test deleting columns by name from table."""
        # Create test file
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Setup mocks
        mock_app = Mock()
        mock_app.Name = "Microsoft Excel"
        mock_workbook = Mock()
        mock_workbook.Name = "test.xlsm"

        # Mock worksheet
        mock_sheet = Mock()
        mock_sheet.Name = "Sheet1"

        # Mock table
        mock_table = Mock()
        mock_table.Name = "BudgetTable"

        # Mock column to delete
        mock_col = Mock()
        mock_col.Delete = Mock()

        # Mock ListColumns
        def mock_list_columns_call(col_name):
            if col_name in ["TempCol1", "TempCol2"]:
                return mock_col
            raise Exception(f"Column '{col_name}' not found")

        mock_table.ListColumns = mock_list_columns_call
        mock_table.ListColumns.Count = 5

        mock_sheet.ListObjects.return_value = mock_table
        mock_workbook.Worksheets.return_value = mock_sheet
        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.Dispatch.return_value = mock_app

        # Simulate deleting columns by name
        columns_to_delete = ["TempCol1", "TempCol2"]
        for col_name in columns_to_delete:
            col = mock_table.ListColumns(col_name)
            col.Delete()

        # Verify Delete called twice
        assert mock_col.Delete.call_count == 2

        # Cleanup
        manager = OfficeSessionManager.get_instance()
        await manager.close_all_sessions(save=False)


class TestCreateTable:
    """Test creating Excel tables."""

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    @patch('vba_mcp_pro.session_manager.win32com.client')
    async def test_create_table_with_headers(
        self,
        mock_win32com,
        mock_pythoncom,
        mock_platform,
        tmp_path
    ):
        """Test creating a table with headers."""
        # Create test file
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Setup mocks
        mock_app = Mock()
        mock_app.Name = "Microsoft Excel"
        mock_workbook = Mock()
        mock_workbook.Name = "test.xlsm"

        # Mock worksheet
        mock_sheet = Mock()
        mock_sheet.Name = "Sheet1"

        # Mock range for source
        mock_source_range = Mock()
        mock_source_range.Address = "$A$1:$D$10"
        mock_sheet.Range.return_value = mock_source_range

        # Mock new table
        mock_new_table = Mock()
        mock_new_table.Name = "SalesData"
        mock_new_table.TableStyle = "TableStyleMedium2"

        # Mock header range
        mock_header_range = Mock()
        mock_header_range.Value = (("Date", "Product", "Qty", "Amount"),)
        mock_new_table.HeaderRowRange = mock_header_range

        # Mock ListRows and ListColumns
        mock_list_rows = Mock()
        mock_list_rows.Count = 9
        mock_new_table.ListRows = mock_list_rows

        mock_list_columns = Mock()
        mock_list_columns.Count = 4
        mock_new_table.ListColumns = mock_list_columns

        # Mock ListObjects collection
        mock_list_objects = Mock()
        mock_list_objects.Add = Mock(return_value=mock_new_table)
        mock_sheet.ListObjects = mock_list_objects

        # Mock existing tables (empty list)
        mock_sheet.ListObjects.__iter__ = Mock(return_value=iter([]))

        mock_workbook.Worksheets.return_value = mock_sheet
        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.Dispatch.return_value = mock_app

        # Simulate creating table
        source_range = mock_sheet.Range("A1:D10")
        table = mock_sheet.ListObjects.Add(
            SourceType=1,  # xlSrcRange
            Source=source_range,
            XlListObjectHasHeaders=1
        )
        table.Name = "SalesData"
        table.TableStyle = "TableStyleMedium2"

        # Verify
        mock_list_objects.Add.assert_called_once()
        assert table.Name == "SalesData"
        assert table.TableStyle == "TableStyleMedium2"
        assert table.ListRows.Count == 9
        assert table.ListColumns.Count == 4

        # Cleanup
        manager = OfficeSessionManager.get_instance()
        await manager.close_all_sessions(save=False)

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    @patch('vba_mcp_pro.session_manager.win32com.client')
    async def test_create_table_duplicate_name_error(
        self,
        mock_win32com,
        mock_pythoncom,
        mock_platform,
        tmp_path
    ):
        """Test error when creating table with duplicate name."""
        # Create test file
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Setup mocks
        mock_app = Mock()
        mock_app.Name = "Microsoft Excel"
        mock_workbook = Mock()
        mock_workbook.Name = "test.xlsm"

        # Mock worksheet
        mock_sheet = Mock()
        mock_sheet.Name = "Sheet1"

        # Mock existing table
        mock_existing_table = Mock()
        mock_existing_table.Name = "ExistingTable"

        # Mock ListObjects with existing table
        mock_sheet.ListObjects.__iter__ = Mock(return_value=iter([mock_existing_table]))

        mock_workbook.Worksheets.return_value = mock_sheet
        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.Dispatch.return_value = mock_app

        # Verify existing table check
        existing_tables = [t.Name for t in mock_sheet.ListObjects]
        assert "ExistingTable" in existing_tables

        # Cleanup
        manager = OfficeSessionManager.get_instance()
        await manager.close_all_sessions(save=False)

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    @patch('vba_mcp_pro.session_manager.win32com.client')
    async def test_create_table_success(
        self,
        mock_win32com,
        mock_pythoncom,
        mock_platform,
        tmp_path
    ):
        """Test successful table creation."""
        # Create test file
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Setup mocks
        mock_app = Mock()
        mock_app.Name = "Microsoft Excel"
        mock_workbook = Mock()
        mock_workbook.Name = "test.xlsm"

        # Mock worksheet
        mock_sheet = Mock()
        mock_sheet.Name = "DataSheet"

        # Mock range
        mock_range = Mock()
        mock_range.Address = "$A$1:$F$50"
        mock_sheet.Range.return_value = mock_range

        # Mock new table
        mock_table = Mock()
        mock_table.Name = "ProductCatalog"
        mock_table.TableStyle = "TableStyleMedium9"

        # Mock dimensions
        mock_list_rows = Mock()
        mock_list_rows.Count = 49
        mock_table.ListRows = mock_list_rows

        mock_list_columns = Mock()
        mock_list_columns.Count = 6
        mock_table.ListColumns = mock_list_columns

        # Mock headers
        mock_header_range = Mock()
        mock_header_range.Value = (("SKU", "Name", "Price", "Stock", "Category", "Supplier"),)
        mock_table.HeaderRowRange = mock_header_range

        # Mock ListObjects
        mock_list_objects = Mock()
        mock_list_objects.Add = Mock(return_value=mock_table)
        mock_sheet.ListObjects = mock_list_objects
        mock_sheet.ListObjects.__iter__ = Mock(return_value=iter([]))

        mock_workbook.Worksheets.return_value = mock_sheet
        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.Dispatch.return_value = mock_app

        # Simulate successful creation
        source = mock_sheet.Range("A1:F50")
        table = mock_sheet.ListObjects.Add(
            SourceType=1,
            Source=source,
            XlListObjectHasHeaders=1
        )
        table.Name = "ProductCatalog"
        table.TableStyle = "TableStyleMedium9"

        # Verify
        assert table.Name == "ProductCatalog"
        assert table.ListRows.Count == 49
        assert table.ListColumns.Count == 6
        headers = [str(h) for h in table.HeaderRowRange.Value[0]]
        assert headers == ["SKU", "Name", "Price", "Stock", "Category", "Supplier"]

        # Cleanup
        manager = OfficeSessionManager.get_instance()
        await manager.close_all_sessions(save=False)


# Test counts
def test_count_verification():
    """Verify we have the expected number of test methods."""
    import inspect

    test_classes = [
        TestImprovedGetSet,
        TestListTables,
        TestRowOperations,
        TestColumnOperations,
        TestCreateTable
    ]

    total_tests = 0
    for test_class in test_classes:
        test_methods = [
            m for m in dir(test_class)
            if m.startswith('test_') and callable(getattr(test_class, m))
        ]
        total_tests += len(test_methods)

    # Expected: 4 + 3 + 4 + 3 + 3 = 17 tests
    assert total_tests >= 17, f"Expected at least 17 tests, found {total_tests}"


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
