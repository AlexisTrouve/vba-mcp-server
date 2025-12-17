"""
Unit tests for Office automation tools.
"""

import pytest
import json
from pathlib import Path
from unittest.mock import Mock, MagicMock, patch, AsyncMock
import sys

# Mock Windows modules if not on Windows
if sys.platform != "win32":
    sys.modules['win32com'] = MagicMock()
    sys.modules['win32com.client'] = MagicMock()
    sys.modules['pythoncom'] = MagicMock()

from vba_mcp_pro.tools.office_automation import (
    open_in_office_tool,
    close_office_file_tool,
    list_open_files_tool,
    run_macro_tool,
    get_worksheet_data_tool,
    set_worksheet_data_tool,
    _normalize_com_data,
    _validate_data,
    _get_cell_address
)
from vba_mcp_pro.session_manager import OfficeSessionManager


class TestOpenInOfficeTool:
    """Tests for open_in_office_tool."""

    @pytest.mark.asyncio
    async def test_open_file_not_found(self):
        """Test error when file doesn't exist."""
        with pytest.raises(FileNotFoundError):
            await open_in_office_tool("/nonexistent/file.xlsm")

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Linux')
    async def test_open_non_windows(self, mock_platform, tmp_path):
        """Test error when not on Windows."""
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        with pytest.raises(RuntimeError, match="Windows"):
            await open_in_office_tool(str(test_file))

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    @patch('vba_mcp_pro.session_manager.win32com.client')
    async def test_open_excel_file(
        self,
        mock_win32com,
        mock_pythoncom,
        mock_platform,
        tmp_path
    ):
        """Test opening an Excel file."""
        # Create test file
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Setup mocks
        mock_app = Mock()
        mock_app.Name = "Microsoft Excel"
        mock_workbook = Mock()
        mock_workbook.Name = "test.xlsm"
        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.Dispatch.return_value = mock_app

        # Call tool
        result = await open_in_office_tool(str(test_file), read_only=False)

        # Assertions
        assert "Office File Opened" in result
        assert "test.xlsm" in result
        assert "Excel" in result
        assert "Editable" in result
        assert "Visible" in result

        # Cleanup
        manager = OfficeSessionManager.get_instance()
        await manager.close_all_sessions(save=False)


class TestCloseOfficeFileTool:
    """Tests for close_office_file_tool."""

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    @patch('vba_mcp_pro.session_manager.win32com.client')
    async def test_close_existing_session(
        self,
        mock_win32com,
        mock_pythoncom,
        mock_platform,
        tmp_path
    ):
        """Test closing an existing session."""
        # Create test file
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Setup mocks
        mock_app = Mock()
        mock_app.Name = "Microsoft Excel"
        mock_workbook = Mock()
        mock_workbook.Name = "test.xlsm"
        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.Dispatch.return_value = mock_app

        # Open first
        await open_in_office_tool(str(test_file))

        # Close
        result = await close_office_file_tool(str(test_file), save_changes=True)

        # Assertions
        assert "Office File Closed" in result
        assert "test.xlsm" in result
        assert "Saved" in result
        mock_workbook.Save.assert_called_once()
        mock_app.Quit.assert_called_once()

    @pytest.mark.asyncio
    async def test_close_nonexistent_session(self, tmp_path):
        """Test error when closing non-existent session."""
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        with pytest.raises(ValueError, match="No open session"):
            await close_office_file_tool(str(test_file))


class TestListOpenFilesTool:
    """Tests for list_open_files_tool."""

    @pytest.mark.asyncio
    async def test_list_no_sessions(self):
        """Test listing when no sessions open."""
        # Reset manager
        manager = OfficeSessionManager.get_instance()
        await manager.close_all_sessions(save=False)

        result = await list_open_files_tool()
        assert "No Office Files Currently Open" in result

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    @patch('vba_mcp_pro.session_manager.win32com.client')
    async def test_list_with_sessions(
        self,
        mock_win32com,
        mock_pythoncom,
        mock_platform,
        tmp_path
    ):
        """Test listing with open sessions."""
        # Create test files
        file1 = tmp_path / "test1.xlsm"
        file1.write_text("test")
        file2 = tmp_path / "test2.xlsm"
        file2.write_text("test")

        # Setup mocks
        mock_app = Mock()
        mock_app.Name = "Microsoft Excel"
        mock_workbook = Mock()
        mock_workbook.Name = "test.xlsm"
        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.Dispatch.return_value = mock_app

        # Open both files
        await open_in_office_tool(str(file1))
        await open_in_office_tool(str(file2))

        # List
        result = await list_open_files_tool()

        # Assertions
        assert "Open Office File Sessions" in result
        assert "Total sessions: 2" in result
        assert "test1.xlsm" in result
        assert "test2.xlsm" in result

        # Cleanup
        manager = OfficeSessionManager.get_instance()
        await manager.close_all_sessions(save=False)


class TestRunMacroTool:
    """Tests for run_macro_tool."""

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    @patch('vba_mcp_pro.session_manager.win32com.client')
    async def test_run_macro_success(
        self,
        mock_win32com,
        mock_pythoncom,
        mock_platform,
        tmp_path
    ):
        """Test successful macro execution."""
        # Create test file
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Setup mocks
        mock_app = Mock()
        mock_app.Name = "Microsoft Excel"
        mock_app.Run.return_value = 42  # Return value from Function
        mock_workbook = Mock()
        mock_workbook.Name = "test.xlsm"
        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.Dispatch.return_value = mock_app

        # Run macro
        result = await run_macro_tool(
            str(test_file),
            "Module1.Calculate",
            arguments=[10, 20]
        )

        # Assertions
        assert "Macro Executed Successfully" in result
        assert "Module1.Calculate" in result
        assert "Return value: 42" in result
        mock_app.Run.assert_called_once_with("Module1.Calculate", 10, 20)

        # Cleanup
        manager = OfficeSessionManager.get_instance()
        await manager.close_all_sessions(save=False)

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    @patch('vba_mcp_pro.session_manager.win32com.client')
    async def test_run_macro_not_found(
        self,
        mock_win32com,
        mock_pythoncom,
        mock_platform,
        tmp_path
    ):
        """Test error when macro not found."""
        # Create test file
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Setup mocks
        mock_app = Mock()
        mock_app.Name = "Microsoft Excel"

        # Simulate COM error for macro not found
        import pythoncom as real_pythoncom
        mock_app.Run.side_effect = Exception("cannot run the macro")

        mock_workbook = Mock()
        mock_workbook.Name = "test.xlsm"
        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.Dispatch.return_value = mock_app

        # Mock pythoncom.com_error
        mock_pythoncom.com_error = Exception

        # Run macro
        with pytest.raises(ValueError, match="not found"):
            await run_macro_tool(str(test_file), "NonExistent.Macro")

        # Cleanup
        manager = OfficeSessionManager.get_instance()
        await manager.close_all_sessions(save=False)


class TestGetWorksheetDataTool:
    """Tests for get_worksheet_data_tool."""

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    @patch('vba_mcp_pro.session_manager.win32com.client')
    async def test_get_data_from_excel(
        self,
        mock_win32com,
        mock_pythoncom,
        mock_platform,
        tmp_path
    ):
        """Test reading data from Excel."""
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

        # Mock range with 2x3 data
        mock_range = Mock()
        mock_range.Rows.Count = 2
        mock_range.Columns.Count = 3
        mock_range.Value = (
            ("A1", "B1", "C1"),
            ("A2", "B2", "C2")
        )
        mock_sheet.Range.return_value = mock_range
        mock_sheet.UsedRange = mock_range

        mock_workbook.Worksheets.return_value = mock_sheet

        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.Dispatch.return_value = mock_app

        # Get data
        result = await get_worksheet_data_tool(
            str(test_file),
            "Sheet1",
            range="A1:C2"
        )

        # Assertions
        assert "Data Retrieved from Excel" in result
        assert "Sheet1" in result
        assert "2 rows" in result
        assert "3 columns" in result

        # Check JSON data
        assert "A1" in result
        assert "C2" in result

        # Cleanup
        manager = OfficeSessionManager.get_instance()
        await manager.close_all_sessions(save=False)


class TestSetWorksheetDataTool:
    """Tests for set_worksheet_data_tool."""

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    @patch('vba_mcp_pro.session_manager.win32com.client')
    async def test_set_data_to_excel(
        self,
        mock_win32com,
        mock_pythoncom,
        mock_platform,
        tmp_path
    ):
        """Test writing data to Excel."""
        # Create test file
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Setup mocks
        mock_app = Mock()
        mock_app.Name = "Microsoft Excel"
        mock_app.Calculation = -4105  # xlCalculationAutomatic
        mock_workbook = Mock()
        mock_workbook.Name = "test.xlsm"

        # Mock worksheet
        mock_sheet = Mock()
        mock_sheet.Name = "Sheet1"

        # Mock cells and range
        mock_cell_start = Mock()
        mock_cell_start.Row = 1
        mock_cell_start.Column = 1
        mock_cell_end = Mock()
        mock_range_target = Mock()

        mock_sheet.Range.return_value = mock_cell_start
        mock_sheet.Cells.return_value = mock_cell_start
        mock_sheet.Range.return_value = mock_range_target

        mock_workbook.Worksheets.return_value = mock_sheet

        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.Dispatch.return_value = mock_app

        # Test data
        test_data = [
            [1, 2, 3],
            [4, 5, 6]
        ]

        # Set data
        result = await set_worksheet_data_tool(
            str(test_file),
            "Sheet1",
            test_data,
            start_cell="A1"
        )

        # Assertions
        assert "Data Written to Excel" in result
        assert "Sheet1" in result
        assert "2 rows" in result
        assert "3 columns" in result

        # Cleanup
        manager = OfficeSessionManager.get_instance()
        await manager.close_all_sessions(save=False)

    def test_validate_data_valid(self):
        """Test data validation with valid data."""
        data = [[1, 2, 3], [4, 5, 6]]
        _validate_data(data)  # Should not raise

    def test_validate_data_not_list(self):
        """Test data validation with non-list."""
        with pytest.raises(ValueError, match="must be a 2D array"):
            _validate_data("not a list")

    def test_validate_data_empty(self):
        """Test data validation with empty data."""
        with pytest.raises(ValueError, match="cannot be empty"):
            _validate_data([])

    def test_validate_data_not_rectangular(self):
        """Test data validation with non-rectangular data."""
        data = [[1, 2, 3], [4, 5]]  # Different row lengths
        with pytest.raises(ValueError, match="same length"):
            _validate_data(data)


class TestHelperFunctions:
    """Tests for helper functions."""

    def test_normalize_com_data_single_cell(self):
        """Test normalizing single cell data."""
        result = _normalize_com_data("A1", 1, 1)
        assert result == [["A1"]]

    def test_normalize_com_data_single_row(self):
        """Test normalizing single row data."""
        data = ("A1", "B1", "C1")
        result = _normalize_com_data(data, 1, 3)
        assert result == [["A1", "B1", "C1"]]

    def test_normalize_com_data_single_column(self):
        """Test normalizing single column data."""
        data = ("A1", "A2", "A3")
        result = _normalize_com_data(data, 3, 1)
        assert result == [["A1"], ["A2"], ["A3"]]

    def test_normalize_com_data_2d(self):
        """Test normalizing 2D data."""
        data = (
            ("A1", "B1"),
            ("A2", "B2")
        )
        result = _normalize_com_data(data, 2, 2)
        assert result == [["A1", "B1"], ["A2", "B2"]]

    def test_normalize_com_data_none(self):
        """Test normalizing None (empty range)."""
        result = _normalize_com_data(None, 2, 3)
        assert result == [[None, None, None], [None, None, None]]

    def test_get_cell_address(self):
        """Test converting row/col to Excel address."""
        assert _get_cell_address(1, 1) == "A1"
        assert _get_cell_address(10, 26) == "Z10"
        assert _get_cell_address(1, 27) == "AA1"
        assert _get_cell_address(100, 28) == "AB100"
        assert _get_cell_address(5, 702) == "ZZ5"


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
