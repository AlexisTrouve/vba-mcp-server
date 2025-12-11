"""
Tests for Office Handler
"""

import pytest
import sys
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from lib.office_handler import OfficeHandler


class TestOfficeHandler:
    """Test suite for Office Handler"""

    @pytest.fixture
    def handler(self):
        """Create an OfficeHandler instance"""
        return OfficeHandler()

    def test_supported_formats(self, handler):
        """Test that supported formats are correctly defined"""
        assert '.xlsm' in handler.SUPPORTED_FORMATS
        assert '.xlsb' in handler.SUPPORTED_FORMATS
        assert '.accdb' in handler.SUPPORTED_FORMATS
        assert '.docm' in handler.SUPPORTED_FORMATS
        assert '.pptm' in handler.SUPPORTED_FORMATS

    def test_parse_module_name_simple(self, handler):
        """Test parsing simple module names"""
        assert handler._parse_module_name("Module1") == "Module1"
        assert handler._parse_module_name("ThisWorkbook") == "ThisWorkbook"

    def test_parse_module_name_with_path(self, handler):
        """Test parsing module names from paths"""
        assert handler._parse_module_name("VBA/Module1") == "Module1"
        assert handler._parse_module_name("path/to/Module2") == "Module2"

    def test_determine_module_type_workbook(self, handler):
        """Test detection of workbook module"""
        module_type = handler._determine_module_type("ThisWorkbook", "VBA/ThisWorkbook")
        assert module_type == "workbook"

    def test_determine_module_type_worksheet(self, handler):
        """Test detection of worksheet modules"""
        module_type = handler._determine_module_type("Sheet1", "VBA/Sheet1")
        assert module_type == "worksheet"

        module_type = handler._determine_module_type("Sheet2", "VBA/Sheet2")
        assert module_type == "worksheet"

    def test_determine_module_type_userform(self, handler):
        """Test detection of UserForm modules"""
        module_type = handler._determine_module_type("UserForm1", "VBA/UserForm1")
        assert module_type == "form"

    def test_determine_module_type_class(self, handler):
        """Test detection of class modules"""
        module_type = handler._determine_module_type("MyClass", "VBA/class/MyClass")
        assert module_type == "class"

    def test_determine_module_type_standard(self, handler):
        """Test detection of standard modules"""
        module_type = handler._determine_module_type("Module1", "VBA/Module1")
        assert module_type == "standard"

    def test_extract_nonexistent_file(self, handler):
        """Test extraction from non-existent file raises error"""
        fake_path = Path("nonexistent_file.xlsm")

        with pytest.raises(FileNotFoundError):
            handler.extract_vba_project(fake_path)

    def test_extract_unsupported_format(self, handler):
        """Test extraction from unsupported format raises error"""
        # Create a temporary text file
        temp_file = Path("temp_test.txt")
        temp_file.write_text("test")

        try:
            with pytest.raises(ValueError, match="Unsupported format"):
                handler.extract_vba_project(temp_file)
        finally:
            if temp_file.exists():
                temp_file.unlink()

    @pytest.mark.skipif(
        not Path("examples/test_simple.xlsm").exists(),
        reason="Test Excel file not available"
    )
    def test_extract_real_excel_file(self, handler):
        """Test extraction from a real Excel file (if available)"""
        test_file = Path("examples/test_simple.xlsm")

        result = handler.extract_vba_project(test_file)

        assert "modules" in result
        assert isinstance(result["modules"], list)

        # If there are modules, check structure
        if result["modules"]:
            module = result["modules"][0]
            assert "name" in module
            assert "type" in module
            assert "code" in module
            assert "line_count" in module

    def test_module_type_case_insensitive(self, handler):
        """Test that module type detection is case-insensitive"""
        # Lowercase
        assert handler._determine_module_type("thisworkbook", "VBA/thisworkbook") == "workbook"
        assert handler._determine_module_type("sheet1", "VBA/sheet1") == "worksheet"

        # Uppercase
        assert handler._determine_module_type("THISWORKBOOK", "VBA/THISWORKBOOK") == "workbook"
        assert handler._determine_module_type("SHEET1", "VBA/SHEET1") == "worksheet"

        # Mixed case
        assert handler._determine_module_type("ThisWorkBook", "VBA/ThisWorkBook") == "workbook"
        assert handler._determine_module_type("ShEeT1", "VBA/ShEeT1") == "worksheet"
