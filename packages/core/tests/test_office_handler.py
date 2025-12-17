"""Tests for OfficeHandler."""
import pytest
from pathlib import Path

from vba_mcp_core.lib.office_handler import OfficeHandler


class TestOfficeHandler:
    """Test suite for OfficeHandler class."""

    def test_supported_formats(self):
        """Test that supported formats are defined."""
        handler = OfficeHandler()
        assert '.xlsm' in handler.SUPPORTED_FORMATS
        assert '.xlsb' in handler.SUPPORTED_FORMATS
        assert '.docm' in handler.SUPPORTED_FORMATS
        assert '.accdb' in handler.SUPPORTED_FORMATS

    def test_extract_vba_project_file_not_found(self, non_existent_file):
        """Test extraction with non-existent file."""
        handler = OfficeHandler()
        with pytest.raises(FileNotFoundError):
            handler.extract_vba_project(non_existent_file)

    def test_extract_vba_project_unsupported_format(self, tmp_path):
        """Test extraction with unsupported file format."""
        # Create a dummy .txt file
        test_file = tmp_path / "test.txt"
        test_file.write_text("dummy content")

        handler = OfficeHandler()
        with pytest.raises(ValueError, match="Unsupported format"):
            handler.extract_vba_project(test_file)

    @pytest.mark.integration
    def test_extract_vba_project_success(self, sample_xlsm):
        """Test successful VBA extraction from sample file."""
        handler = OfficeHandler()
        result = handler.extract_vba_project(sample_xlsm)

        assert "modules" in result
        assert isinstance(result["modules"], list)
        assert len(result["modules"]) > 0

        # Check module structure
        module = result["modules"][0]
        assert "name" in module
        assert "type" in module
        assert "code" in module
        assert "line_count" in module

        # Verify our test module
        test_module = next((m for m in result["modules"] if m["name"] == "TestModule"), None)
        assert test_module is not None
        assert "HelloWorld" in test_module["code"]
        assert "SumArray" in test_module["code"]

    def test_parse_module_name(self):
        """Test module name parsing."""
        handler = OfficeHandler()

        # Test with path separator
        assert handler._parse_module_name("VBA/Module1") == "Module1"

        # Test without path separator
        assert handler._parse_module_name("Module1") == "Module1"

    def test_determine_module_type(self):
        """Test module type determination."""
        handler = OfficeHandler()

        # Test workbook
        assert handler._determine_module_type("ThisWorkbook", "VBA/ThisWorkbook") == "workbook"

        # Test worksheet
        assert handler._determine_module_type("Sheet1", "VBA/Sheet1") == "worksheet"

        # Test form
        assert handler._determine_module_type("UserForm1", "VBA/UserForm1") == "form"

        # Test class
        assert handler._determine_module_type("MyClass", "VBA/Class Modules/MyClass") == "class"

        # Test standard module
        assert handler._determine_module_type("Module1", "VBA/Module1") == "standard"
