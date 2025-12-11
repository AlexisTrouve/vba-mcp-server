"""
Tests for MCP Tools
"""

import pytest
import asyncio
import sys
import json
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from tools.extract import extract_vba_tool
from tools.list_modules import list_modules_tool
from tools.analyze import analyze_structure_tool


class TestExtractTool:
    """Test suite for extract_vba_tool"""

    @pytest.mark.asyncio
    async def test_extract_nonexistent_file(self):
        """Test extraction from non-existent file"""
        with pytest.raises(FileNotFoundError):
            await extract_vba_tool("nonexistent.xlsm")

    @pytest.mark.asyncio
    async def test_extract_unsupported_format(self):
        """Test extraction from unsupported format"""
        # Create a temporary text file
        temp_file = Path("temp_test.txt")
        temp_file.write_text("test")

        try:
            with pytest.raises(ValueError):
                await extract_vba_tool(str(temp_file))
        finally:
            if temp_file.exists():
                temp_file.unlink()

    @pytest.mark.skipif(
        not Path("examples/test_simple.xlsm").exists(),
        reason="Test Excel file not available"
    )
    @pytest.mark.asyncio
    async def test_extract_real_file(self):
        """Test extraction from a real Excel file"""
        test_file = "examples/test_simple.xlsm"

        result = await extract_vba_tool(test_file)

        # Should return a string (formatted output)
        assert isinstance(result, str)
        assert len(result) > 0

        # Should contain expected markers
        assert "VBA Extraction Results" in result or "No VBA macros found" in result

    @pytest.mark.skipif(
        not Path("examples/test_simple.xlsm").exists(),
        reason="Test Excel file not available"
    )
    @pytest.mark.asyncio
    async def test_extract_specific_module(self):
        """Test extraction of a specific module"""
        test_file = "examples/test_simple.xlsm"

        # Try to extract "Module1" (common default name)
        # This might fail if the module doesn't exist, which is expected
        try:
            result = await extract_vba_tool(test_file, module_name="Module1")
            assert isinstance(result, str)
        except ValueError as e:
            # Module not found is acceptable
            assert "not found" in str(e).lower()


class TestListModulesTool:
    """Test suite for list_modules_tool"""

    @pytest.mark.asyncio
    async def test_list_nonexistent_file(self):
        """Test listing modules from non-existent file"""
        with pytest.raises(FileNotFoundError):
            await list_modules_tool("nonexistent.xlsm")

    @pytest.mark.skipif(
        not Path("examples/test_simple.xlsm").exists(),
        reason="Test Excel file not available"
    )
    @pytest.mark.asyncio
    async def test_list_real_file(self):
        """Test listing modules from a real Excel file"""
        test_file = "examples/test_simple.xlsm"

        result = await list_modules_tool(test_file)

        # Should return a string
        assert isinstance(result, str)
        assert len(result) > 0

        # Should contain expected markers
        assert "VBA Modules" in result or "No VBA modules found" in result


class TestAnalyzeTool:
    """Test suite for analyze_structure_tool"""

    @pytest.mark.asyncio
    async def test_analyze_nonexistent_file(self):
        """Test analyzing non-existent file"""
        with pytest.raises(FileNotFoundError):
            await analyze_structure_tool("nonexistent.xlsm")

    @pytest.mark.skipif(
        not Path("examples/test_simple.xlsm").exists(),
        reason="Test Excel file not available"
    )
    @pytest.mark.asyncio
    async def test_analyze_real_file(self):
        """Test analyzing a real Excel file"""
        test_file = "examples/test_simple.xlsm"

        result = await analyze_structure_tool(test_file)

        # Should return a string
        assert isinstance(result, str)
        assert len(result) > 0

        # Should contain expected markers
        assert "VBA Structure Analysis" in result or "No VBA modules found" in result


# Integration test
class TestToolsIntegration:
    """Integration tests for all tools together"""

    @pytest.mark.skipif(
        not Path("examples/test_simple.xlsm").exists(),
        reason="Test Excel file not available"
    )
    @pytest.mark.asyncio
    async def test_full_workflow(self):
        """Test the complete workflow: list → extract → analyze"""
        test_file = "examples/test_simple.xlsm"

        # Step 1: List modules
        modules_result = await list_modules_tool(test_file)
        assert isinstance(modules_result, str)

        # Step 2: Extract all VBA
        extract_result = await extract_vba_tool(test_file)
        assert isinstance(extract_result, str)

        # Step 3: Analyze structure
        analyze_result = await analyze_structure_tool(test_file)
        assert isinstance(analyze_result, str)

        # All results should be non-empty
        assert len(modules_result) > 0
        assert len(extract_result) > 0
        assert len(analyze_result) > 0
