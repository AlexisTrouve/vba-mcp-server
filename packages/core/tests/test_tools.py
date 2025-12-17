"""Tests for core tools (extract, list_modules, analyze)."""
import json
import pytest

from vba_mcp_core.tools.extract import extract_vba_tool
from vba_mcp_core.tools.list_modules import list_modules_tool


class TestExtractVBATool:
    """Test suite for extract_vba_tool."""

    @pytest.mark.asyncio
    async def test_extract_file_not_found(self, non_existent_file):
        """Test extraction with non-existent file."""
        with pytest.raises(FileNotFoundError):
            await extract_vba_tool(str(non_existent_file))

    @pytest.mark.asyncio
    @pytest.mark.integration
    async def test_extract_success(self, sample_xlsm):
        """Test successful VBA extraction."""
        result = await extract_vba_tool(str(sample_xlsm))

        # Result should be formatted text
        assert isinstance(result, str)
        assert "VBA Extraction Results" in result
        assert "TestModule" in result
        assert "HelloWorld" in result

    @pytest.mark.asyncio
    @pytest.mark.integration
    async def test_extract_specific_module(self, sample_xlsm):
        """Test extracting a specific module."""
        result = await extract_vba_tool(str(sample_xlsm), module_name="TestModule")

        assert isinstance(result, str)
        assert "TestModule" in result
        assert "HelloWorld" in result

    @pytest.mark.asyncio
    @pytest.mark.integration
    async def test_extract_nonexistent_module(self, sample_xlsm):
        """Test extracting a module that doesn't exist."""
        with pytest.raises(ValueError, match="Module .* not found"):
            await extract_vba_tool(str(sample_xlsm), module_name="NonExistentModule")


class TestListModulesTool:
    """Test suite for list_modules_tool."""

    @pytest.mark.asyncio
    async def test_list_file_not_found(self, non_existent_file):
        """Test listing modules with non-existent file."""
        with pytest.raises(FileNotFoundError):
            await list_modules_tool(str(non_existent_file))

    @pytest.mark.asyncio
    @pytest.mark.integration
    async def test_list_success(self, sample_xlsm):
        """Test successful module listing."""
        result = await list_modules_tool(str(sample_xlsm))

        assert isinstance(result, str)
        assert "VBA Modules in" in result
        assert "TestModule" in result
        assert "Total:" in result

    @pytest.mark.asyncio
    async def test_list_no_vba(self, tmp_path):
        """Test listing modules in file without VBA."""
        # This test would need a real .xlsm without VBA
        # For now, we'll skip it
        pytest.skip("Requires .xlsm file without VBA macros")
