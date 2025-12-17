"""Tests for VBA MCP Lite server."""
import pytest
from mcp.types import Tool, TextContent

from vba_mcp_lite.server import app, list_tools, call_tool


class TestMCPServer:
    """Test suite for MCP server functionality."""

    def test_server_initialized(self):
        """Test that server is properly initialized."""
        assert app is not None
        assert app.name == "vba-mcp-server"

    @pytest.mark.asyncio
    async def test_list_tools(self):
        """Test that list_tools returns correct tools."""
        tools = await list_tools()

        assert isinstance(tools, list)
        assert len(tools) == 3

        tool_names = [tool.name for tool in tools]
        assert "extract_vba" in tool_names
        assert "list_modules" in tool_names
        assert "analyze_structure" in tool_names

        # Check extract_vba tool
        extract_tool = next(t for t in tools if t.name == "extract_vba")
        assert "file_path" in extract_tool.inputSchema["properties"]
        assert "module_name" in extract_tool.inputSchema["properties"]
        assert "file_path" in extract_tool.inputSchema["required"]

    @pytest.mark.asyncio
    async def test_call_unknown_tool(self):
        """Test calling an unknown tool."""
        result = await call_tool("unknown_tool", {})

        assert isinstance(result, list)
        assert len(result) > 0
        assert isinstance(result[0], TextContent)
        assert "Unknown tool" in result[0].text or "Error" in result[0].text

    @pytest.mark.asyncio
    async def test_call_extract_vba_missing_file(self):
        """Test extract_vba with missing file."""
        result = await call_tool("extract_vba", {"file_path": "/does/not/exist.xlsm"})

        assert isinstance(result, list)
        assert isinstance(result[0], TextContent)
        assert "Error" in result[0].text
        assert "not found" in result[0].text.lower()

    @pytest.mark.asyncio
    @pytest.mark.integration
    async def test_call_extract_vba_success(self, sample_xlsm):
        """Test successful extract_vba call."""
        result = await call_tool("extract_vba", {"file_path": str(sample_xlsm)})

        assert isinstance(result, list)
        assert isinstance(result[0], TextContent)
        assert "VBA Extraction Results" in result[0].text
        assert "TestModule" in result[0].text

    @pytest.mark.asyncio
    @pytest.mark.integration
    async def test_call_list_modules_success(self, sample_xlsm):
        """Test successful list_modules call."""
        result = await call_tool("list_modules", {"file_path": str(sample_xlsm)})

        assert isinstance(result, list)
        assert isinstance(result[0], TextContent)
        assert "VBA Modules in" in result[0].text
        assert "TestModule" in result[0].text

    @pytest.mark.asyncio
    @pytest.mark.integration
    async def test_call_analyze_structure_success(self, sample_xlsm):
        """Test successful analyze_structure call."""
        result = await call_tool("analyze_structure", {"file_path": str(sample_xlsm)})

        assert isinstance(result, list)
        assert isinstance(result[0], TextContent)
        assert "VBA Structure Analysis" in result[0].text
        assert "Metrics" in result[0].text
        assert "Total Procedures" in result[0].text

    @pytest.mark.asyncio
    async def test_call_with_permission_error(self, tmp_path):
        """Test handling of permission errors."""
        # This test is platform-dependent and may not work on all systems
        pytest.skip("Permission error testing is platform-dependent")
