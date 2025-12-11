"""
VBA MCP Server - Lite Edition

Open source MCP server for VBA extraction and analysis.
"""

__version__ = "0.1.0"

# Re-export core components
from vba_mcp_core import OfficeHandler, VBAParser
from vba_mcp_core.tools import extract_vba_tool, list_modules_tool, analyze_structure_tool

__all__ = [
    "OfficeHandler",
    "VBAParser",
    "extract_vba_tool",
    "list_modules_tool",
    "analyze_structure_tool",
    "__version__",
]
