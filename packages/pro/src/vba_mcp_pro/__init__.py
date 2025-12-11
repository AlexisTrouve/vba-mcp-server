"""
VBA MCP Server - Pro Edition

Commercial MCP server with advanced VBA manipulation features.
"""

__version__ = "0.1.0"

# Re-export core and lite components
from vba_mcp_core import OfficeHandler, VBAParser
from vba_mcp_core.tools import extract_vba_tool, list_modules_tool, analyze_structure_tool

# Pro-only exports
from .tools import inject_vba_tool, refactor_tool, backup_tool

__all__ = [
    # Core
    "OfficeHandler",
    "VBAParser",
    # Lite tools
    "extract_vba_tool",
    "list_modules_tool",
    "analyze_structure_tool",
    # Pro tools
    "inject_vba_tool",
    "refactor_tool",
    "backup_tool",
    "__version__",
]
