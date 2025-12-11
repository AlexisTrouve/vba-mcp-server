#!/usr/bin/env python3
"""
VBA MCP Server - Main Entry Point

Model Context Protocol server for extracting and analyzing VBA code
from Microsoft Office files.

Version: 1.0.0 (Lite)
Author: Alexis Trouve
License: MIT (See LICENSE file)
"""

import asyncio
import sys
from pathlib import Path

# Add src directory to path for imports
sys.path.insert(0, str(Path(__file__).parent))

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent

from tools.extract import extract_vba_tool
from tools.list_modules import list_modules_tool
from tools.analyze import analyze_structure_tool


# Initialize MCP server
app = Server("vba-mcp-server")


@app.list_tools()
async def list_tools() -> list[Tool]:
    """
    List all available tools in the MCP server.

    Returns:
        List of Tool objects describing available functionality
    """
    return [
        Tool(
            name="extract_vba",
            description=(
                "Extract VBA source code from Microsoft Office files (.xlsm, .xlsb, .accdb, .docm). "
                "Returns the complete VBA code with module information, procedures, and metadata."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {
                        "type": "string",
                        "description": "Absolute path to the Office file"
                    },
                    "module_name": {
                        "type": "string",
                        "description": "Optional: specific module to extract. If omitted, extracts all modules."
                    }
                },
                "required": ["file_path"]
            }
        ),
        Tool(
            name="list_modules",
            description=(
                "List all VBA modules in an Office file without extracting the code. "
                "Provides a quick overview of module names, types, and line counts."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {
                        "type": "string",
                        "description": "Absolute path to the Office file"
                    }
                },
                "required": ["file_path"]
            }
        ),
        Tool(
            name="analyze_structure",
            description=(
                "Analyze VBA code structure, dependencies, and complexity metrics. "
                "Returns information about procedures, function calls, dependencies between modules, "
                "and code quality metrics."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {
                        "type": "string",
                        "description": "Absolute path to the Office file"
                    },
                    "module_name": {
                        "type": "string",
                        "description": "Optional: analyze specific module only"
                    }
                },
                "required": ["file_path"]
            }
        )
    ]


@app.call_tool()
async def call_tool(name: str, arguments: dict) -> list[TextContent]:
    """
    Handle tool calls from MCP clients (Claude Code).

    Args:
        name: Tool name to execute
        arguments: Tool-specific arguments

    Returns:
        List of TextContent with tool results

    Raises:
        ValueError: If tool name is unknown
    """
    try:
        if name == "extract_vba":
            result = await extract_vba_tool(
                file_path=arguments["file_path"],
                module_name=arguments.get("module_name")
            )
        elif name == "list_modules":
            result = await list_modules_tool(
                file_path=arguments["file_path"]
            )
        elif name == "analyze_structure":
            result = await analyze_structure_tool(
                file_path=arguments["file_path"],
                module_name=arguments.get("module_name")
            )
        else:
            raise ValueError(f"Unknown tool: {name}")

        # Format successful response
        return [TextContent(
            type="text",
            text=result
        )]

    except FileNotFoundError as e:
        return [TextContent(
            type="text",
            text=f"❌ Error: File not found - {str(e)}"
        )]
    except PermissionError as e:
        return [TextContent(
            type="text",
            text=f"❌ Error: Permission denied - {str(e)}"
        )]
    except Exception as e:
        return [TextContent(
            type="text",
            text=f"❌ Error: {type(e).__name__} - {str(e)}"
        )]


async def main():
    """
    Main entry point for the MCP server.
    Starts the server with stdio transport.
    """
    # Run the server using stdio transport (for local Claude Code)
    async with stdio_server() as (read_stream, write_stream):
        await app.run(
            read_stream,
            write_stream,
            app.create_initialization_options()
        )


if __name__ == "__main__":
    # Run the async server
    asyncio.run(main())
