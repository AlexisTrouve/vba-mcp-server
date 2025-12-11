#!/usr/bin/env python3
"""
VBA MCP Server - Pro Edition

Model Context Protocol server with advanced VBA manipulation features.

Version: 0.1.0 (Pro - Commercial)
Author: Alexis Trouve
License: Commercial (See LICENSE file)
"""

import asyncio
import sys
from pathlib import Path

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent

# Core tools (from lite)
from vba_mcp_core.tools import extract_vba_tool, list_modules_tool, analyze_structure_tool

# Pro tools
from .tools import inject_vba_tool, refactor_tool, backup_tool


# Initialize MCP server
app = Server("vba-mcp-server-pro")


@app.list_tools()
async def list_tools() -> list[Tool]:
    """List all available tools (lite + pro)."""
    return [
        # === LITE TOOLS ===
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
                        "description": "Optional: specific module to extract"
                    }
                },
                "required": ["file_path"]
            }
        ),
        Tool(
            name="list_modules",
            description="List all VBA modules in an Office file without extracting code.",
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
            description="Analyze VBA code structure, dependencies, and complexity metrics.",
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
        ),

        # === PRO TOOLS ===
        Tool(
            name="inject_vba",
            description=(
                "[PRO] Inject modified VBA code back into Office files. "
                "Creates automatic backup before modification."
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
                        "description": "Name of module to update/create"
                    },
                    "code": {
                        "type": "string",
                        "description": "VBA code to inject"
                    },
                    "create_backup": {
                        "type": "boolean",
                        "description": "Create backup before modification (default: true)"
                    }
                },
                "required": ["file_path", "module_name", "code"]
            }
        ),
        Tool(
            name="refactor_vba",
            description=(
                "[PRO] AI-powered refactoring suggestions for VBA code. "
                "Analyzes complexity, naming, structure, and suggests improvements."
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
                        "description": "Optional: specific module to analyze"
                    },
                    "refactor_type": {
                        "type": "string",
                        "enum": ["all", "complexity", "naming", "structure"],
                        "description": "Type of refactoring analysis (default: all)"
                    }
                },
                "required": ["file_path"]
            }
        ),
        Tool(
            name="backup_vba",
            description=(
                "[PRO] Manage backups of Office files. "
                "Create, list, restore, or delete backups."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {
                        "type": "string",
                        "description": "Absolute path to the Office file"
                    },
                    "action": {
                        "type": "string",
                        "enum": ["create", "list", "restore", "delete"],
                        "description": "Action to perform (default: create)"
                    },
                    "backup_id": {
                        "type": "string",
                        "description": "Backup ID for restore/delete actions"
                    }
                },
                "required": ["file_path"]
            }
        )
    ]


@app.call_tool()
async def call_tool(name: str, arguments: dict) -> list[TextContent]:
    """Handle tool calls from MCP clients."""
    try:
        # Lite tools
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
        # Pro tools
        elif name == "inject_vba":
            result = await inject_vba_tool(
                file_path=arguments["file_path"],
                module_name=arguments["module_name"],
                code=arguments["code"],
                create_backup=arguments.get("create_backup", True)
            )
        elif name == "refactor_vba":
            result = await refactor_tool(
                file_path=arguments["file_path"],
                module_name=arguments.get("module_name"),
                refactor_type=arguments.get("refactor_type", "all")
            )
        elif name == "backup_vba":
            result = await backup_tool(
                file_path=arguments["file_path"],
                action=arguments.get("action", "create"),
                backup_id=arguments.get("backup_id")
            )
        else:
            raise ValueError(f"Unknown tool: {name}")

        return [TextContent(type="text", text=result)]

    except FileNotFoundError as e:
        return [TextContent(type="text", text=f"Error: File not found - {str(e)}")]
    except PermissionError as e:
        return [TextContent(type="text", text=f"Error: Permission denied - {str(e)}")]
    except Exception as e:
        return [TextContent(type="text", text=f"Error: {type(e).__name__} - {str(e)}")]


async def main():
    """Main entry point for the MCP server."""
    async with stdio_server() as (read_stream, write_stream):
        await app.run(
            read_stream,
            write_stream,
            app.create_initialization_options()
        )


def run():
    """Entry point for console script."""
    asyncio.run(main())


if __name__ == "__main__":
    run()
