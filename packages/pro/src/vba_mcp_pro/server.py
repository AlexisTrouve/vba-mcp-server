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
import platform
import logging
from pathlib import Path

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent

# Configure logging
logger = logging.getLogger(__name__)


def _check_wsl_environment():
    """
    Detect if running in WSL and log a warning.

    WSL (Windows Subsystem for Linux) cannot properly run pywin32 COM automation
    because COM is Windows-specific. This will cause Excel.Application.Visible
    and other properties to fail.

    Returns:
        bool: True if WSL detected, False otherwise
    """
    if platform.system() != "Windows":
        try:
            with open("/proc/version", "r") as f:
                if "microsoft" in f.read().lower():
                    logger.warning(
                        "WSL DETECTED: VBA MCP Pro is running in WSL. "
                        "COM automation has limited functionality in WSL. "
                        "For full functionality, run on native Windows Python. "
                        "See documentation for WSL workarounds."
                    )
                    return True
        except FileNotFoundError:
            pass

    return False


# Core tools (from lite)
from vba_mcp_core.tools import extract_vba_tool, list_modules_tool, analyze_structure_tool

# Pro tools
from .tools import (
    inject_vba_tool,
    refactor_tool,
    backup_tool,
    validate_vba_code_tool,
    open_in_office_tool,
    run_macro_tool,
    get_worksheet_data_tool,
    set_worksheet_data_tool,
    close_office_file_tool,
    list_open_files_tool,
    list_macros_tool,
    list_tables_tool,
    insert_rows_tool,
    delete_rows_tool,
    insert_columns_tool,
    delete_columns_tool,
    create_table_tool
)
from .session_manager import OfficeSessionManager


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
        ),
        Tool(
            name="validate_vba_code",
            description="[PRO] Validate VBA code syntax without injecting it into a file.",
            inputSchema={
                "type": "object",
                "properties": {
                    "code": {
                        "type": "string",
                        "description": "VBA code to validate"
                    },
                    "file_type": {
                        "type": "string",
                        "enum": ["excel", "word"],
                        "description": "Target Office application (default: excel)"
                    }
                },
                "required": ["code"]
            }
        ),

        # === INTERACTIVE OFFICE AUTOMATION TOOLS ===
        Tool(
            name="open_in_office",
            description=(
                "[PRO] Open Office file interactively with visible UI. "
                "File remains open for further operations."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {
                        "type": "string",
                        "description": "Absolute path to Office file"
                    },
                    "read_only": {
                        "type": "boolean",
                        "description": "Open in read-only mode (default: false)"
                    }
                },
                "required": ["file_path"]
            }
        ),
        Tool(
            name="run_macro",
            description=(
                "[PRO] Execute a VBA macro in an Office file. "
                "Format: 'ModuleName.MacroName' or 'MacroName'."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {
                        "type": "string",
                        "description": "Absolute path to Office file"
                    },
                    "macro_name": {
                        "type": "string",
                        "description": "Macro name (e.g., 'Module1.Calculate' or 'Calculate')"
                    },
                    "arguments": {
                        "type": "array",
                        "items": {},
                        "description": "Arguments to pass to macro (optional)"
                    },
                    "enable_macros": {
                        "type": "boolean",
                        "description": "Enable macros by temporarily lowering AutomationSecurity (default: true)"
                    }
                },
                "required": ["file_path", "macro_name"]
            }
        ),
        Tool(
            name="get_worksheet_data",
            description=(
                "[PRO] Read data from Excel worksheet or Access table. "
                "Returns JSON array."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {
                        "type": "string",
                        "description": "Absolute path to Office file"
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Worksheet/table name"
                    },
                    "range": {
                        "type": "string",
                        "description": "Cell range (e.g., 'A1:D10') or null for entire sheet"
                    },
                    "include_formulas": {
                        "type": "boolean",
                        "description": "Return formulas instead of values (default: false)"
                    }
                },
                "required": ["file_path", "sheet_name"]
            }
        ),
        Tool(
            name="set_worksheet_data",
            description=(
                "[PRO] Write data to Excel worksheet. "
                "Data must be 2D array."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {
                        "type": "string",
                        "description": "Absolute path to Excel file"
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Worksheet name (created if doesn't exist)"
                    },
                    "data": {
                        "type": "array",
                        "items": {"type": "array"},
                        "description": "2D array of values [[row1], [row2], ...]"
                    },
                    "start_cell": {
                        "type": "string",
                        "description": "Top-left cell (default: 'A1')"
                    },
                    "clear_existing": {
                        "type": "boolean",
                        "description": "Clear sheet before writing (default: false)"
                    }
                },
                "required": ["file_path", "sheet_name", "data"]
            }
        ),
        Tool(
            name="close_office_file",
            description="[PRO] Close an open Office file session.",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {
                        "type": "string",
                        "description": "Absolute path to Office file"
                    },
                    "save_changes": {
                        "type": "boolean",
                        "description": "Save before closing (default: true)"
                    }
                },
                "required": ["file_path"]
            }
        ),
        Tool(
            name="list_open_files",
            description="[PRO] List all currently open Office file sessions.",
            inputSchema={
                "type": "object",
                "properties": {}
            }
        ),
        Tool(
            name="list_macros",
            description="[PRO] List all public macros (Subs and Functions) in an Office file.",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {
                        "type": "string",
                        "description": "Absolute path to Office file"
                    }
                },
                "required": ["file_path"]
            }
        ),

        # === EXCEL TABLES TOOLS ===
        Tool(
            name="list_tables",
            description="[PRO] List all Excel Tables (ListObjects) in a file or sheet.",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {
                        "type": "string",
                        "description": "Absolute path to Excel file"
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Optional sheet name (if None, lists tables from all sheets)"
                    }
                },
                "required": ["file_path"]
            }
        ),
        Tool(
            name="insert_rows",
            description="[PRO] Insert row(s) in worksheet or Excel table.",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {
                        "type": "string",
                        "description": "Absolute path to Excel file"
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Sheet name"
                    },
                    "position": {
                        "type": "integer",
                        "description": "Row number (1-based) or position in table"
                    },
                    "count": {
                        "type": "integer",
                        "description": "Number of rows to insert (default: 1)"
                    },
                    "table_name": {
                        "type": "string",
                        "description": "Optional table name to insert in table context"
                    }
                },
                "required": ["file_path", "sheet_name", "position"]
            }
        ),
        Tool(
            name="delete_rows",
            description="[PRO] Delete row(s) from worksheet or Excel table.",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {
                        "type": "string",
                        "description": "Absolute path to Excel file"
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Sheet name"
                    },
                    "start_row": {
                        "type": "integer",
                        "description": "First row to delete (1-based)"
                    },
                    "end_row": {
                        "type": "integer",
                        "description": "Last row to delete (inclusive, optional)"
                    },
                    "table_name": {
                        "type": "string",
                        "description": "Optional table name to delete from table"
                    }
                },
                "required": ["file_path", "sheet_name", "start_row"]
            }
        ),
        Tool(
            name="insert_columns",
            description="[PRO] Insert column(s) in worksheet or Excel table.",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {
                        "type": "string",
                        "description": "Absolute path to Excel file"
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Sheet name"
                    },
                    "position": {
                        "description": "Column number (1-based) or letter (A, B, etc.)"
                    },
                    "count": {
                        "type": "integer",
                        "description": "Number of columns to insert (default: 1)"
                    },
                    "table_name": {
                        "type": "string",
                        "description": "Optional table name to insert in table"
                    },
                    "header_name": {
                        "type": "string",
                        "description": "Optional header name for new column (tables only)"
                    }
                },
                "required": ["file_path", "sheet_name", "position"]
            }
        ),
        Tool(
            name="delete_columns",
            description="[PRO] Delete column(s) from worksheet or Excel table.",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {
                        "type": "string",
                        "description": "Absolute path to Excel file"
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Sheet name"
                    },
                    "column": {
                        "description": "Column number, letter, or list of column names (for tables)"
                    },
                    "table_name": {
                        "type": "string",
                        "description": "Optional table name to delete from table"
                    }
                },
                "required": ["file_path", "sheet_name", "column"]
            }
        ),
        Tool(
            name="create_table",
            description="[PRO] Convert a range to an Excel Table (ListObject).",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {
                        "type": "string",
                        "description": "Absolute path to Excel file"
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Sheet name"
                    },
                    "range": {
                        "type": "string",
                        "description": "Range to convert (e.g., 'A1:D10')"
                    },
                    "table_name": {
                        "type": "string",
                        "description": "Name for the new table"
                    },
                    "has_headers": {
                        "type": "boolean",
                        "description": "First row contains headers (default: true)"
                    },
                    "style": {
                        "type": "string",
                        "description": "Excel table style (default: TableStyleMedium2)"
                    }
                },
                "required": ["file_path", "sheet_name", "range", "table_name"]
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
        elif name == "validate_vba_code":
            result = await validate_vba_code_tool(
                code=arguments["code"],
                file_type=arguments.get("file_type", "excel")
            )
        # Interactive Office tools
        elif name == "open_in_office":
            result = await open_in_office_tool(
                file_path=arguments["file_path"],
                read_only=arguments.get("read_only", False)
            )
        elif name == "run_macro":
            result = await run_macro_tool(
                file_path=arguments["file_path"],
                macro_name=arguments["macro_name"],
                arguments=arguments.get("arguments"),
                enable_macros=arguments.get("enable_macros", True)
            )
        elif name == "get_worksheet_data":
            result = await get_worksheet_data_tool(
                file_path=arguments["file_path"],
                sheet_name=arguments["sheet_name"],
                range=arguments.get("range"),
                include_formulas=arguments.get("include_formulas", False)
            )
        elif name == "set_worksheet_data":
            result = await set_worksheet_data_tool(
                file_path=arguments["file_path"],
                sheet_name=arguments["sheet_name"],
                data=arguments["data"],
                start_cell=arguments.get("start_cell", "A1"),
                clear_existing=arguments.get("clear_existing", False)
            )
        elif name == "close_office_file":
            result = await close_office_file_tool(
                file_path=arguments["file_path"],
                save_changes=arguments.get("save_changes", True)
            )
        elif name == "list_open_files":
            result = await list_open_files_tool()
        elif name == "list_macros":
            result = await list_macros_tool(
                file_path=arguments["file_path"]
            )
        elif name == "list_tables":
            result = await list_tables_tool(
                file_path=arguments["file_path"],
                sheet_name=arguments.get("sheet_name")
            )
        elif name == "insert_rows":
            result = await insert_rows_tool(
                file_path=arguments["file_path"],
                sheet_name=arguments["sheet_name"],
                position=arguments["position"],
                count=arguments.get("count", 1),
                table_name=arguments.get("table_name")
            )
        elif name == "delete_rows":
            result = await delete_rows_tool(
                file_path=arguments["file_path"],
                sheet_name=arguments["sheet_name"],
                start_row=arguments["start_row"],
                end_row=arguments.get("end_row"),
                table_name=arguments.get("table_name")
            )
        elif name == "insert_columns":
            result = await insert_columns_tool(
                file_path=arguments["file_path"],
                sheet_name=arguments["sheet_name"],
                position=arguments["position"],
                count=arguments.get("count", 1),
                table_name=arguments.get("table_name"),
                header_name=arguments.get("header_name")
            )
        elif name == "delete_columns":
            result = await delete_columns_tool(
                file_path=arguments["file_path"],
                sheet_name=arguments["sheet_name"],
                column=arguments["column"],
                table_name=arguments.get("table_name")
            )
        elif name == "create_table":
            result = await create_table_tool(
                file_path=arguments["file_path"],
                sheet_name=arguments["sheet_name"],
                range=arguments["range"],
                table_name=arguments["table_name"],
                has_headers=arguments.get("has_headers", True),
                style=arguments.get("style", "TableStyleMedium2")
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
    # Check for WSL environment and log warning if detected
    _check_wsl_environment()

    # Initialize session manager
    manager = OfficeSessionManager.get_instance()
    manager.start_cleanup_task()

    try:
        async with stdio_server() as (read_stream, write_stream):
            await app.run(
                read_stream,
                write_stream,
                app.create_initialization_options()
            )
    finally:
        # Cleanup all sessions on shutdown
        await manager.stop_cleanup_task()
        await manager.close_all_sessions(save=True)


def run():
    """Entry point for console script."""
    asyncio.run(main())


if __name__ == "__main__":
    run()
