"""
VBA Extraction Tool

Extracts VBA source code from Microsoft Office files.
"""

import json
from pathlib import Path
from typing import Optional

from ..lib.office_handler import OfficeHandler
from ..lib.vba_parser import VBAParser


async def extract_vba_tool(file_path: str, module_name: Optional[str] = None) -> str:
    """
    Extract VBA code from an Office file.

    Args:
        file_path: Absolute path to Office file
        module_name: Optional specific module to extract

    Returns:
        JSON string with extraction results

    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If file format unsupported or module not found
    """
    # Validate file exists
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    # Initialize handler
    handler = OfficeHandler()

    # Extract VBA project
    try:
        vba_project = handler.extract_vba_project(path)
    except Exception as e:
        raise ValueError(f"Failed to extract VBA: {str(e)}")

    # If no VBA found
    if not vba_project or not vba_project.get("modules"):
        return json.dumps({
            "status": "error",
            "error": "No VBA macros found in file",
            "code": "NO_VBA",
            "file_info": {
                "path": str(path),
                "format": path.suffix.lstrip('.'),
                "size_bytes": path.stat().st_size
            }
        }, indent=2)

    # Filter by module name if specified
    modules = vba_project["modules"]
    if module_name:
        modules = [m for m in modules if m["name"] == module_name]
        if not modules:
            raise ValueError(f"Module '{module_name}' not found in file")

    # Parse each module
    parser = VBAParser()
    parsed_modules = []

    for module in modules:
        parsed = parser.parse_module(module)
        parsed_modules.append(parsed)

    # Build response
    result = {
        "status": "success",
        "modules": parsed_modules,
        "file_info": {
            "path": str(path),
            "format": path.suffix.lstrip('.'),
            "size_bytes": path.stat().st_size,
            "total_modules": len(vba_project["modules"]),
            "extracted_modules": len(parsed_modules)
        }
    }

    # Format as readable text for Claude
    text_output = _format_extraction_output(result)

    return text_output


def _format_extraction_output(result: dict) -> str:
    """
    Format extraction result as readable text.

    Args:
        result: Extraction result dictionary

    Returns:
        Formatted text output
    """
    lines = []

    # Header
    file_info = result["file_info"]
    lines.append(f"**VBA Extraction Results**")
    lines.append(f"File: {file_info['path']}")
    lines.append(f"Format: .{file_info['format']}")
    lines.append(f"Extracted: {file_info['extracted_modules']} of {file_info['total_modules']} modules")
    lines.append("")

    # Modules
    for module in result["modules"]:
        lines.append(f"## {module['name']} ({module['type']})")
        lines.append("")
        lines.append(f"**Lines:** {module['line_count']}")
        lines.append(f"**Procedures:** {', '.join(p['name'] for p in module['procedures']) if module['procedures'] else 'None'}")
        lines.append("")
        lines.append("```vba")
        lines.append(module['code'])
        lines.append("```")
        lines.append("")

    return "\n".join(lines)
