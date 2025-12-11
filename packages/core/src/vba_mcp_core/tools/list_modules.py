"""
List Modules Tool

Lists all VBA modules in an Office file without extracting code.
"""

from pathlib import Path

from ..lib.office_handler import OfficeHandler


async def list_modules_tool(file_path: str) -> str:
    """
    List all VBA modules in an Office file.

    Args:
        file_path: Absolute path to Office file

    Returns:
        Formatted text listing modules

    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If file format unsupported
    """
    # Validate file
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    # Extract VBA project (metadata only)
    handler = OfficeHandler()
    vba_project = handler.extract_vba_project(path)

    if not vba_project or not vba_project.get("modules"):
        return f"No VBA modules found in {path.name}"

    # Format output
    lines = []
    lines.append(f"**VBA Modules in {path.name}**")
    lines.append("")

    modules = vba_project["modules"]
    total_lines = 0

    for i, module in enumerate(modules, 1):
        line_count = module.get("line_count", 0)
        total_lines += line_count

        module_type = module.get("type", "unknown")
        lines.append(f"{i}. **{module['name']}** ({module_type})")
        lines.append(f"   - {line_count} lines")
        lines.append("")

    lines.append(f"**Total:** {len(modules)} modules, {total_lines} lines of code")

    return "\n".join(lines)
