"""
VBA Refactoring Tool (PRO)

AI-powered refactoring suggestions for VBA code.
"""

from pathlib import Path
from typing import Optional, List, Dict

from vba_mcp_core import OfficeHandler, VBAParser


async def refactor_tool(
    file_path: str,
    module_name: Optional[str] = None,
    refactor_type: str = "all"
) -> str:
    """
    Analyze VBA code and suggest refactoring improvements.

    Args:
        file_path: Absolute path to Office file
        module_name: Optional specific module to refactor
        refactor_type: Type of refactoring (all, complexity, naming, structure)

    Returns:
        Refactoring suggestions report

    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If file format unsupported
    """
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    # Extract and analyze
    handler = OfficeHandler()
    vba_project = handler.extract_vba_project(path)

    if not vba_project or not vba_project.get("modules"):
        return f"No VBA code to refactor in {path.name}"

    modules = vba_project["modules"]
    if module_name:
        modules = [m for m in modules if m["name"] == module_name]
        if not modules:
            raise ValueError(f"Module '{module_name}' not found")

    # Analyze and generate suggestions
    parser = VBAParser()
    all_suggestions: List[Dict] = []

    for module in modules:
        parsed = parser.parse_module(module)
        suggestions = _analyze_for_refactoring(parsed, refactor_type)
        all_suggestions.extend(suggestions)

    # Format output
    lines = [
        f"**VBA Refactoring Suggestions: {path.name}**",
        "",
        f"Analyzed: {len(modules)} module(s)",
        f"Suggestions found: {len(all_suggestions)}",
        "",
    ]

    if not all_suggestions:
        lines.append("No refactoring suggestions - code looks good!")
    else:
        # Group by severity
        high = [s for s in all_suggestions if s["severity"] == "high"]
        medium = [s for s in all_suggestions if s["severity"] == "medium"]
        low = [s for s in all_suggestions if s["severity"] == "low"]

        if high:
            lines.append("### High Priority")
            for s in high:
                lines.append(f"- **{s['module']}.{s['location']}**: {s['message']}")
            lines.append("")

        if medium:
            lines.append("### Medium Priority")
            for s in medium:
                lines.append(f"- **{s['module']}.{s['location']}**: {s['message']}")
            lines.append("")

        if low:
            lines.append("### Low Priority")
            for s in low[:5]:  # Limit low priority
                lines.append(f"- **{s['module']}.{s['location']}**: {s['message']}")
            if len(low) > 5:
                lines.append(f"  ... and {len(low) - 5} more")
            lines.append("")

    return "\n".join(lines)


def _analyze_for_refactoring(module: Dict, refactor_type: str) -> List[Dict]:
    """
    Analyze a module for refactoring opportunities.

    Args:
        module: Parsed module dictionary
        refactor_type: Type of analysis to perform

    Returns:
        List of suggestion dictionaries
    """
    suggestions = []
    module_name = module.get("name", "Unknown")

    for proc in module.get("procedures", []):
        # Complexity check
        if refactor_type in ("all", "complexity"):
            complexity = proc.get("complexity", 1)
            if complexity > 15:
                suggestions.append({
                    "module": module_name,
                    "location": proc["name"],
                    "type": "complexity",
                    "severity": "high",
                    "message": f"Very high complexity ({complexity}). Split into smaller functions."
                })
            elif complexity > 10:
                suggestions.append({
                    "module": module_name,
                    "location": proc["name"],
                    "type": "complexity",
                    "severity": "medium",
                    "message": f"High complexity ({complexity}). Consider refactoring."
                })

        # Naming check
        if refactor_type in ("all", "naming"):
            name = proc["name"]
            if len(name) < 3:
                suggestions.append({
                    "module": module_name,
                    "location": name,
                    "type": "naming",
                    "severity": "low",
                    "message": "Very short name. Use descriptive names."
                })
            if name[0].islower() and proc["type"] in ("Sub", "Function"):
                suggestions.append({
                    "module": module_name,
                    "location": name,
                    "type": "naming",
                    "severity": "low",
                    "message": "Procedure names should start with uppercase (PascalCase)."
                })

        # Length check
        if refactor_type in ("all", "structure"):
            line_count = proc.get("line_end", 0) - proc.get("line_start", 0)
            if line_count > 50:
                suggestions.append({
                    "module": module_name,
                    "location": proc["name"],
                    "type": "structure",
                    "severity": "medium",
                    "message": f"Long procedure ({line_count} lines). Consider splitting."
                })

    return suggestions
