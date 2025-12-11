"""
Structure Analysis Tool

Analyzes VBA code structure, dependencies, and complexity.
"""

from pathlib import Path
from typing import Optional

from lib.office_handler import OfficeHandler
from lib.vba_parser import VBAParser


async def analyze_structure_tool(file_path: str, module_name: Optional[str] = None) -> str:
    """
    Analyze VBA code structure and dependencies.

    Args:
        file_path: Absolute path to Office file
        module_name: Optional specific module to analyze

    Returns:
        Formatted analysis report

    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If file format unsupported or module not found
    """
    # Validate file
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    # Extract and parse
    handler = OfficeHandler()
    vba_project = handler.extract_vba_project(path)

    if not vba_project or not vba_project.get("modules"):
        return f"ℹ️ No VBA code to analyze in {path.name}"

    # Filter by module if specified
    modules = vba_project["modules"]
    if module_name:
        modules = [m for m in modules if m["name"] == module_name]
        if not modules:
            raise ValueError(f"Module '{module_name}' not found")

    # Analyze each module
    parser = VBAParser()
    all_procedures = []
    dependencies = {}

    for module in modules:
        parsed = parser.parse_module(module)

        # Collect procedures
        for proc in parsed.get("procedures", []):
            proc["module"] = module["name"]
            all_procedures.append(proc)

        # Track dependencies
        deps = parsed.get("dependencies", [])
        if deps:
            dependencies[module["name"]] = deps

    # Calculate metrics
    total_procedures = len(all_procedures)
    total_lines = sum(m.get("line_count", 0) for m in modules)

    complexities = [p.get("complexity", 1) for p in all_procedures]
    avg_complexity = sum(complexities) / len(complexities) if complexities else 0
    max_complexity = max(complexities) if complexities else 0

    # Format output
    lines = []
    lines.append(f"📊 **VBA Structure Analysis: {path.name}**")
    lines.append("")

    # Metrics
    lines.append("### Metrics")
    lines.append(f"- **Total Modules:** {len(modules)}")
    lines.append(f"- **Total Procedures:** {total_procedures}")
    lines.append(f"- **Total Lines:** {total_lines}")
    lines.append(f"- **Avg Complexity:** {avg_complexity:.1f}")
    lines.append(f"- **Max Complexity:** {max_complexity}")
    lines.append("")

    # Complexity assessment
    if avg_complexity <= 5:
        lines.append("✅ Code complexity is **good** - well structured")
    elif avg_complexity <= 10:
        lines.append("⚠️ Code complexity is **moderate** - consider refactoring complex procedures")
    else:
        lines.append("❌ Code complexity is **high** - refactoring recommended")
    lines.append("")

    # Procedures
    if all_procedures:
        lines.append("### Procedures")
        for proc in sorted(all_procedures, key=lambda p: p.get("complexity", 0), reverse=True)[:10]:
            complexity = proc.get("complexity", 1)
            complexity_icon = "🔴" if complexity > 10 else "🟡" if complexity > 5 else "🟢"

            calls = proc.get("calls", [])
            calls_str = f" → Calls: {', '.join(calls[:3])}" if calls else ""

            lines.append(f"- **{proc['module']}.{proc['name']}** "
                        f"({proc['type']}) {complexity_icon} Complexity: {complexity}{calls_str}")
        lines.append("")

    # Dependencies
    if dependencies:
        lines.append("### Dependencies")
        for module, deps in dependencies.items():
            if deps:
                lines.append(f"- **{module}** → {', '.join(deps)}")
        lines.append("")

    # Recommendations
    lines.append("### Recommendations")
    high_complexity = [p for p in all_procedures if p.get("complexity", 0) > 10]
    if high_complexity:
        lines.append(f"- Refactor {len(high_complexity)} high-complexity procedure(s):")
        for proc in high_complexity[:5]:
            lines.append(f"  - `{proc['module']}.{proc['name']}` (complexity: {proc['complexity']})")
    else:
        lines.append("- ✅ No high-complexity procedures detected")

    return "\n".join(lines)
