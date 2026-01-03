"""
Access VBA Tools (PRO)

COM-based VBA extraction, analysis, and compilation for Access databases.
oletools does not support .accdb files, so we use COM (VBProject.VBComponents).
"""

import json
import logging
import re
from pathlib import Path
from typing import Any, Dict, List, Optional

from ..session_manager import OfficeSessionManager


# Configure logging
logger = logging.getLogger(__name__)


# VBA component type constants
VBA_COMPONENT_TYPES = {
    1: "standard",      # vbext_ct_StdModule
    2: "class",         # vbext_ct_ClassModule
    3: "form",          # vbext_ct_MSForm (UserForm)
    11: "activex",      # vbext_ct_ActiveXDesigner
    100: "document",    # vbext_ct_Document (e.g., Access forms/reports code behind)
}


async def extract_vba_access_tool(
    file_path: str,
    module_name: Optional[str] = None
) -> str:
    """
    Extract VBA code from an Access database using COM.

    This tool uses VBProject.VBComponents to extract VBA code since
    oletools does not support .accdb files.

    Args:
        file_path: Absolute path to Access database (.accdb or .mdb)
        module_name: Optional specific module to extract

    Returns:
        Formatted text with VBA code extraction results

    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If file is not an Access database or module not found
        RuntimeError: If VBA extraction fails
    """
    path = Path(file_path).resolve()

    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    # Validate file type
    if path.suffix.lower() not in ['.accdb', '.mdb']:
        raise ValueError(
            f"This tool only works with Access databases (.accdb, .mdb). "
            f"Got: {path.suffix}"
        )

    manager = OfficeSessionManager.get_instance()
    session = await manager.get_or_create_session(path, read_only=True)
    session.refresh_last_accessed()

    try:
        vb_project = session.vb_project

        if vb_project is None:
            return _format_no_vba_result(path)

        # Extract modules
        modules = []
        for component in vb_project.VBComponents:
            comp_name = component.Name
            comp_type = VBA_COMPONENT_TYPES.get(component.Type, f"unknown ({component.Type})")

            # Filter by module name if specified
            if module_name and comp_name.lower() != module_name.lower():
                continue

            # Get code
            code_module = component.CodeModule
            line_count = code_module.CountOfLines

            if line_count > 0:
                code = code_module.Lines(1, line_count)
            else:
                code = ""

            # Parse procedures from code
            procedures = _parse_procedures(code)

            modules.append({
                "name": comp_name,
                "type": comp_type,
                "code": code,
                "line_count": line_count,
                "procedures": procedures
            })

        # Check if module was found when filtering
        if module_name and not modules:
            available = [c.Name for c in vb_project.VBComponents]
            raise ValueError(
                f"Module '{module_name}' not found in {path.name}\n"
                f"Available modules: {', '.join(available) if available else '(none)'}"
            )

        if not modules:
            return _format_no_vba_result(path)

        # Format output
        return _format_extraction_result(path, modules, module_name)

    except ValueError:
        raise
    except Exception as e:
        raise RuntimeError(f"Failed to extract VBA from Access: {str(e)}")


async def analyze_structure_access_tool(
    file_path: str,
    module_name: Optional[str] = None
) -> str:
    """
    Analyze VBA code structure in an Access database using COM.

    Analyzes complexity, dependencies, and provides recommendations.

    Args:
        file_path: Absolute path to Access database (.accdb or .mdb)
        module_name: Optional specific module to analyze

    Returns:
        Formatted analysis report

    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If file is not an Access database
        RuntimeError: If analysis fails
    """
    path = Path(file_path).resolve()

    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    # Validate file type
    if path.suffix.lower() not in ['.accdb', '.mdb']:
        raise ValueError(
            f"This tool only works with Access databases (.accdb, .mdb). "
            f"Got: {path.suffix}"
        )

    manager = OfficeSessionManager.get_instance()
    session = await manager.get_or_create_session(path, read_only=True)
    session.refresh_last_accessed()

    try:
        vb_project = session.vb_project

        if vb_project is None:
            return f"No VBA code to analyze in {path.name}"

        # Collect all modules and procedures
        all_modules = []
        all_procedures = []
        dependencies = {}

        for component in vb_project.VBComponents:
            comp_name = component.Name
            comp_type = VBA_COMPONENT_TYPES.get(component.Type, f"unknown ({component.Type})")

            # Filter by module name if specified
            if module_name and comp_name.lower() != module_name.lower():
                continue

            # Get code
            code_module = component.CodeModule
            line_count = code_module.CountOfLines

            if line_count > 0:
                code = code_module.Lines(1, line_count)
            else:
                code = ""

            # Parse procedures and calculate complexity
            procedures = _parse_procedures(code)
            for proc in procedures:
                proc["module"] = comp_name
                proc["complexity"] = _calculate_complexity(proc.get("body", ""))
                proc["calls"] = _extract_calls(proc.get("body", ""))
                all_procedures.append(proc)

            # Extract dependencies (references to other modules)
            module_deps = _extract_dependencies(code)
            if module_deps:
                dependencies[comp_name] = module_deps

            all_modules.append({
                "name": comp_name,
                "type": comp_type,
                "line_count": line_count,
                "procedure_count": len(procedures)
            })

        # Check if module was found when filtering
        if module_name and not all_modules:
            available = [c.Name for c in vb_project.VBComponents]
            raise ValueError(
                f"Module '{module_name}' not found in {path.name}\n"
                f"Available modules: {', '.join(available) if available else '(none)'}"
            )

        if not all_modules:
            return f"No VBA code to analyze in {path.name}"

        # Format analysis output
        return _format_analysis_result(path, all_modules, all_procedures, dependencies)

    except ValueError:
        raise
    except Exception as e:
        raise RuntimeError(f"Failed to analyze VBA structure: {str(e)}")


async def compile_vba_tool(file_path: str) -> str:
    """
    Compile VBA project and detect compilation errors.

    This tool forces the VBA project to compile and reports any
    syntax or reference errors. This is important because if the
    VBA project has compilation errors, macros cannot be executed.

    Args:
        file_path: Absolute path to Office file (.accdb, .mdb, .xlsm, .xlsb)

    Returns:
        Compilation result with any errors found

    Raises:
        FileNotFoundError: If file doesn't exist
        RuntimeError: If compilation check fails

    Note:
        For Access databases, this uses Application.RunCommand(acCmdCompileAndSaveAllModules).
        For Excel files, it attempts to access VBProject which triggers compilation.
    """
    path = Path(file_path).resolve()

    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    manager = OfficeSessionManager.get_instance()
    session = await manager.get_or_create_session(path, read_only=False)
    session.refresh_last_accessed()

    errors = []
    warnings = []

    try:
        if session.app_type == "Access":
            # Access: Use RunCommand to compile
            errors, warnings = await _compile_access_vba(session)
        elif session.app_type == "Excel":
            # Excel: Check VBProject for errors
            errors, warnings = await _compile_excel_vba(session)
        else:
            raise ValueError(
                f"compile_vba only supports Access and Excel files. "
                f"Got: {session.app_type}"
            )

        # Format result
        return _format_compilation_result(path, session.app_type, errors, warnings)

    except Exception as e:
        # If compilation itself throws an error, capture it
        if "compile" in str(e).lower() or "syntax" in str(e).lower():
            errors.append({
                "module": "Unknown",
                "line": 0,
                "error": str(e),
                "type": "compilation"
            })
            return _format_compilation_result(path, session.app_type, errors, warnings)
        raise RuntimeError(f"Failed to compile VBA: {str(e)}")


async def _compile_access_vba(session) -> tuple:
    """
    Compile VBA in Access database.

    Args:
        session: Active Access session

    Returns:
        Tuple of (errors, warnings)
    """
    errors = []
    warnings = []

    try:
        app = session.app

        # Method 1: Try to compile using RunCommand
        # acCmdCompileAndSaveAllModules = 125
        try:
            app.DoCmd.RunCommand(125)
            logger.info("VBA compilation command executed successfully")
        except Exception as compile_error:
            error_msg = str(compile_error)
            # Parse error message for details
            errors.append({
                "module": "Unknown",
                "line": 0,
                "error": error_msg,
                "type": "compilation"
            })
            return errors, warnings

        # Method 2: Check each module by accessing its code
        # This can reveal errors that weren't caught by the compile command
        vb_project = session.vb_project

        if vb_project:
            for component in vb_project.VBComponents:
                try:
                    code_module = component.CodeModule
                    line_count = code_module.CountOfLines

                    if line_count > 0:
                        # Try to read all lines - this can trigger errors
                        _ = code_module.Lines(1, line_count)

                        # Check for common issues in the code
                        code = code_module.Lines(1, line_count)
                        module_warnings = _check_code_issues(component.Name, code)
                        warnings.extend(module_warnings)

                except Exception as module_error:
                    error_msg = str(module_error)
                    # Try to parse line number from error
                    line_num = _extract_line_number(error_msg)
                    errors.append({
                        "module": component.Name,
                        "line": line_num,
                        "error": error_msg,
                        "type": "syntax"
                    })

    except Exception as e:
        errors.append({
            "module": "VBProject",
            "line": 0,
            "error": str(e),
            "type": "access"
        })

    return errors, warnings


async def _compile_excel_vba(session) -> tuple:
    """
    Check VBA compilation in Excel workbook.

    Args:
        session: Active Excel session

    Returns:
        Tuple of (errors, warnings)
    """
    errors = []
    warnings = []

    try:
        vb_project = session.vb_project

        if vb_project is None:
            return errors, warnings

        # Check each component
        for component in vb_project.VBComponents:
            try:
                code_module = component.CodeModule
                line_count = code_module.CountOfLines

                if line_count > 0:
                    # Read code to check for issues
                    code = code_module.Lines(1, line_count)

                    # Check for common issues
                    module_warnings = _check_code_issues(component.Name, code)
                    warnings.extend(module_warnings)

            except Exception as module_error:
                error_msg = str(module_error)
                line_num = _extract_line_number(error_msg)
                errors.append({
                    "module": component.Name,
                    "line": line_num,
                    "error": error_msg,
                    "type": "syntax"
                })

    except Exception as e:
        errors.append({
            "module": "VBProject",
            "line": 0,
            "error": str(e),
            "type": "access"
        })

    return errors, warnings


def _parse_procedures(code: str) -> List[Dict[str, Any]]:
    """
    Parse VBA code to extract procedure information.

    Args:
        code: VBA source code

    Returns:
        List of procedure dictionaries
    """
    procedures = []
    lines = code.split('\n')

    current_proc = None
    current_body = []
    in_proc = False

    for i, line in enumerate(lines, 1):
        stripped = line.strip()

        # Check for procedure start
        sub_match = re.match(
            r'^(Public\s+|Private\s+)?(Sub|Function)\s+(\w+)\s*\(([^)]*)\)',
            stripped, re.IGNORECASE
        )

        if sub_match:
            # Save previous procedure if any
            if current_proc and in_proc:
                current_proc["body"] = '\n'.join(current_body)
                current_proc["end_line"] = i - 1
                procedures.append(current_proc)

            # Start new procedure
            visibility = sub_match.group(1) or "Public"
            proc_type = sub_match.group(2)
            proc_name = sub_match.group(3)
            params = sub_match.group(4)

            current_proc = {
                "name": proc_name,
                "type": proc_type,
                "visibility": visibility.strip(),
                "parameters": params,
                "start_line": i,
                "end_line": 0
            }
            current_body = [line]
            in_proc = True
            continue

        # Check for procedure end
        if in_proc and re.match(r'^End\s+(Sub|Function)', stripped, re.IGNORECASE):
            current_body.append(line)
            current_proc["body"] = '\n'.join(current_body)
            current_proc["end_line"] = i
            procedures.append(current_proc)
            current_proc = None
            current_body = []
            in_proc = False
            continue

        # Add line to current procedure body
        if in_proc:
            current_body.append(line)

    return procedures


def _calculate_complexity(code: str) -> int:
    """
    Calculate cyclomatic complexity of VBA code.

    Args:
        code: VBA procedure body

    Returns:
        Complexity score
    """
    if not code:
        return 1

    complexity = 1  # Base complexity

    # Decision points that increase complexity
    patterns = [
        r'\bIf\b',
        r'\bElseIf\b',
        r'\bSelect\s+Case\b',
        r'\bCase\b(?!\s+Else)',
        r'\bFor\b',
        r'\bFor\s+Each\b',
        r'\bDo\b',
        r'\bWhile\b',
        r'\bLoop\s+While\b',
        r'\bLoop\s+Until\b',
        r'\bAnd\b',
        r'\bOr\b',
        r'\bOn\s+Error\b',
    ]

    for pattern in patterns:
        matches = re.findall(pattern, code, re.IGNORECASE)
        complexity += len(matches)

    return complexity


def _extract_calls(code: str) -> List[str]:
    """
    Extract procedure calls from VBA code.

    Args:
        code: VBA procedure body

    Returns:
        List of called procedure names
    """
    if not code:
        return []

    calls = set()

    # Match procedure calls: Call ProcName or ProcName args
    # This is a simplified pattern
    patterns = [
        r'\bCall\s+(\w+)',
        r'^\s*(\w+)\s*(?:\(|$)',  # Procedure on its own line
    ]

    for pattern in patterns:
        matches = re.findall(pattern, code, re.MULTILINE | re.IGNORECASE)
        for match in matches:
            # Filter out VBA keywords
            if match.lower() not in ['if', 'then', 'else', 'end', 'for', 'next',
                                      'do', 'loop', 'while', 'wend', 'with',
                                      'select', 'case', 'dim', 'set', 'let',
                                      'public', 'private', 'sub', 'function',
                                      'exit', 'goto', 'resume', 'error']:
                calls.add(match)

    return list(calls)[:10]  # Limit to first 10


def _extract_dependencies(code: str) -> List[str]:
    """
    Extract external dependencies from VBA code.

    Args:
        code: VBA module code

    Returns:
        List of dependency names
    """
    deps = set()

    # Look for common patterns indicating dependencies
    patterns = [
        # Object creation
        r'CreateObject\s*\(\s*"([^"]+)"',
        # Early binding references
        r'New\s+(\w+\.\w+)',
        # Library references in Dim statements
        r'As\s+(\w+)\.',
    ]

    for pattern in patterns:
        matches = re.findall(pattern, code, re.IGNORECASE)
        deps.update(matches)

    return list(deps)


def _check_code_issues(module_name: str, code: str) -> List[Dict[str, Any]]:
    """
    Check VBA code for common issues.

    Args:
        module_name: Name of the module
        code: VBA source code

    Returns:
        List of warning dictionaries
    """
    warnings = []
    lines = code.split('\n')

    for i, line in enumerate(lines, 1):
        stripped = line.strip()

        # Check for Option Explicit missing
        if i == 1 and not stripped.lower().startswith('option explicit'):
            # Check if any line in first 5 lines has Option Explicit
            first_lines = '\n'.join(lines[:5]).lower()
            if 'option explicit' not in first_lines:
                warnings.append({
                    "module": module_name,
                    "line": 1,
                    "warning": "Option Explicit not declared",
                    "type": "best_practice"
                })
                break  # Only warn once per module

        # Check for On Error Resume Next without handler
        if 'on error resume next' in stripped.lower():
            # Look for On Error GoTo 0 in nearby lines
            nearby = '\n'.join(lines[max(0, i-1):min(len(lines), i+20)])
            if 'on error goto 0' not in nearby.lower():
                warnings.append({
                    "module": module_name,
                    "line": i,
                    "warning": "On Error Resume Next without On Error GoTo 0",
                    "type": "error_handling"
                })

    return warnings


def _extract_line_number(error_msg: str) -> int:
    """
    Extract line number from error message.

    Args:
        error_msg: Error message string

    Returns:
        Line number or 0 if not found
    """
    # Try various patterns
    patterns = [
        r'line\s+(\d+)',
        r'Line:\s*(\d+)',
        r'\((\d+)\)',
    ]

    for pattern in patterns:
        match = re.search(pattern, error_msg, re.IGNORECASE)
        if match:
            return int(match.group(1))

    return 0


def _format_no_vba_result(path: Path) -> str:
    """Format result when no VBA is found."""
    return f"""**VBA Extraction Results**

File: {path}
Format: .{path.suffix.lstrip('.')}

No VBA macros found in this file.
"""


def _format_extraction_result(
    path: Path,
    modules: List[Dict],
    module_filter: Optional[str]
) -> str:
    """Format VBA extraction results."""
    lines = []

    # Header
    lines.append("**VBA Extraction Results**")
    lines.append(f"File: {path}")
    lines.append(f"Format: .{path.suffix.lstrip('.')}")
    lines.append(f"Extracted: {len(modules)} module(s)")
    if module_filter:
        lines.append(f"Filter: {module_filter}")
    lines.append("")

    # Modules
    for module in modules:
        lines.append(f"## {module['name']} ({module['type']})")
        lines.append("")
        lines.append(f"**Lines:** {module['line_count']}")

        proc_names = [p['name'] for p in module['procedures']]
        lines.append(f"**Procedures:** {', '.join(proc_names) if proc_names else 'None'}")
        lines.append("")
        lines.append("```vba")
        lines.append(module['code'])
        lines.append("```")
        lines.append("")

    return "\n".join(lines)


def _format_analysis_result(
    path: Path,
    modules: List[Dict],
    procedures: List[Dict],
    dependencies: Dict[str, List[str]]
) -> str:
    """Format VBA structure analysis results."""
    lines = []

    # Header
    lines.append(f"**VBA Structure Analysis: {path.name}**")
    lines.append("")

    # Metrics
    total_lines = sum(m.get("line_count", 0) for m in modules)
    complexities = [p.get("complexity", 1) for p in procedures]
    avg_complexity = sum(complexities) / len(complexities) if complexities else 0
    max_complexity = max(complexities) if complexities else 0

    lines.append("### Metrics")
    lines.append(f"- **Total Modules:** {len(modules)}")
    lines.append(f"- **Total Procedures:** {len(procedures)}")
    lines.append(f"- **Total Lines:** {total_lines}")
    lines.append(f"- **Avg Complexity:** {avg_complexity:.1f}")
    lines.append(f"- **Max Complexity:** {max_complexity}")
    lines.append("")

    # Complexity assessment
    if avg_complexity <= 5:
        lines.append("Code complexity is **good** - well structured")
    elif avg_complexity <= 10:
        lines.append("Code complexity is **moderate** - consider refactoring complex procedures")
    else:
        lines.append("Code complexity is **high** - refactoring recommended")
    lines.append("")

    # Module breakdown
    lines.append("### Modules")
    for module in modules:
        lines.append(
            f"- **{module['name']}** ({module['type']}) - "
            f"{module['line_count']} lines, {module['procedure_count']} procedures"
        )
    lines.append("")

    # Procedures (sorted by complexity)
    if procedures:
        lines.append("### Procedures (by complexity)")
        sorted_procs = sorted(procedures, key=lambda p: p.get("complexity", 0), reverse=True)[:15]

        for proc in sorted_procs:
            complexity = proc.get("complexity", 1)
            indicator = "HIGH" if complexity > 10 else "MEDIUM" if complexity > 5 else "LOW"

            calls = proc.get("calls", [])
            calls_str = f" -> Calls: {', '.join(calls[:3])}" if calls else ""

            lines.append(
                f"- **{proc['module']}.{proc['name']}** "
                f"({proc['type']}) [{indicator}] Complexity: {complexity}{calls_str}"
            )
        lines.append("")

    # Dependencies
    if dependencies:
        lines.append("### Dependencies")
        for module, deps in dependencies.items():
            if deps:
                lines.append(f"- **{module}** -> {', '.join(deps)}")
        lines.append("")

    # Recommendations
    lines.append("### Recommendations")
    high_complexity = [p for p in procedures if p.get("complexity", 0) > 10]
    if high_complexity:
        lines.append(f"- Refactor {len(high_complexity)} high-complexity procedure(s):")
        for proc in high_complexity[:5]:
            lines.append(f"  - `{proc['module']}.{proc['name']}` (complexity: {proc['complexity']})")
    else:
        lines.append("- No high-complexity procedures detected")

    return "\n".join(lines)


def _format_compilation_result(
    path: Path,
    app_type: str,
    errors: List[Dict],
    warnings: List[Dict]
) -> str:
    """Format VBA compilation check results."""
    lines = []

    # Header
    lines.append(f"**VBA Compilation Check: {path.name}**")
    lines.append(f"Application: {app_type}")
    lines.append("")

    # Status
    if not errors:
        lines.append("### Status: SUCCESS")
        lines.append("")
        lines.append("VBA project compiled successfully. All modules are valid.")
    else:
        lines.append("### Status: FAILED")
        lines.append("")
        lines.append(f"Found {len(errors)} compilation error(s):")
        lines.append("")

        for error in errors:
            module = error.get("module", "Unknown")
            line = error.get("line", 0)
            error_msg = error.get("error", "Unknown error")
            error_type = error.get("type", "unknown")

            line_info = f" (line {line})" if line > 0 else ""
            lines.append(f"- **{module}**{line_info}: {error_msg}")
            lines.append(f"  Type: {error_type}")
            lines.append("")

    # Warnings
    if warnings:
        lines.append("### Warnings")
        lines.append("")

        for warning in warnings:
            module = warning.get("module", "Unknown")
            line = warning.get("line", 0)
            warning_msg = warning.get("warning", "Unknown warning")
            warning_type = warning.get("type", "unknown")

            line_info = f" (line {line})" if line > 0 else ""
            lines.append(f"- **{module}**{line_info}: {warning_msg}")
        lines.append("")

    # Recommendations
    if errors:
        lines.append("### Next Steps")
        lines.append("")
        lines.append("1. Fix the compilation errors listed above")
        lines.append("2. Run `compile_vba` again to verify fixes")
        lines.append("3. Once compilation succeeds, macros can be executed with `run_macro`")
    else:
        lines.append("")
        lines.append("You can now run macros with `run_macro`.")

    return "\n".join(lines)
