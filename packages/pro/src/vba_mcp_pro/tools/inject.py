"""
VBA Injection Tool (PRO)

Injects modified VBA code back into Office files.
"""

from pathlib import Path
from typing import Optional, Tuple
import shutil
from datetime import datetime
import sys
import platform
import logging

logger = logging.getLogger(__name__)


def _configure_excel_app(app, visible: bool = False, display_alerts: bool = False):
    """
    Configure Excel application properties with graceful error handling.

    In WSL or service environments, some properties may not be settable.
    This function tries to set them but continues if they fail.

    Args:
        app: Excel.Application COM object
        visible: Whether Excel should be visible
        display_alerts: Whether to show alerts

    Returns:
        dict: Status of each property setting
    """
    status = {
        'visible': False,
        'display_alerts': False,
        'screen_updating': False
    }

    # Try to set Visible
    try:
        app.Visible = visible
        status['visible'] = True
        logger.debug(f"Excel.Visible set to {visible}")
    except Exception as e:
        logger.warning(f"Could not set Excel.Visible={visible}: {e}. Continuing anyway.")

    # Try to set DisplayAlerts
    try:
        app.DisplayAlerts = display_alerts
        status['display_alerts'] = True
        logger.debug(f"Excel.DisplayAlerts set to {display_alerts}")
    except Exception as e:
        logger.warning(f"Could not set Excel.DisplayAlerts={display_alerts}: {e}")

    # Try to set ScreenUpdating (performance optimization)
    try:
        app.ScreenUpdating = False
        status['screen_updating'] = True
        logger.debug("Excel.ScreenUpdating set to False")
    except Exception as e:
        logger.warning(f"Could not set Excel.ScreenUpdating=False: {e}")

    return status


def _detect_non_ascii(code: str) -> Tuple[bool, str]:
    """
    Detect non-ASCII characters in VBA code.

    VBA only supports ASCII characters (ord <= 127).

    Args:
        code: VBA code to check

    Returns:
        (has_non_ascii, error_message) tuple
    """
    non_ascii_chars = []
    for i, char in enumerate(code):
        if ord(char) > 127:
            # Find line number for better error reporting
            line_num = code[:i].count('\n') + 1
            non_ascii_chars.append((char, i, line_num))

    if non_ascii_chars:
        unique_chars = set(c for c, _, _ in non_ascii_chars)

        message = (
            f"VBA only supports ASCII characters.\n\n"
            f"Found {len(non_ascii_chars)} non-ASCII character(s): {', '.join(repr(c) for c in unique_chars)}\n\n"
            f"Common replacements:\n"
            f"  ✓ → [OK] or (check)\n"
            f"  ✗ → [ERROR] or (x)\n"
            f"  → → ->\n"
            f"  ➤ → >>\n"
            f"  • → *\n"
            f"  — → -\n"
            f"  " " → \" \"\n"
            f"  ' ' → ' '\n"
            f"  … → ...\n\n"
            f"First occurrence at line {non_ascii_chars[0][2]}"
        )
        return True, message

    return False, ""


def _suggest_ascii_replacement(code: str) -> Tuple[str, str]:
    """
    Suggest ASCII replacements for common Unicode characters.

    Args:
        code: VBA code with potential Unicode characters

    Returns:
        (suggested_code, changes_description) tuple
    """
    replacements = {
        '✓': '[OK]',
        '✗': '[ERROR]',
        '→': '->',
        '➤': '>>',
        '•': '*',
        '—': '-',
        '–': '-',
        '"': '"',
        '"': '"',
        ''': "'",
        ''': "'",
        '…': '...',
        '×': 'x',
        '÷': '/',
        '≤': '<=',
        '≥': '>=',
        '≠': '<>',
    }

    suggested = code
    changes = []

    for unicode_char, ascii_replacement in replacements.items():
        if unicode_char in suggested:
            count = suggested.count(unicode_char)
            suggested = suggested.replace(unicode_char, ascii_replacement)
            changes.append(f"  {repr(unicode_char)} → {repr(ascii_replacement)} ({count} occurrence(s))")

    if changes:
        changes_description = "Suggested replacements:\n" + "\n".join(changes)
        return suggested, changes_description
    else:
        return code, "No common Unicode characters found to replace automatically."


def _normalize_vba_code(code: str, strip_access_defaults: bool = True) -> str:
    """
    Normalize VBA code for comparison.

    VBA editor may add/remove blank lines, normalize whitespace, etc.
    This function normalizes code to make comparison more reliable.

    Args:
        code: VBA code to normalize
        strip_access_defaults: If True, strip Access-specific default lines
                              like "Option Compare Database" that Access adds
                              automatically to new modules

    Returns:
        Normalized code string
    """
    lines = code.splitlines()
    normalized_lines = []

    # Access-specific lines that are automatically added
    access_defaults = [
        "Option Compare Database",
        "Option Compare Text",
        "Option Compare Binary",
    ]

    for line in lines:
        stripped_line = line.strip()

        # Skip Access default lines if requested
        if strip_access_defaults and stripped_line in access_defaults:
            continue

        # Keep the line but strip trailing whitespace
        # Don't strip leading whitespace as indentation matters in VBA
        normalized_lines.append(line.rstrip())

    # Remove leading and trailing blank lines
    while normalized_lines and not normalized_lines[0].strip():
        normalized_lines.pop(0)
    while normalized_lines and not normalized_lines[-1].strip():
        normalized_lines.pop()

    return '\n'.join(normalized_lines)


def _check_vba_syntax(code: str) -> Tuple[bool, Optional[str]]:
    """
    Check VBA code syntax for common errors using pattern matching.

    This performs basic syntax checks before attempting COM-based validation.

    Args:
        code: VBA code to check

    Returns:
        (success, error_message) tuple
    """
    lines = code.splitlines()

    # Track block nesting
    if_count = 0
    for_count = 0
    while_count = 0
    do_count = 0
    with_count = 0
    select_count = 0
    sub_count = 0
    function_count = 0

    for line_num, line in enumerate(lines, 1):
        stripped = line.strip()

        # Skip empty lines and comments
        if not stripped or stripped.startswith("'") or stripped.startswith("Rem "):
            continue

        # Remove inline comments for analysis
        if "'" in stripped:
            # Simple approach: split on first '
            stripped = stripped.split("'")[0].strip()

        # Check for block start/end
        # If/Then/End If
        if stripped.startswith("If ") and " Then" in stripped and not stripped.endswith(" _"):
            # Check if it's a single-line If or multi-line If
            # Single-line: If condition Then statement (something after Then on same line)
            # Multi-line: If condition Then (nothing or just comment after Then)
            after_then = stripped.split(" Then", 1)[1].strip()
            # Remove any inline comment from after_then
            if "'" in after_then:
                after_then = after_then.split("'")[0].strip()

            # If there's nothing after Then, or just a colon, it's multi-line
            if not after_then or after_then == ":":
                if_count += 1
            # Otherwise it's single-line If (doesn't need End If)
        elif stripped.startswith("ElseIf ") and " Then" in stripped:
            pass  # ElseIf doesn't change count
        elif stripped.startswith("Else") and not stripped.startswith("ElseIf"):
            pass  # Else doesn't change count
        elif stripped.startswith("End If") or stripped == "End If":
            if_count -= 1
            if if_count < 0:
                return False, f"Line {line_num}: 'End If' without matching 'If'"

        # For/Next
        elif stripped.startswith("For "):
            for_count += 1
        elif stripped.startswith("Next"):
            for_count -= 1
            if for_count < 0:
                return False, f"Line {line_num}: 'Next' without matching 'For'"

        # While/Wend
        elif stripped.startswith("While "):
            while_count += 1
        elif stripped.startswith("Wend"):
            while_count -= 1
            if while_count < 0:
                return False, f"Line {line_num}: 'Wend' without matching 'While'"

        # Do/Loop
        elif stripped.startswith("Do") and (stripped == "Do" or stripped.startswith("Do While") or stripped.startswith("Do Until")):
            do_count += 1
        elif stripped.startswith("Loop"):
            do_count -= 1
            if do_count < 0:
                return False, f"Line {line_num}: 'Loop' without matching 'Do'"

        # With/End With
        elif stripped.startswith("With "):
            with_count += 1
        elif stripped.startswith("End With"):
            with_count -= 1
            if with_count < 0:
                return False, f"Line {line_num}: 'End With' without matching 'With'"

        # Select/End Select
        elif stripped.startswith("Select Case "):
            select_count += 1
        elif stripped.startswith("End Select"):
            select_count -= 1
            if select_count < 0:
                return False, f"Line {line_num}: 'End Select' without matching 'Select Case'"

        # Sub/End Sub
        elif stripped.startswith("Sub ") or stripped.startswith("Public Sub ") or stripped.startswith("Private Sub "):
            sub_count += 1
        elif stripped.startswith("End Sub"):
            sub_count -= 1
            if sub_count < 0:
                return False, f"Line {line_num}: 'End Sub' without matching 'Sub'"

        # Function/End Function
        elif stripped.startswith("Function ") or stripped.startswith("Public Function ") or stripped.startswith("Private Function "):
            function_count += 1
        elif stripped.startswith("End Function"):
            function_count -= 1
            if function_count < 0:
                return False, f"Line {line_num}: 'End Function' without matching 'Function'"

    # Check for unclosed blocks
    errors = []
    if if_count > 0:
        errors.append(f"{if_count} unclosed 'If' block(s) - missing 'End If'")
    if for_count > 0:
        errors.append(f"{for_count} unclosed 'For' loop(s) - missing 'Next'")
    if while_count > 0:
        errors.append(f"{while_count} unclosed 'While' loop(s) - missing 'Wend'")
    if do_count > 0:
        errors.append(f"{do_count} unclosed 'Do' loop(s) - missing 'Loop'")
    if with_count > 0:
        errors.append(f"{with_count} unclosed 'With' block(s) - missing 'End With'")
    if select_count > 0:
        errors.append(f"{select_count} unclosed 'Select Case' block(s) - missing 'End Select'")
    if sub_count > 0:
        errors.append(f"{sub_count} unclosed 'Sub' procedure(s) - missing 'End Sub'")
    if function_count > 0:
        errors.append(f"{function_count} unclosed 'Function' procedure(s) - missing 'End Function'")

    if errors:
        return False, "VBA Syntax Error:\n  " + "\n  ".join(errors)

    return True, None


def _compile_vba_module(vb_module) -> Tuple[bool, Optional[str]]:
    """
    Validate a VBA module by forcing VBA to parse the code.

    Uses ProcOfLine() which forces semantic checks on the code.
    This is more thorough than just reading lines.

    Args:
        vb_module: VBA component to validate

    Returns:
        (success, error_message) tuple

    Raises:
        Exception: For unexpected validation errors (no longer masked)
    """
    try:
        import pythoncom

        code_module = vb_module.CodeModule
        line_count = code_module.CountOfLines

        if line_count == 0:
            return True, None

        # Read all code first - forces basic parsing
        try:
            full_code = code_module.Lines(1, line_count)
        except pythoncom.com_error as e:
            return False, f"Failed to read code: {str(e)}"

        # PRE-CHECK: Basic syntax validation
        syntax_ok, syntax_error = _check_vba_syntax(full_code)
        if not syntax_ok:
            return False, syntax_error

        # Use ProcOfLine to force semantic checks on each line
        # This triggers VBA's internal parser more thoroughly
        for line_num in range(1, min(line_count + 1, 1000)):
            try:
                # ProcOfLine raises exception if there's a syntax error
                proc_name = code_module.ProcOfLine(line_num, 0)
            except pythoncom.com_error as e:
                error_msg = str(e)
                if "Compile error" in error_msg or "Syntax error" in error_msg:
                    return False, f"Syntax error at line {line_num}: {error_msg}"
                # Other errors are OK (e.g., line not in a procedure)

        # Verify basic module properties
        try:
            _ = vb_module.Name
            _ = code_module.CountOfDeclarationLines
        except pythoncom.com_error as e:
            return False, f"Module validation error: {str(e)}"

        return True, None

    except Exception as e:
        # No longer mask exceptions - propagate unexpected errors
        error_str = str(e)
        if "Compile error" in error_str or "Syntax error" in error_str:
            return False, f"Validation error: {error_str}"
        # For other unexpected errors, raise instead of masking
        raise


async def _verify_injection_via_session(
    session,  # OfficeSession
    module_name: str,
    expected_code: str
) -> Tuple[bool, Optional[str]]:
    """
    Verify VBA injection via existing session (for Access files).

    Since Access locks the file, we can't reopen it for verification.
    Instead, we verify by reading the code from the session's VB project.

    Args:
        session: OfficeSession instance
        module_name: Name of module to verify
        expected_code: Code that should be in the module

    Returns:
        (success, error_message) tuple
    """
    try:
        vb_project = session.vb_project

        # Search for the module
        for component in vb_project.VBComponents:
            if component.Name == module_name:
                code_module = component.CodeModule
                if code_module.CountOfLines > 0:
                    actual_code = code_module.Lines(1, code_module.CountOfLines)

                    # Normalize whitespace for comparison
                    expected_normalized = _normalize_vba_code(expected_code)
                    actual_normalized = _normalize_vba_code(actual_code)

                    logger.debug(f"Session verify - Expected: {len(expected_normalized)} chars, Actual: {len(actual_normalized)} chars")

                    if actual_normalized != expected_normalized:
                        logger.warning(f"Session verify code mismatch:")
                        logger.warning(f"Expected (first 200): {expected_normalized[:200]}")
                        logger.warning(f"Actual (first 200): {actual_normalized[:200]}")
                        return False, f"Code mismatch (expected {len(expected_normalized)} chars, got {len(actual_normalized)} chars)"
                else:
                    return False, "Module exists but is empty"
                break
        else:
            return False, f"Module '{module_name}' not found in session"

        return True, None

    except Exception as e:
        logger.error(f"Session verification exception: {str(e)}", exc_info=True)
        return False, f"Session verification failed: {str(e)}"


async def _verify_injection(
    file_path: Path,
    module_name: str,
    expected_code: str
) -> Tuple[bool, Optional[str]]:
    """
    Verify that VBA injection actually persisted by reopening the file.

    This detects cases where Save() succeeds but the module doesn't persist.
    Opens file in read-only mode to check module exists and code matches.

    Args:
        file_path: Path to Office file
        module_name: Name of module to verify
        expected_code: Code that should be in the module

    Returns:
        (success, error_message) tuple
    """
    import win32com.client
    import pythoncom

    # Don't call CoInitialize - we're already in an initialized COM context
    # from the session manager
    app = None
    file_obj = None

    try:
        # Determine application type from file extension
        file_ext = file_path.suffix.lower()
        if file_ext in ['.xlsm', '.xlsb', '.xls']:
            app = win32com.client.Dispatch("Excel.Application")
            _configure_excel_app(app, visible=False, display_alerts=False)
            file_obj = app.Workbooks.Open(str(file_path), ReadOnly=True)
        elif file_ext in ['.docm', '.doc']:
            app = win32com.client.Dispatch("Word.Application")
            # Word also has Visible property, handle gracefully
            try:
                app.Visible = False
                logger.debug("Word.Visible set to False")
            except Exception as e:
                logger.warning(f"Could not set Word.Visible=False: {e}. Continuing anyway.")
            try:
                app.DisplayAlerts = False
                logger.debug("Word.DisplayAlerts set to False")
            except Exception as e:
                logger.warning(f"Could not set Word.DisplayAlerts=False: {e}")
            file_obj = app.Documents.Open(str(file_path), ReadOnly=True)
            vb_project = file_obj.VBProject
        elif file_ext in ['.accdb', '.mdb']:
            app = win32com.client.Dispatch("Access.Application")
            try:
                app.Visible = False
                logger.debug("Access.Visible set to False")
            except Exception as e:
                logger.warning(f"Could not set Access.Visible=False: {e}. Continuing anyway.")
            # Access doesn't support ReadOnly in OpenCurrentDatabase
            # But we only read, so it's safe
            app.OpenCurrentDatabase(str(file_path))
            file_obj = app  # Access uses app itself as file container
            vb_project = app.VBE.ActiveVBProject
        else:
            return False, f"Unsupported file type for verification: {file_ext}"

        if file_ext not in ['.accdb', '.mdb']:
            vb_project = file_obj.VBProject

        # Search for the module
        for component in vb_project.VBComponents:
            if component.Name == module_name:
                code_module = component.CodeModule
                if code_module.CountOfLines > 0:
                    actual_code = code_module.Lines(1, code_module.CountOfLines)

                    # Normalize whitespace for comparison
                    # VBA may add extra blank lines or normalize whitespace
                    expected_normalized = _normalize_vba_code(expected_code)
                    actual_normalized = _normalize_vba_code(actual_code)

                    logger.debug(f"Expected code length: {len(expected_normalized)}, Actual: {len(actual_normalized)}")

                    if actual_normalized != expected_normalized:
                        # Log the difference for debugging
                        logger.warning(f"Code mismatch detected:")
                        logger.warning(f"Expected (first 200 chars): {expected_normalized[:200]}")
                        logger.warning(f"Actual (first 200 chars): {actual_normalized[:200]}")
                        return False, f"Code mismatch in saved file (expected {len(expected_normalized)} chars, got {len(actual_normalized)} chars)"
                else:
                    return False, "Module exists but is empty"
                break
        else:
            return False, f"Module '{module_name}' not found in saved file"

        return True, None

    except Exception as e:
        logger.error(f"Verification exception: {str(e)}", exc_info=True)
        return False, f"Verification failed: {str(e)}"

    finally:
        # Clean up COM objects
        # For Access, file_obj == app, so we only close database and quit
        if file_obj and file_obj != app:
            try:
                file_obj.Close(SaveChanges=False)
            except Exception as e:
                logger.warning(f"Error closing file during verification: {e}")
        elif file_obj == app:
            # Access: close current database before quitting
            try:
                app.CloseCurrentDatabase()
            except Exception as e:
                logger.warning(f"Error closing Access database during verification: {e}")

        if app:
            try:
                app.Quit()
            except Exception as e:
                logger.warning(f"Error quitting app during verification: {e}")

        # Don't call CoUninitialize - we didn't call CoInitialize


async def inject_vba_tool(
    file_path: str,
    module_name: str,
    code: str,
    create_backup: bool = True
) -> str:
    """
    Inject VBA code into an Office file.

    ENHANCED: Now validates code before and after injection with automatic rollback.

    Args:
        file_path: Absolute path to Office file
        module_name: Name of module to update/create
        code: VBA code to inject
        create_backup: Whether to create backup before modification

    Returns:
        Success message with details

    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If file format unsupported or code validation fails
        PermissionError: If file is locked
        ImportError: If pywin32 not available (Windows only)
    """
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    # Check platform
    if platform.system() != "Windows":
        raise RuntimeError(
            "VBA injection is only supported on Windows. "
            "Install this package on a Windows machine with Microsoft Office."
        )

    # PRE-VALIDATION: Check for non-ASCII characters
    has_non_ascii, ascii_error = _detect_non_ascii(code)
    if has_non_ascii:
        # Try to suggest replacements
        suggested_code, suggestions = _suggest_ascii_replacement(code)
        raise ValueError(
            f"Invalid VBA Code - Non-ASCII Characters Detected\n\n"
            f"{ascii_error}\n\n"
            f"{suggestions}\n\n"
            f"Please replace these characters with ASCII equivalents and try again."
        )

    # Import Windows-specific modules
    try:
        import win32com.client
        import pythoncom
    except ImportError:
        raise ImportError(
            "pywin32 is required for VBA injection. "
            "Install with: pip install vba-mcp-server-pro[windows]"
        )

    # Create backup if requested
    # Note: If file is already open in an existing session, backup may fail
    backup_path = None
    if create_backup:
        try:
            backup_path = _create_backup(path)
        except PermissionError as e:
            # File is likely open in another process or session
            logger.warning(f"Could not create backup (file may be open): {e}")
            # Continue without backup - the session will handle save/restore

    # Determine application type
    file_ext = path.suffix.lower()
    if file_ext in ['.xlsm', '.xlsb', '.xls']:
        app_name = "Excel.Application"
        open_method = "Workbooks.Open"
    elif file_ext in ['.docm', '.doc']:
        app_name = "Word.Application"
        open_method = "Documents.Open"
    elif file_ext in ['.accdb', '.mdb']:
        app_name = "Access.Application"
        open_method = "OpenCurrentDatabase"
    else:
        raise ValueError(f"Unsupported file type: {file_ext}")

    # Inject VBA code using session_manager
    try:
        from vba_mcp_pro.session_manager import OfficeSessionManager
        manager = OfficeSessionManager.get_instance()
        session = await manager.get_or_create_session(path)

        result = await _inject_vba_via_session(
            session=session,
            module_name=module_name,
            code=code,
            backup_path=backup_path
        )
    except Exception as e:
        # If injection fails and we created a backup, inform user
        if backup_path:
            raise RuntimeError(
                f"Injection failed: {str(e)}\n"
                f"Original file preserved. Backup at: {backup_path}"
            ) from e
        raise

    # Format result
    result_lines = [
        f"**VBA Injection Successful**",
        f"",
        f"File: {path.name}",
        f"Module: {module_name}",
        f"Code length: {len(code)} characters",
        f"Lines of code: {len(code.splitlines())}",
        f"Action: {result['action']}",
        f"Validation: {'Passed' if result.get('validated') else 'Skipped'}",
        f"Verified: {'Yes' if result.get('verified') else 'No'}",
    ]

    if backup_path:
        result_lines.append(f"Backup: {backup_path.name}")

    return "\n".join(result_lines)


async def _inject_vba_via_session(
    session,  # OfficeSession
    module_name: str,
    code: str,
    backup_path: Optional[Path] = None
) -> dict:
    """
    Inject VBA code via existing OfficeSession.

    Uses session.app, session.file_obj, session.vb_project to avoid concurrent access.
    No CoInitialize/CoUninitialize (managed by session).
    No app.Quit() (session stays open).

    Args:
        session: OfficeSession instance
        module_name: Module name
        code: VBA code to inject
        backup_path: Path to backup file (for potential rollback)

    Returns:
        Dictionary with injection result details

    Raises:
        ValueError: If code validation fails after injection
        PermissionError: If VBA project access is denied
    """
    import pythoncom

    action = "updated"
    old_code = None
    vb_component = None

    try:
        # Access VBA project via session
        vb_project = session.vb_project

        # Access VBA project
        try:
            # Check if module exists
            existing_component = None
            for component in vb_project.VBComponents:
                if component.Name == module_name:
                    existing_component = component
                    break

            if existing_component:
                # Update existing module - save old code for rollback
                vb_component = existing_component
                code_module = vb_component.CodeModule

                # Save old code for potential rollback
                if code_module.CountOfLines > 0:
                    old_code = code_module.Lines(1, code_module.CountOfLines)

                # Delete all existing code
                if code_module.CountOfLines > 0:
                    code_module.DeleteLines(1, code_module.CountOfLines)

                # Insert new code
                code_module.AddFromString(code)
                action = "updated"
            else:
                # Create new module
                # Determine module type (default to standard module)
                from win32com.client import constants
                try:
                    vb_component = vb_project.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
                    vb_component.Name = module_name
                    vb_component.CodeModule.AddFromString(code)
                    action = "created"
                except AttributeError:
                    # Fallback if constants not available
                    vb_component = vb_project.VBComponents.Add(1)
                    vb_component.Name = module_name
                    vb_component.CodeModule.AddFromString(code)
                    action = "created"
                except Exception as e:
                    logger.error(f"Failed to add component: {e}")
                    raise

            # POST-VALIDATION: Try to compile/validate the module
            compile_success, compile_error = _compile_vba_module(vb_component)

            if not compile_success:
                # ROLLBACK: Restore old code or delete module
                if old_code:
                    # Restore old code
                    code_module = vb_component.CodeModule
                    code_module.DeleteLines(1, code_module.CountOfLines)
                    code_module.AddFromString(old_code)
                else:
                    # Delete newly created module
                    vb_project.VBComponents.Remove(vb_component)

                raise ValueError(
                    f"VBA Code Validation Failed\n\n"
                    f"{compile_error}\n\n"
                    f"Code was NOT injected. File unchanged.\n"
                    f"{'Old code restored.' if old_code else 'Module not created.'}"
                )

        except ValueError:
            # Re-raise validation errors
            raise
        except pythoncom.com_error as e:
            error_msg = str(e).lower()
            if "permission" in error_msg or "access denied" in error_msg:
                raise PermissionError(
                    f"Cannot access VBA project. "
                    f"Ensure 'Trust access to the VBA project object model' is enabled in Office.\n"
                    f"Error: {str(e)}"
                ) from e
            # Other COM errors - propagate with original type
            raise RuntimeError(f"COM error during injection: {str(e)}") from e
        except Exception as e:
            # Non-COM errors - propagate
            raise

        # Save file (session stays open)
        if session.app_type == "Excel":
            session.file_obj.Save()
        elif session.app_type == "Word":
            session.file_obj.Save()
        elif session.app_type == "Access":
            # Access auto-saves, but we can force a save via DoCmd
            try:
                session.app.DoCmd.Save()
            except Exception as e:
                # DoCmd.Save may fail, Access still auto-saves
                logger.debug(f"DoCmd.Save failed (expected for some objects): {e}")

        # POST-SAVE VERIFICATION: Verify injection actually persisted
        if session.app_type == "Access":
            # For Access, verify via existing session (file is locked by session)
            success, error = await _verify_injection_via_session(session, module_name, code)
        else:
            # For Excel/Word, verify by reopening the file
            success, error = await _verify_injection(session.file_path, module_name, code)

        if not success:
            # Verification failed - restore backup if exists
            if backup_path and backup_path.exists():
                if session.app_type != "Access":
                    # Can only copy over file if it's not locked
                    import shutil
                    shutil.copy2(backup_path, session.file_path)
                    raise ValueError(
                        f"Injection verification failed: {error}\n"
                        f"File restored from backup: {backup_path}"
                    )
                else:
                    # For Access, we need to close session, restore, reopen
                    raise ValueError(
                        f"Injection verification failed: {error}\n"
                        f"Backup available at: {backup_path}\n"
                        f"Close Access manually and restore from backup if needed."
                    )
            else:
                raise ValueError(f"Injection verification failed: {error}")

        return {"action": action, "module": module_name, "validated": True, "verified": True}

    except Exception as e:
        # If we have a backup, restore it on error
        if backup_path and backup_path.exists():
            try:
                shutil.copy2(backup_path, session.file_path)
                logger.info(f"Restored file from backup after error: {backup_path}")
            except Exception as restore_error:
                logger.error(f"Failed to restore backup: {restore_error}")
        raise


def _create_backup(file_path: Path) -> Path:
    """
    Create a timestamped backup of the file.

    Args:
        file_path: Path to file to backup

    Returns:
        Path to backup file
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_name = f"{file_path.stem}_backup_{timestamp}{file_path.suffix}"
    backup_dir = file_path.parent / ".vba_backups"
    backup_dir.mkdir(exist_ok=True)
    backup_path = backup_dir / backup_name

    shutil.copy2(file_path, backup_path)

    return backup_path
