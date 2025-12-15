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

    pythoncom.CoInitialize()
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
        else:
            return False, f"Unsupported file type for verification: {file_ext}"

        vb_project = file_obj.VBProject

        # Search for the module
        for component in vb_project.VBComponents:
            if component.Name == module_name:
                code_module = component.CodeModule
                if code_module.CountOfLines > 0:
                    actual_code = code_module.Lines(1, code_module.CountOfLines)

                    # Compare code (strip whitespace for comparison)
                    if actual_code.strip() != expected_code.strip():
                        return False, "Code mismatch in saved file"
                else:
                    return False, "Module exists but is empty"
                break
        else:
            return False, f"Module '{module_name}' not found in saved file"

        return True, None

    except Exception as e:
        return False, f"Verification failed: {str(e)}"

    finally:
        # Clean up COM objects
        if file_obj:
            try:
                file_obj.Close(SaveChanges=False)
            except Exception as e:
                logger.warning(f"Error closing file during verification: {e}")

        if app:
            try:
                app.Quit()
            except Exception as e:
                logger.warning(f"Error quitting app during verification: {e}")

        try:
            pythoncom.CoUninitialize()
        except Exception as e:
            logger.warning(f"Error uninitializing COM during verification: {e}")


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
    if create_backup:
        backup_path = _create_backup(path)
    else:
        backup_path = None

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
            # Access auto-saves
            pass

        # POST-SAVE VERIFICATION: Verify injection actually persisted
        success, error = await _verify_injection(session.file_path, module_name, code)
        if not success:
            # Verification failed - restore backup if exists
            if backup_path and backup_path.exists():
                import shutil
                shutil.copy2(backup_path, session.file_path)
                raise ValueError(
                    f"Injection verification failed: {error}\n"
                    f"File restored from backup: {backup_path}"
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
