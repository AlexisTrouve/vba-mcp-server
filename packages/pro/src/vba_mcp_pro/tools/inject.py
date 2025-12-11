"""
VBA Injection Tool (PRO)

Injects modified VBA code back into Office files.
"""

from pathlib import Path
from typing import Optional
import shutil
from datetime import datetime
import sys
import platform


async def inject_vba_tool(
    file_path: str,
    module_name: str,
    code: str,
    create_backup: bool = True
) -> str:
    """
    Inject VBA code into an Office file.

    Args:
        file_path: Absolute path to Office file
        module_name: Name of module to update/create
        code: VBA code to inject
        create_backup: Whether to create backup before modification

    Returns:
        Success message with details

    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If file format unsupported
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

    # Inject VBA code
    try:
        result = _inject_vba_windows(
            path=path,
            module_name=module_name,
            code=code,
            app_name=app_name,
            open_method=open_method
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
        f"Action: {result['action']}",
    ]

    if backup_path:
        result_lines.append(f"Backup: {backup_path.name}")

    return "\n".join(result_lines)


def _inject_vba_windows(
    path: Path,
    module_name: str,
    code: str,
    app_name: str,
    open_method: str
) -> dict:
    """
    Inject VBA code using Windows COM automation.

    Args:
        path: Path to Office file
        module_name: Module name
        code: VBA code to inject
        app_name: COM application name (e.g., "Excel.Application")
        open_method: Method to open file

    Returns:
        Dictionary with injection result details
    """
    import win32com.client
    import pythoncom

    # Initialize COM
    pythoncom.CoInitialize()

    app = None
    file_obj = None
    action = "updated"

    try:
        # Create application instance
        app = win32com.client.Dispatch(app_name)
        app.Visible = False
        app.DisplayAlerts = False

        # Open file
        abs_path = str(path.resolve())
        if "Excel" in app_name:
            file_obj = app.Workbooks.Open(abs_path)
            vb_project = file_obj.VBProject
        elif "Word" in app_name:
            file_obj = app.Documents.Open(abs_path)
            vb_project = file_obj.VBProject
        elif "Access" in app_name:
            app.OpenCurrentDatabase(abs_path)
            vb_project = app.VBE.ActiveVBProject
        else:
            raise ValueError(f"Unknown application type: {app_name}")

        # Access VBA project
        try:
            # Check if module exists
            vb_component = None
            for component in vb_project.VBComponents:
                if component.Name == module_name:
                    vb_component = component
                    break

            if vb_component:
                # Update existing module
                code_module = vb_component.CodeModule
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
                except:
                    # Fallback if constants not available
                    vb_component = vb_project.VBComponents.Add(1)
                    vb_component.Name = module_name
                    vb_component.CodeModule.AddFromString(code)
                    action = "created"

        except Exception as e:
            raise PermissionError(
                f"Cannot access VBA project. "
                f"Ensure 'Trust access to the VBA project object model' is enabled in Office.\n"
                f"Error: {str(e)}"
            ) from e

        # Save and close
        if "Excel" in app_name:
            file_obj.Save()
            file_obj.Close(SaveChanges=True)
        elif "Word" in app_name:
            file_obj.Save()
            file_obj.Close(SaveChanges=True)
        elif "Access" in app_name:
            app.CloseCurrentDatabase()

        return {"action": action, "module": module_name}

    finally:
        # Cleanup
        if app:
            try:
                app.Quit()
            except:
                pass
        pythoncom.CoUninitialize()


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
