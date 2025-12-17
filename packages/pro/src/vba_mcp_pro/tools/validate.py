"""
VBA code validation tool.
"""

import platform
import logging
from pathlib import Path
from typing import Dict, Any

logger = logging.getLogger(__name__)

if platform.system() == "Windows":
    import win32com.client
    import pythoncom


async def validate_vba_code_tool(
    code: str,
    file_type: str = "excel"
) -> str:
    """
    Validate VBA code without injecting it into a file.

    Creates a temporary Office file, injects code, compiles, then deletes.

    Args:
        code: VBA code to validate
        file_type: "excel", "word", or "access"

    Returns:
        Validation result with errors if any
    """

    if platform.system() != "Windows":
        raise RuntimeError("VBA validation requires Windows + Microsoft Office")

    # Check for non-ASCII
    from .inject import _detect_non_ascii, _check_vba_syntax
    has_non_ascii, ascii_error = _detect_non_ascii(code)
    if has_non_ascii:
        return f"### VBA Validation Failed\n\n{ascii_error}"

    # PRE-CHECK: Basic syntax validation (before creating temp file)
    syntax_ok, syntax_error = _check_vba_syntax(code)
    if not syntax_ok:
        return f"### VBA Validation Failed\n\n{syntax_error}"

    # Create temporary file
    import tempfile

    temp_dir = Path(tempfile.gettempdir())

    if file_type.lower() == "excel":
        app_name = "Excel.Application"
        temp_file = temp_dir / "vba_validation_temp.xlsm"
        file_format = 52  # xlOpenXMLWorkbookMacroEnabled
    elif file_type.lower() == "word":
        app_name = "Word.Application"
        temp_file = temp_dir / "vba_validation_temp.docm"
        file_format = 13  # wdFormatXMLDocumentMacroEnabled
    else:
        raise ValueError(f"Unsupported file type: {file_type}")

    try:
        pythoncom.CoInitialize()

        app = win32com.client.Dispatch(app_name)

        # Try to set Visible=False, but continue if it fails (WSL compatibility)
        try:
            app.Visible = False
            logger.debug("App.Visible set to False")
        except Exception as e:
            logger.warning(f"Could not set App.Visible=False: {e}. Continuing anyway.")

        try:
            app.DisplayAlerts = False
            logger.debug("App.DisplayAlerts set to False")
        except Exception as e:
            logger.warning(f"Could not set App.DisplayAlerts=False: {e}")

        try:
            # Create new file
            if file_type.lower() == "excel":
                file_obj = app.Workbooks.Add()
            elif file_type.lower() == "word":
                file_obj = app.Documents.Add()

            # Save as macro-enabled
            file_obj.SaveAs(str(temp_file.absolute()), FileFormat=file_format)

            # Inject code
            vb_module = file_obj.VBProject.VBComponents.Add(1)
            vb_module.Name = "ValidationTest"
            vb_module.CodeModule.AddFromString(code)

            # Try to compile
            from .inject import _compile_vba_module
            compile_success, compile_error = _compile_vba_module(vb_module)

            # Close without saving
            file_obj.Close(SaveChanges=False)
            app.Quit()

            # Build result
            if compile_success:
                output = f"### VBA Code Valid ✓\n\n"
                output += f"**Lines:** {len(code.splitlines())}\n"
                output += f"**Syntax:** Valid\n"
                output += f"**Compilation:** Successful\n"
                return output
            else:
                output = f"### VBA Code Invalid ✗\n\n"
                output += f"**Error:** {compile_error}\n"
                return output

        finally:
            # Cleanup temp file
            if temp_file.exists():
                try:
                    temp_file.unlink()
                except:
                    pass

            pythoncom.CoUninitialize()

    except Exception as e:
        return f"### VBA Validation Error\n\n{str(e)}"
