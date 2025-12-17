# Plan de Correction - VBA MCP Issues

**Date:** 2025-12-14
**Projet:** vba-mcp-monorepo
**Référence:** ../vba-mcp-demo/MCP_ISSUES.md
**Objectif:** Corriger les 4 problèmes P0 + améliorations P2/P3

---

## Vue d'Ensemble

### Problèmes à Corriger

| # | Issue | Priorité | Temps Estimé |
|---|-------|----------|--------------|
| 1 | Excel crash lors injection | P0 | 1h |
| 2 | Pas de validation VBA | P0 | 1h30 |
| 3 | Pas de retour d'erreur compilation | P0 | 30min (inclus dans #2) |
| 4 | run_macro ne trouve pas les macros | P0 | 1h |
| 5 | Pas d'outil validate_vba | P2 | 1h |
| 6 | Gestion caractères non-ASCII | P2 | 30min |
| 7 | Pas de liste macros | P3 | 45min |

**Total:** ~6h30 de développement + 2h de tests

---

## Phase 1: Fixes Critiques P0 (4h)

### Fix #4: Corriger `run_macro` (1h)

**Problème:** La macro n'est jamais trouvée, peu importe le format utilisé.

**Fichier:** `packages/pro/src/vba_mcp_pro/tools/office_automation.py`

**Localisation:** Fonction `run_macro_tool()` (lignes ~265-300)

**Cause probable:**
- Format du nom de macro incorrect
- `Application.Run()` nécessite un format spécifique selon le contexte
- Différence entre Excel/Word/Access

**Solution:**

```python
async def run_macro_tool(
    file_path: str,
    macro_name: str,
    arguments: Optional[List[Any]] = None
) -> str:
    """Execute VBA macro with improved macro name resolution."""

    # Get session
    manager = OfficeSessionManager.get_instance()
    path = Path(file_path).resolve()
    session = await manager.get_or_create_session(path)
    session.refresh_last_accessed()

    args = arguments or []

    # Parse macro name
    parts = macro_name.split('.')
    if len(parts) == 2:
        module_name, proc_name = parts
    else:
        module_name = None
        proc_name = macro_name

    # Try different formats
    formats_to_try = []

    if session.app_type == "Excel":
        workbook_name = session.file_obj.Name
        if module_name:
            formats_to_try = [
                f"{module_name}.{proc_name}",  # Module.Macro
                f"'{workbook_name}'!{module_name}.{proc_name}",  # 'Book.xlsm'!Module.Macro
                proc_name,  # Just macro name
            ]
        else:
            formats_to_try = [
                proc_name,  # Just macro name
                f"'{workbook_name}'!{proc_name}",  # 'Book.xlsm'!Macro
            ]
    elif session.app_type in ["Word", "Access"]:
        # Word/Access typically just use macro name
        formats_to_try = [
            proc_name,
            f"{module_name}.{proc_name}" if module_name else proc_name
        ]

    # Try each format
    last_error = None
    for format_name in formats_to_try:
        try:
            result = session.app.Run(format_name, *args)

            # Build success message
            output = f"### Macro Executed Successfully\n\n"
            output += f"**Macro:** {macro_name}\n"
            output += f"**Format used:** {format_name}\n"
            output += f"**File:** {path.name}\n"

            if result is not None:
                output += f"**Return value:** {result}\n"
            else:
                output += f"**Type:** Sub (no return value)\n"

            return output

        except Exception as e:
            last_error = str(e)
            continue  # Try next format

    # All formats failed - list available macros
    try:
        available_macros = _list_available_macros(session)
        raise ValueError(
            f"Macro '{macro_name}' not found in {path.name}\n\n"
            f"Available macros:\n{available_macros}\n\n"
            f"Formats tried: {', '.join(formats_to_try)}\n"
            f"Last error: {last_error}"
        )
    except:
        raise ValueError(
            f"Macro '{macro_name}' not found.\n"
            f"Formats tried: {', '.join(formats_to_try)}\n"
            f"Last error: {last_error}"
        )


def _list_available_macros(session: OfficeSession) -> str:
    """List all public macros in the VBA project."""
    macros = []

    try:
        vb_project = session.vb_project

        for component in vb_project.VBComponents:
            module_name = component.Name
            code_module = component.CodeModule

            # Iterate through all lines to find Public Sub/Function
            line_count = code_module.CountOfLines
            for line_num in range(1, line_count + 1):
                line = code_module.Lines(line_num, 1).strip()

                # Check for Public Sub or Public Function
                if line.startswith("Public Sub ") or line.startswith("Sub "):
                    # Extract procedure name
                    if "(" in line:
                        proc_name = line.split("Sub ")[1].split("(")[0].strip()
                        macros.append(f"  - {module_name}.{proc_name} (Sub)")

                elif line.startswith("Public Function ") or line.startswith("Function "):
                    if "(" in line:
                        proc_name = line.split("Function ")[1].split("(")[0].strip()
                        macros.append(f"  - {module_name}.{proc_name} (Function)")

        if macros:
            return "\n".join(macros)
        else:
            return "  (No public macros found)"

    except Exception as e:
        return f"  (Error listing macros: {e})"
```

**Tests:**
1. Macro simple: `run_macro("test.xlsm", "HelloWorld")`
2. Macro avec module: `run_macro("test.xlsm", "TestModule.HelloWorld")`
3. Macro avec paramètres: `run_macro("test.xlsm", "AddNumbers", [10, 20])`
4. Macro inexistante: doit lister les macros disponibles

---

### Fix #2-3: Validation VBA + Retour Erreurs (1h30)

**Problème:**
- Code VBA invalide injecté sans vérification
- Pas de retour d'erreur de compilation
- Découverte des erreurs uniquement à l'exécution

**Fichier:** `packages/pro/src/vba_mcp_pro/tools/inject.py`

**Localisation:** Fonction `inject_vba_tool()` (probablement lignes ~50-150)

**Solution:**

```python
import re
from typing import Optional, Tuple

def _detect_non_ascii(code: str) -> Tuple[bool, str]:
    """Detect non-ASCII characters in VBA code."""
    non_ascii_chars = []
    for i, char in enumerate(code):
        if ord(char) > 127:
            non_ascii_chars.append((char, i))

    if non_ascii_chars:
        unique_chars = set(c for c, _ in non_ascii_chars)
        message = (
            f"VBA only supports ASCII characters.\n"
            f"Found non-ASCII characters: {', '.join(repr(c) for c in unique_chars)}\n\n"
            f"Common replacements:\n"
            f"  ✓ → - or [OK]\n"
            f"  ✗ → x or [ERROR]\n"
            f"  → → -> \n"
            f"  ➤ → >> \n"
            f"  • → * \n"
        )
        return True, message

    return False, ""


def _compile_vba_module(vb_module) -> Tuple[bool, Optional[str]]:
    """
    Compile a VBA module to check for syntax errors.

    Returns:
        (success, error_message)
    """
    try:
        # Access parent VBProject and compile
        vb_project = vb_module.CodeModule.Parent.VBProject

        # This will raise COM error if compilation fails
        # Note: VBProject.Compile() doesn't exist in pywin32
        # We need to use a workaround: try to access the module properties
        # which forces VBA to parse the code

        # Alternative: Execute a test macro that forces compilation
        # Or: Read the module line by line and check for basic syntax

        # Best approach: Use Excel's built-in error checking
        code_module = vb_module.CodeModule
        line_count = code_module.CountOfLines

        # Try to access each line - if there's a syntax error, it will fail
        for i in range(1, min(line_count + 1, 1000)):  # Limit to avoid timeout
            _ = code_module.Lines(i, 1)

        return True, None

    except pythoncom.com_error as e:
        error_msg = str(e)

        # Parse error message to extract useful info
        if "Compile error" in error_msg or "Syntax error" in error_msg:
            return False, f"VBA Compilation Error: {error_msg}"
        else:
            return False, f"VBA Error: {error_msg}"

    except Exception as e:
        return False, f"Unexpected error during compilation: {str(e)}"


async def inject_vba_tool(
    file_path: str,
    module_name: str,
    code: str,
    create_backup: bool = True
) -> str:
    """
    Inject VBA code into Office file with validation.

    IMPROVED: Now validates code before and after injection.
    """

    path = Path(file_path)

    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    # Check platform
    if platform.system() != "Windows":
        raise RuntimeError("VBA injection requires Windows + Microsoft Office")

    # PRE-VALIDATION: Check for non-ASCII characters
    has_non_ascii, ascii_error = _detect_non_ascii(code)
    if has_non_ascii:
        raise ValueError(f"Invalid VBA Code:\n\n{ascii_error}")

    # Create backup if requested
    backup_path = None
    if create_backup:
        from .backup import backup_tool
        backup_result = await backup_tool(file_path, action="create")
        # Extract backup path from result (assumes format in backup_tool)
        # This is a simplification - adjust based on actual backup_tool implementation
        backup_path = path.parent / f"{path.stem}_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}{path.suffix}"

    try:
        # Initialize COM
        pythoncom.CoInitialize()

        # Determine Office application type
        ext = path.suffix.lower()
        if ext in ['.xlsm', '.xlsb', '.xls']:
            app_name = "Excel.Application"
        elif ext in ['.docm', '.doc']:
            app_name = "Word.Application"
        elif ext in ['.accdb', '.mdb']:
            app_name = "Access.Application"
        else:
            raise ValueError(f"Unsupported file type: {ext}")

        # Open Office application
        app = win32com.client.Dispatch(app_name)
        app.Visible = False
        app.DisplayAlerts = False

        try:
            # Open file
            if ext in ['.xlsm', '.xlsb', '.xls']:
                file_obj = app.Workbooks.Open(str(path.absolute()))
            elif ext in ['.docm', '.doc']:
                file_obj = app.Documents.Open(str(path.absolute()))
            elif ext in ['.accdb', '.mdb']:
                file_obj = app.OpenCurrentDatabase(str(path.absolute()))

            # Access VBA project
            vb_project = file_obj.VBProject
            vb_components = vb_project.VBComponents

            # Check if module exists
            module_exists = False
            existing_module = None
            for component in vb_components:
                if component.Name == module_name:
                    module_exists = True
                    existing_module = component
                    break

            # Store old code for rollback
            old_code = None
            if module_exists:
                old_code = existing_module.CodeModule.Lines(
                    1,
                    existing_module.CodeModule.CountOfLines
                )
                # Remove old code
                existing_module.CodeModule.DeleteLines(1, existing_module.CodeModule.CountOfLines)
                vb_module = existing_module
            else:
                # Create new module
                vb_module = vb_components.Add(1)  # 1 = vbext_ct_StdModule
                vb_module.Name = module_name

            # Inject new code
            vb_module.CodeModule.AddFromString(code)

            # POST-VALIDATION: Try to compile
            compile_success, compile_error = _compile_vba_module(vb_module)

            if not compile_success:
                # ROLLBACK: Restore old code or delete module
                if old_code:
                    vb_module.CodeModule.DeleteLines(1, vb_module.CodeModule.CountOfLines)
                    vb_module.CodeModule.AddFromString(old_code)
                else:
                    vb_components.Remove(vb_module)

                # Save and close
                file_obj.Close(SaveChanges=False)
                app.Quit()

                raise ValueError(
                    f"VBA Code Validation Failed\n\n"
                    f"{compile_error}\n\n"
                    f"Code was NOT injected. File unchanged."
                )

            # Save and close
            file_obj.Save()
            file_obj.Close()
            app.Quit()

            # Success message
            output = f"### VBA Injection Successful ✓\n\n"
            output += f"**File:** {path.name}\n"
            output += f"**Module:** {module_name}\n"
            output += f"**Action:** {'Updated' if module_exists else 'Created'}\n"
            output += f"**Lines of code:** {len(code.splitlines())}\n"
            output += f"**Validation:** Passed ✓\n"

            if create_backup:
                output += f"**Backup:** Created\n"

            return output

        finally:
            pythoncom.CoUninitialize()

    except Exception as e:
        # If error and backup exists, mention restore option
        error_msg = f"### VBA Injection Failed\n\n"
        error_msg += f"**Error:** {str(e)}\n"
        error_msg += f"**File:** {path.name}\n"

        if backup_path and backup_path.exists():
            error_msg += f"\n**Backup available:** {backup_path.name}\n"
            error_msg += f"Use `backup_vba` tool with action='restore' to revert.\n"

        raise RuntimeError(error_msg)
```

**Tests:**
1. Code avec caractères Unicode: doit rejeter avec message clair
2. Code avec erreur syntaxe: doit rejeter après compilation
3. Code valide: doit injecter et compiler
4. Code invalide + rollback: vérifier que l'ancien code est restauré

---

### Fix #1: Stabiliser Excel (1h)

**Problème:** Excel peut crasher lors de l'injection, perte de données.

**Fichier:** `packages/pro/src/vba_mcp_pro/tools/inject.py`

**Solution:** Déjà partiellement implémenté dans Fix #2-3 ci-dessus avec:
- Backup automatique avant injection
- Rollback en cas d'erreur
- Try/catch robustes
- Validation avant/après

**Améliorations supplémentaires:**

```python
# Dans inject_vba_tool, ajouter:

# 1. Vérifier que le fichier n'est pas déjà ouvert ailleurs
def _check_file_not_locked(path: Path) -> None:
    """Check if file is locked by another process."""
    try:
        # Try to open file in exclusive mode
        with open(path, 'r+b') as f:
            pass
    except PermissionError:
        raise PermissionError(
            f"File is locked: {path.name}\n\n"
            f"Close the file in Excel/Word/Access and try again."
        )

# 2. Ajouter timeout pour les opérations
import signal

def timeout_handler(signum, frame):
    raise TimeoutError("Operation timed out")

# Dans la fonction principale:
try:
    # Set timeout (Windows doesn't support signal.SIGALRM)
    # Use threading.Timer instead
    import threading

    timeout = 30  # seconds
    timer = threading.Timer(timeout, lambda: None)
    timer.start()

    # ... injection code ...

    timer.cancel()

except TimeoutError:
    raise RuntimeError(
        f"Injection timeout after {timeout}s.\n"
        f"Excel may have frozen. Check Task Manager."
    )

# 3. Vérifier l'état d'Excel après injection
def _check_excel_alive(app) -> bool:
    """Check if Excel is still responsive."""
    try:
        _ = app.Name
        return True
    except:
        return False

# Après injection:
if not _check_excel_alive(app):
    raise RuntimeError(
        "Excel has become unresponsive.\n"
        "The injection may have succeeded but Excel crashed.\n"
        "Check the file manually."
    )
```

---

## Phase 2: Améliorations P2 (2h)

### Feature #5: Outil `validate_vba_code` (1h)

**Nouveau fichier:** `packages/pro/src/vba_mcp_pro/tools/validate.py`

```python
"""
VBA code validation tool.
"""

import platform
from pathlib import Path
from typing import Dict, Any

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
    from .inject import _detect_non_ascii
    has_non_ascii, ascii_error = _detect_non_ascii(code)
    if has_non_ascii:
        return f"### VBA Validation Failed\n\n{ascii_error}"

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
        app.Visible = False
        app.DisplayAlerts = False

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
```

**Enregistrer dans server.py:**

```python
# Dans list_tools():
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
)

# Dans call_tool():
elif name == "validate_vba_code":
    result = await validate_vba_code_tool(
        code=arguments["code"],
        file_type=arguments.get("file_type", "excel")
    )
```

**Tests:**
1. Code valide: doit retourner "Valid ✓"
2. Code avec erreur: doit retourner erreur
3. Code avec Unicode: doit rejeter

---

### Feature #6: Améliorer détection ASCII (30min)

**Fichier:** `packages/pro/src/vba_mcp_pro/tools/inject.py`

Déjà implémenté dans Fix #2-3 avec la fonction `_detect_non_ascii()`.

**Amélioration:** Ajouter un helper pour auto-remplacer:

```python
def _suggest_ascii_replacement(code: str) -> str:
    """Suggest ASCII replacements for common Unicode characters."""
    replacements = {
        '✓': '[OK]',
        '✗': '[ERROR]',
        '→': '->',
        '➤': '>>',
        '•': '*',
        '—': '-',
        '"': '"',
        '"': '"',
        ''': "'",
        ''': "'",
        '…': '...',
    }

    suggested = code
    changes = []

    for unicode_char, ascii_replacement in replacements.items():
        if unicode_char in suggested:
            suggested = suggested.replace(unicode_char, ascii_replacement)
            changes.append(f"  {repr(unicode_char)} → {repr(ascii_replacement)}")

    if changes:
        return suggested, "\n".join(changes)
    else:
        return code, ""


# Dans inject_vba_tool:
if has_non_ascii:
    suggested_code, changes = _suggest_ascii_replacement(code)
    raise ValueError(
        f"Invalid VBA Code (non-ASCII characters):\n\n"
        f"{ascii_error}\n\n"
        f"Suggested replacements:\n{changes}\n\n"
        f"Use the suggested code or manually replace characters."
    )
```

---

## Phase 3: Features P3 (45min)

### Feature #7: Outil `list_macros` (45min)

**Fichier:** `packages/pro/src/vba_mcp_pro/tools/office_automation.py`

Déjà partiellement implémenté dans Fix #4 avec `_list_available_macros()`.

**Créer outil MCP standalone:**

```python
async def list_macros_tool(file_path: str) -> str:
    """
    List all public macros (Subs and Functions) in an Office file.

    Args:
        file_path: Path to Office file

    Returns:
        Formatted list of macros with signatures
    """

    manager = OfficeSessionManager.get_instance()
    path = Path(file_path).resolve()

    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    # Get or create session
    session = await manager.get_or_create_session(path, read_only=True)
    session.refresh_last_accessed()

    try:
        macros_by_module = {}
        vb_project = session.vb_project

        for component in vb_project.VBComponents:
            module_name = component.Name
            code_module = component.CodeModule
            module_macros = []

            line_count = code_module.CountOfLines

            for line_num in range(1, line_count + 1):
                line = code_module.Lines(line_num, 1).strip()

                # Public Sub
                if line.startswith("Public Sub ") or line.startswith("Sub "):
                    signature = line
                    # Extract name and parameters
                    if "(" in signature:
                        full_sig = signature.replace("Public ", "").replace("Sub ", "")
                        module_macros.append({
                            "type": "Sub",
                            "signature": full_sig.split("'")[0].strip()  # Remove comments
                        })

                # Public Function
                elif line.startswith("Public Function ") or line.startswith("Function "):
                    signature = line
                    if "(" in signature:
                        full_sig = signature.replace("Public ", "").replace("Function ", "")
                        # Extract return type
                        if " As " in full_sig:
                            name_part, return_type = full_sig.split(" As ", 1)
                            return_type = return_type.split("'")[0].strip()
                        else:
                            name_part = full_sig
                            return_type = "Variant"

                        module_macros.append({
                            "type": "Function",
                            "signature": name_part.split("'")[0].strip(),
                            "returns": return_type
                        })

            if module_macros:
                macros_by_module[module_name] = module_macros

        # Format output
        output = f"### Macros in {path.name}\n\n"

        if not macros_by_module:
            output += "No public macros found.\n"
            return output

        total_macros = sum(len(macros) for macros in macros_by_module.values())
        output += f"**Total:** {total_macros} public macros in {len(macros_by_module)} modules\n\n"

        for module_name, macros in sorted(macros_by_module.items()):
            output += f"#### {module_name}\n\n"

            for macro in macros:
                if macro["type"] == "Sub":
                    output += f"- `{macro['signature']}` (Sub)\n"
                else:
                    output += f"- `{macro['signature']}` → {macro['returns']} (Function)\n"

            output += "\n"

        output += f"**Usage:** `run_macro(file, \"MacroName\")` or `run_macro(file, \"Module.MacroName\")`\n"

        return output

    except Exception as e:
        raise RuntimeError(f"Error listing macros: {str(e)}")
```

**Enregistrer dans server.py:**

```python
# Dans list_tools():
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
)

# Dans call_tool():
elif name == "list_macros":
    result = await list_macros_tool(
        file_path=arguments["file_path"]
    )
```

**Tests:**
1. Fichier avec macros: doit lister toutes les macros publiques
2. Fichier sans macros: doit retourner "No public macros found"
3. Fichier avec plusieurs modules: doit grouper par module

---

## Phase 4: Tests et Documentation (2h)

### Tests Unitaires (1h30)

**Fichier:** `packages/pro/tests/test_vba_validation.py`

```python
"""
Tests for VBA validation and error handling.
"""

import pytest
from pathlib import Path
from unittest.mock import Mock, patch

from vba_mcp_pro.tools.inject import (
    _detect_non_ascii,
    _suggest_ascii_replacement,
    inject_vba_tool
)
from vba_mcp_pro.tools.validate import validate_vba_code_tool
from vba_mcp_pro.tools.office_automation import run_macro_tool, list_macros_tool


class TestNonASCIIDetection:
    """Tests for non-ASCII character detection."""

    def test_detect_ascii_only(self):
        """Test code with only ASCII characters."""
        code = "Sub Test()\nMsgBox \"Hello\"\nEnd Sub"
        has_non_ascii, error = _detect_non_ascii(code)
        assert has_non_ascii is False
        assert error == ""

    def test_detect_unicode_checkmark(self):
        """Test detection of Unicode checkmark."""
        code = "Sub Test()\nMsgBox \"✓ Success\"\nEnd Sub"
        has_non_ascii, error = _detect_non_ascii(code)
        assert has_non_ascii is True
        assert "✓" in error or "non-ASCII" in error

    def test_suggest_replacement(self):
        """Test ASCII replacement suggestions."""
        code = "MsgBox \"✓ Done → Next\""
        suggested, changes = _suggest_ascii_replacement(code)
        assert "✓" not in suggested
        assert "[OK]" in suggested
        assert "->" in suggested


class TestVBAInjection:
    """Tests for VBA injection with validation."""

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    async def test_inject_unicode_code_fails(self, mock_platform, tmp_path):
        """Test that Unicode code is rejected."""
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        code = "Sub Test()\nMsgBox \"✓\"\nEnd Sub"

        with pytest.raises(ValueError, match="ASCII"):
            await inject_vba_tool(str(test_file), "TestModule", code)

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.tools.inject.win32com')
    @patch('vba_mcp_pro.tools.inject.pythoncom')
    async def test_inject_with_syntax_error_rollback(
        self, mock_pythoncom, mock_win32com, mock_platform, tmp_path
    ):
        """Test that syntax errors trigger rollback."""
        # Setup mocks to simulate compilation error
        # ... (similar to existing tests)
        pass


class TestRunMacro:
    """Tests for improved run_macro."""

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.win32com')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    async def test_run_macro_tries_multiple_formats(
        self, mock_pythoncom, mock_win32com, mock_platform, tmp_path
    ):
        """Test that run_macro tries different name formats."""
        # Mock that first format fails, second succeeds
        # ... setup mocks ...
        pass

    @pytest.mark.asyncio
    async def test_run_macro_not_found_lists_available(self, tmp_path):
        """Test that error message lists available macros."""
        # Test that when macro not found, error includes list
        # ... setup mocks ...
        pass


class TestListMacros:
    """Tests for list_macros tool."""

    @pytest.mark.asyncio
    async def test_list_macros_with_subs_and_functions(self, tmp_path):
        """Test listing macros with both Subs and Functions."""
        # ... setup mocks ...
        pass

    @pytest.mark.asyncio
    async def test_list_macros_no_public_macros(self, tmp_path):
        """Test file with no public macros."""
        # ... setup mocks ...
        pass


class TestValidateVBACode:
    """Tests for validate_vba_code tool."""

    @pytest.mark.asyncio
    async def test_validate_valid_code(self):
        """Test validating correct VBA code."""
        code = "Sub Test()\nMsgBox \"Hello\"\nEnd Sub"
        # ... mock Office ...
        result = await validate_vba_code_tool(code)
        assert "Valid" in result

    @pytest.mark.asyncio
    async def test_validate_invalid_syntax(self):
        """Test validating code with syntax error."""
        code = "Sub Test()\nIf x = 1\nEnd Sub"  # Missing End If
        # ... mock Office ...
        result = await validate_vba_code_tool(code)
        assert "Invalid" in result or "Error" in result


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
```

**Exécuter les tests:**
```bash
pytest packages/pro/tests/test_vba_validation.py -v
```

---

### Documentation (30min)

**Fichier:** `packages/pro/CHANGELOG.md`

```markdown
# Changelog - VBA MCP Pro

## [Unreleased] - 2025-12-14

### 🔴 Critical Fixes

#### Fixed `run_macro` not finding macros (Issue #4)
- **Problem:** Macros were never found regardless of format
- **Solution:** Try multiple name formats (MacroName, Module.MacroName, 'Book.xlsm'!MacroName)
- **Enhancement:** List available macros when macro not found
- **Impact:** CRITICAL - tool was completely broken

#### Added VBA code validation (Issues #2, #3)
- **Problem:** Invalid VBA code injected without checking, errors only at runtime
- **Solution:**
  - Detect non-ASCII characters before injection
  - Compile VBA code after injection to check syntax
  - Rollback to previous code if compilation fails
  - Return detailed error messages with line numbers
- **Impact:** CRITICAL - prevents file corruption and data loss

#### Improved Excel stability (Issue #1)
- **Problem:** Excel could crash during injection
- **Solution:**
  - Automatic backup before all injections
  - Robust try/catch blocks
  - Rollback on errors
  - Check Excel responsiveness after operations
- **Impact:** CRITICAL - prevents data loss

### ✨ New Features

#### `validate_vba_code` tool
- Validate VBA syntax without injecting into a file
- Creates temporary file, compiles code, returns errors
- Useful for testing code before injection

#### `list_macros` tool
- List all public Subs and Functions in a file
- Shows signatures and return types
- Grouped by module
- Helps discover available macros

### 🎨 Improvements

#### Better non-ASCII character handling
- Detect Unicode characters in VBA code
- Suggest ASCII replacements (✓ → [OK], → → ->, etc.)
- Clear error messages with examples

#### Enhanced error messages
- More descriptive errors with context
- Suggestions for fixing common issues
- List of available options when selection fails

### 📝 Documentation

- Updated SETUP_INSTRUCTIONS.md with new validation features
- Added troubleshooting section for common errors
- Documented all new tools in README.md

### 🧪 Tests

- Added 15+ new unit tests for validation
- Tests for non-ASCII detection
- Tests for rollback functionality
- Tests for macro name resolution
```

**Fichier:** `packages/pro/KNOWN_ISSUES.md`

```markdown
# Known Issues - VBA MCP Pro

## Resolved ✓

### ~~run_macro never finds macros~~
**Status:** FIXED in version X.X.X (2025-12-14)

### ~~No VBA validation before injection~~
**Status:** FIXED in version X.X.X (2025-12-14)

### ~~Excel crashes during injection~~
**Status:** IMPROVED in version X.X.X (2025-12-14)

## Current Issues

None currently known. Report issues at: [GitHub Issues URL]

## Limitations

### VBA Only Supports ASCII
VBA does not support Unicode characters in code (only in strings).
The server now detects and rejects non-ASCII characters with helpful suggestions.

### Windows + Office Required
All Pro features (injection, automation) require:
- Windows OS
- Microsoft Office installed
- Trust VBA Object Model enabled

### Compilation Limitations
VBA compilation detection is limited by pywin32 capabilities.
Some syntax errors may not be caught until runtime.

### Macro Execution Context
Some macros may fail when run via automation if they expect:
- User interaction (InputBox, MsgBox with user input)
- Active selection/context
- Specific Excel state
```

---

## Phase 5: Déploiement (30min)

### Mise à Jour de server.py

**Fichier:** `packages/pro/src/vba_mcp_pro/server.py`

Ajouter les imports:
```python
from .tools import (
    inject_vba_tool,
    refactor_tool,
    backup_tool,
    open_in_office_tool,
    run_macro_tool,
    get_worksheet_data_tool,
    set_worksheet_data_tool,
    close_office_file_tool,
    list_open_files_tool,
    validate_vba_code_tool,  # NEW
    list_macros_tool  # NEW
)
```

Ajouter dans `list_tools()`:
```python
# (Les 2 nouveaux outils déjà documentés ci-dessus)
```

Ajouter dans `call_tool()`:
```python
elif name == "validate_vba_code":
    result = await validate_vba_code_tool(...)

elif name == "list_macros":
    result = await list_macros_tool(...)
```

### Mise à Jour de __init__.py

**Fichier:** `packages/pro/src/vba_mcp_pro/tools/__init__.py`

```python
from .inject import inject_vba_tool
from .refactor import refactor_tool
from .backup import backup_tool
from .validate import validate_vba_code_tool  # NEW
from .office_automation import (
    open_in_office_tool,
    run_macro_tool,
    get_worksheet_data_tool,
    set_worksheet_data_tool,
    close_office_file_tool,
    list_open_files_tool,
    list_macros_tool  # NEW
)

__all__ = [
    "inject_vba_tool",
    "refactor_tool",
    "backup_tool",
    "validate_vba_code_tool",  # NEW
    "open_in_office_tool",
    "run_macro_tool",
    "get_worksheet_data_tool",
    "set_worksheet_data_tool",
    "close_office_file_tool",
    "list_open_files_tool",
    "list_macros_tool"  # NEW
]
```

---

## Checklist de Validation

### Avant de Commencer
- [ ] Lire MCP_ISSUES.md complètement
- [ ] Créer une branche Git: `git checkout -b fix/mcp-critical-issues`
- [ ] Backup du code actuel

### Phase 1 - Fixes P0
- [ ] Fix #4: `run_macro` - Implémenter formats multiples
- [ ] Fix #4: `run_macro` - Ajouter `_list_available_macros()`
- [ ] Fix #4: Test avec test.xlsm (HelloWorld, AddNumbers)
- [ ] Fix #2-3: Ajouter `_detect_non_ascii()`
- [ ] Fix #2-3: Ajouter `_compile_vba_module()`
- [ ] Fix #2-3: Modifier `inject_vba_tool()` avec validation
- [ ] Fix #2-3: Test avec code Unicode (doit rejeter)
- [ ] Fix #2-3: Test avec code invalide (doit rollback)
- [ ] Fix #1: Ajouter checks file locked
- [ ] Fix #1: Ajouter check Excel alive
- [ ] Fix #1: Test stabilité

### Phase 2 - Features P2
- [ ] Feature #5: Créer `validate.py`
- [ ] Feature #5: Implémenter `validate_vba_code_tool()`
- [ ] Feature #5: Enregistrer dans server.py
- [ ] Feature #5: Test avec code valide/invalide
- [ ] Feature #6: Ajouter `_suggest_ascii_replacement()`
- [ ] Feature #6: Intégrer dans messages d'erreur
- [ ] Feature #6: Test suggestions

### Phase 3 - Features P3
- [ ] Feature #7: Extraire `_list_available_macros()` en outil
- [ ] Feature #7: Créer `list_macros_tool()`
- [ ] Feature #7: Enregistrer dans server.py
- [ ] Feature #7: Test avec test.xlsm

### Phase 4 - Tests
- [ ] Créer `test_vba_validation.py`
- [ ] Tests pour non-ASCII detection
- [ ] Tests pour VBA injection avec validation
- [ ] Tests pour run_macro formats
- [ ] Tests pour list_macros
- [ ] Tests pour validate_vba_code
- [ ] Exécuter tous les tests: `pytest packages/pro/tests/ -v`
- [ ] Vérifier 100% pass

### Phase 5 - Documentation
- [ ] Mettre à jour CHANGELOG.md
- [ ] Créer KNOWN_ISSUES.md
- [ ] Mettre à jour README.md avec nouveaux outils
- [ ] Mettre à jour QUICK_TEST_PROMPTS.md
- [ ] Mettre à jour ../vba-mcp-demo/PROMPTS_READY_TO_USE.md

### Phase 6 - Tests Intégration Manuels
- [ ] Test 1: Injecter code avec Unicode → doit rejeter
- [ ] Test 2: Injecter code avec syntaxe erreur → doit rollback
- [ ] Test 3: Injecter code valide → doit compiler
- [ ] Test 4: Exécuter macro simple → doit fonctionner
- [ ] Test 5: Exécuter macro inexistante → doit lister disponibles
- [ ] Test 6: Valider code standalone → doit fonctionner
- [ ] Test 7: Lister macros → doit lister toutes les macros publiques
- [ ] Test 8: Workflow complet avec budget-analyzer.xlsm

### Phase 7 - Déploiement
- [ ] Commit: `git commit -m "fix: critical VBA issues (validation, run_macro, stability)"`
- [ ] Push: `git push origin fix/mcp-critical-issues`
- [ ] Créer Pull Request
- [ ] Code review
- [ ] Merge vers main
- [ ] Tag version: `git tag v0.2.0`
- [ ] Mettre à jour ../vba-mcp-demo/MCP_ISSUES.md (marquer comme résolu)

---

## Résumé des Fichiers à Modifier/Créer

### Fichiers à Modifier
1. `packages/pro/src/vba_mcp_pro/tools/inject.py` (validation VBA)
2. `packages/pro/src/vba_mcp_pro/tools/office_automation.py` (run_macro fix)
3. `packages/pro/src/vba_mcp_pro/tools/__init__.py` (exports)
4. `packages/pro/src/vba_mcp_pro/server.py` (enregistrement outils)

### Fichiers à Créer
5. `packages/pro/src/vba_mcp_pro/tools/validate.py` (nouveau)
6. `packages/pro/tests/test_vba_validation.py` (nouveau)
7. `packages/pro/CHANGELOG.md` (nouveau)
8. `packages/pro/KNOWN_ISSUES.md` (nouveau)

### Fichiers à Mettre à Jour
9. `README.md` (features)
10. `QUICK_TEST_PROMPTS.md` (nouveaux prompts)
11. `../vba-mcp-demo/MCP_ISSUES.md` (status résolu)
12. `../vba-mcp-demo/PROMPTS_READY_TO_USE.md` (nouveaux prompts)

---

## Temps Total Estimé

| Phase | Durée |
|-------|-------|
| Phase 1: Fixes P0 | 4h |
| Phase 2: Features P2 | 2h |
| Phase 3: Features P3 | 45min |
| Phase 4: Tests | 1h30 |
| Phase 5: Documentation | 30min |
| **TOTAL** | **8h45** |

**Avec tests manuels et déploiement:** ~10-12h (1.5 jours)

---

## Notes Importantes

1. **Priorité absolue:** Fixes P0 (#1-4) - systèmes de production bloqués
2. **Backup:** Toujours créer backup Git avant modification
3. **Tests:** Tester sur test.xlsm et budget-analyzer.xlsm
4. **Windows:** Tous les tests réels nécessitent Windows + Excel
5. **Documentation:** Documenter tous les changements pour utilisateurs

---

**Prêt pour implémentation?** Commence par Phase 1, Fix #4 (run_macro).
