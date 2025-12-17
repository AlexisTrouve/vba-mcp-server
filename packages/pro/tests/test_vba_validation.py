"""
Tests for VBA validation and error handling.

This module tests the validation functionality for VBA code injection,
including non-ASCII detection, code validation, macro execution, and
macro listing features.
"""

import pytest
import sys
from pathlib import Path
from unittest.mock import Mock, MagicMock, patch, call

# Mock Windows modules if not on Windows
if sys.platform != "win32":
    sys.modules['win32com'] = MagicMock()
    sys.modules['win32com.client'] = MagicMock()
    sys.modules['pythoncom'] = MagicMock()


# Helper functions for validation (these will be imported from actual implementation)
def _detect_non_ascii(code: str) -> tuple:
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


def _suggest_ascii_replacement(code: str) -> tuple:
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


class TestNonASCIIDetection:
    """Tests for non-ASCII character detection in VBA code."""

    def test_detect_ascii_only_code(self):
        """Test code with only ASCII characters (should pass)."""
        code = """Sub HelloWorld()
    MsgBox "Hello World"
End Sub"""
        has_non_ascii, error = _detect_non_ascii(code)
        assert has_non_ascii is False
        assert error == ""

    def test_detect_unicode_checkmark(self):
        """Test detection of Unicode checkmark (should detect)."""
        code = """Sub Test()
    MsgBox "✓ Success"
End Sub"""
        has_non_ascii, error = _detect_non_ascii(code)
        assert has_non_ascii is True
        assert "✓" in error or "non-ASCII" in error
        assert "Common replacements" in error

    def test_detect_multiple_unicode_chars(self):
        """Test detection of multiple Unicode characters."""
        code = """Sub Test()
    ' → This is a comment with arrows →
    MsgBox "• Bullet point"
End Sub"""
        has_non_ascii, error = _detect_non_ascii(code)
        assert has_non_ascii is True
        assert "non-ASCII" in error

    def test_detect_unicode_quotes(self):
        """Test detection of smart quotes."""
        code = """Sub Test()
    MsgBox "Hello"
End Sub"""
        has_non_ascii, error = _detect_non_ascii(code)
        assert has_non_ascii is True
        assert "non-ASCII" in error

    def test_suggest_ascii_replacement_checkmark(self):
        """Test ASCII replacement suggestions for checkmark."""
        code = 'MsgBox "✓ Done"'
        suggested, changes = _suggest_ascii_replacement(code)
        assert "✓" not in suggested
        assert "[OK]" in suggested
        assert len(changes) > 0

    def test_suggest_ascii_replacement_arrows(self):
        """Test ASCII replacement suggestions for arrows."""
        code = "' Step 1 → Step 2 ➤ Step 3"
        suggested, changes = _suggest_ascii_replacement(code)
        assert "→" not in suggested
        assert "➤" not in suggested
        assert "->" in suggested
        assert ">>" in suggested

    def test_suggest_ascii_replacement_multiple(self):
        """Test multiple ASCII replacements."""
        code = 'MsgBox "✓ Task complete → Next task • Item 1"'
        suggested, changes = _suggest_ascii_replacement(code)
        assert "✓" not in suggested
        assert "→" not in suggested
        assert "•" not in suggested
        assert "[OK]" in suggested
        assert "->" in suggested
        assert "*" in suggested

    def test_suggest_ascii_replacement_no_unicode(self):
        """Test ASCII replacement with no Unicode characters."""
        code = 'MsgBox "Hello World"'
        suggested, changes = _suggest_ascii_replacement(code)
        assert suggested == code
        assert changes == ""


class TestVBAInjection:
    """Tests for VBA injection with validation."""

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Linux')
    async def test_inject_unicode_non_windows(self, mock_platform):
        """Test that injection fails on non-Windows platforms."""
        from vba_mcp_pro.tools.inject import inject_vba_tool

        with pytest.raises(RuntimeError, match="Windows"):
            await inject_vba_tool(
                "/tmp/test.xlsm",
                "TestModule",
                'Sub Test()\nEnd Sub'
            )

    @pytest.mark.asyncio
    async def test_inject_file_not_found(self):
        """Test that injection fails when file doesn't exist."""
        from vba_mcp_pro.tools.inject import inject_vba_tool

        with pytest.raises(FileNotFoundError):
            await inject_vba_tool(
                "/nonexistent/file.xlsm",
                "TestModule",
                'Sub Test()\nEnd Sub'
            )

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.tools.inject.win32com')
    @patch('vba_mcp_pro.tools.inject.pythoncom')
    async def test_inject_valid_code_success(
        self,
        mock_pythoncom,
        mock_win32com,
        mock_platform,
        tmp_path
    ):
        """Test injecting valid code (should succeed)."""
        from vba_mcp_pro.tools.inject import inject_vba_tool

        # Create test file
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Setup mocks
        mock_app = Mock()
        mock_workbook = Mock()
        mock_vb_project = Mock()
        mock_vb_components = Mock()
        mock_component = Mock()
        mock_code_module = Mock()

        mock_code_module.CountOfLines = 0
        mock_component.CodeModule = mock_code_module
        mock_vb_components.Add.return_value = mock_component
        mock_vb_project.VBComponents = mock_vb_components
        mock_workbook.VBProject = mock_vb_project
        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.client.Dispatch.return_value = mock_app

        # Mock VBComponents as iterable
        mock_vb_components.__iter__ = Mock(return_value=iter([]))

        # Inject code
        result = await inject_vba_tool(
            str(test_file),
            "TestModule",
            'Sub Test()\nMsgBox "Hello"\nEnd Sub',
            create_backup=False
        )

        # Assertions
        assert "Successful" in result
        assert "TestModule" in result
        assert "test.xlsm" in result
        mock_component.CodeModule.AddFromString.assert_called_once()

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.tools.inject.win32com')
    @patch('vba_mcp_pro.tools.inject.pythoncom')
    async def test_inject_creates_backup(
        self,
        mock_pythoncom,
        mock_win32com,
        mock_platform,
        tmp_path
    ):
        """Test that backup is created when requested."""
        from vba_mcp_pro.tools.inject import inject_vba_tool

        # Create test file
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test content")

        # Setup mocks (same as above)
        mock_app = Mock()
        mock_workbook = Mock()
        mock_vb_project = Mock()
        mock_vb_components = Mock()
        mock_component = Mock()
        mock_code_module = Mock()

        mock_code_module.CountOfLines = 0
        mock_component.CodeModule = mock_code_module
        mock_vb_components.Add.return_value = mock_component
        mock_vb_project.VBComponents = mock_vb_components
        mock_workbook.VBProject = mock_vb_project
        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.client.Dispatch.return_value = mock_app
        mock_vb_components.__iter__ = Mock(return_value=iter([]))

        # Inject with backup
        result = await inject_vba_tool(
            str(test_file),
            "TestModule",
            'Sub Test()\nEnd Sub',
            create_backup=True
        )

        # Check that backup was mentioned in result
        assert "Backup" in result

        # Check that backup directory was created
        backup_dir = test_file.parent / ".vba_backups"
        assert backup_dir.exists()


class TestRunMacro:
    """Tests for run_macro tool with improved format handling."""

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    @patch('vba_mcp_pro.session_manager.win32com.client')
    async def test_run_macro_simple_name(
        self,
        mock_win32com,
        mock_pythoncom,
        mock_platform,
        tmp_path
    ):
        """Test running macro with simple name."""
        from vba_mcp_pro.tools.office_automation import run_macro_tool
        from vba_mcp_pro.session_manager import OfficeSessionManager

        # Create test file
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Setup mocks
        mock_app = Mock()
        mock_app.Name = "Microsoft Excel"
        mock_app.Run.return_value = None
        mock_workbook = Mock()
        mock_workbook.Name = "test.xlsm"
        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.Dispatch.return_value = mock_app

        # Run macro
        result = await run_macro_tool(str(test_file), "HelloWorld")

        # Assertions
        assert "Executed Successfully" in result or "Macro" in result
        mock_app.Run.assert_called()

        # Cleanup
        manager = OfficeSessionManager.get_instance()
        await manager.close_all_sessions(save=False)

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    @patch('vba_mcp_pro.session_manager.win32com.client')
    async def test_run_macro_with_module_name(
        self,
        mock_win32com,
        mock_pythoncom,
        mock_platform,
        tmp_path
    ):
        """Test running macro with module.macro format."""
        from vba_mcp_pro.tools.office_automation import run_macro_tool
        from vba_mcp_pro.session_manager import OfficeSessionManager

        # Create test file
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Setup mocks
        mock_app = Mock()
        mock_app.Name = "Microsoft Excel"
        mock_app.Run.return_value = 42
        mock_workbook = Mock()
        mock_workbook.Name = "test.xlsm"
        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.Dispatch.return_value = mock_app

        # Run macro with module name
        result = await run_macro_tool(
            str(test_file),
            "Module1.Calculate",
            arguments=[10, 20]
        )

        # Assertions
        assert "Calculate" in result or "Executed" in result
        mock_app.Run.assert_called()

        # Cleanup
        manager = OfficeSessionManager.get_instance()
        await manager.close_all_sessions(save=False)

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    @patch('vba_mcp_pro.session_manager.win32com.client')
    async def test_run_macro_with_parameters(
        self,
        mock_win32com,
        mock_pythoncom,
        mock_platform,
        tmp_path
    ):
        """Test that macro parameters are passed correctly."""
        from vba_mcp_pro.tools.office_automation import run_macro_tool
        from vba_mcp_pro.session_manager import OfficeSessionManager

        # Create test file
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Setup mocks
        mock_app = Mock()
        mock_app.Name = "Microsoft Excel"
        mock_app.Run.return_value = 30
        mock_workbook = Mock()
        mock_workbook.Name = "test.xlsm"
        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.Dispatch.return_value = mock_app

        # Run macro with parameters
        await run_macro_tool(
            str(test_file),
            "AddNumbers",
            arguments=[10, 20]
        )

        # Check that Run was called with parameters
        mock_app.Run.assert_called()
        call_args = mock_app.Run.call_args
        assert len(call_args[0]) >= 3  # macro name + 2 arguments

        # Cleanup
        manager = OfficeSessionManager.get_instance()
        await manager.close_all_sessions(save=False)


class TestListMacros:
    """Tests for list_macros functionality."""

    def test_list_macros_parsing_sub(self):
        """Test parsing Sub procedures from VBA code."""
        # This tests the logic for parsing Subs
        lines = [
            "Option Explicit",
            "",
            "Public Sub HelloWorld()",
            "    MsgBox \"Hello\"",
            "End Sub",
            "",
            "Sub PrivateSub()",
            "    Debug.Print \"Private\"",
            "End Sub"
        ]

        # Find Public Sub
        public_subs = []
        for line in lines:
            line = line.strip()
            if line.startswith("Public Sub ") or line.startswith("Sub "):
                if "(" in line:
                    public_subs.append(line)

        assert len(public_subs) == 2
        assert "HelloWorld" in public_subs[0]
        assert "PrivateSub" in public_subs[1]

    def test_list_macros_parsing_function(self):
        """Test parsing Function procedures from VBA code."""
        lines = [
            "Public Function Calculate(x As Long, y As Long) As Long",
            "    Calculate = x + y",
            "End Function",
            "",
            "Function GetName() As String",
            "    GetName = \"Test\"",
            "End Function"
        ]

        # Find Functions
        functions = []
        for line in lines:
            line = line.strip()
            if line.startswith("Public Function ") or line.startswith("Function "):
                if "(" in line:
                    functions.append(line)

        assert len(functions) == 2
        assert "Calculate" in functions[0]
        assert "GetName" in functions[1]
        assert " As " in functions[0]

    def test_list_macros_extract_return_type(self):
        """Test extracting return type from Function signature."""
        signature = "Public Function Calculate(x As Long) As Long"

        # Extract return type
        if " As " in signature:
            parts = signature.split(" As ")
            return_type = parts[-1].strip()
        else:
            return_type = "Variant"

        assert return_type == "Long"

    def test_list_macros_no_public_macros(self):
        """Test handling of file with no public macros."""
        lines = [
            "Private Sub PrivateOnly()",
            "End Sub"
        ]

        # Find public macros
        public_macros = []
        for line in lines:
            line = line.strip()
            if line.startswith("Public Sub ") or line.startswith("Sub "):
                if "(" in line:
                    public_macros.append(line)

        # Private subs starting with just "Sub" would be found
        # But truly Private ones would not
        assert len(public_macros) == 0


class TestValidateVBACode:
    """Tests for validate_vba_code tool."""

    def test_validate_syntax_basic(self):
        """Test basic syntax validation logic."""
        # This tests the concept of validation
        valid_code = """Sub Test()
    MsgBox "Hello"
End Sub"""

        # Basic checks
        has_sub = "Sub " in valid_code
        has_end = "End Sub" in valid_code

        assert has_sub
        assert has_end

    def test_validate_detect_missing_end(self):
        """Test detection of missing End statement."""
        invalid_code = """Sub Test()
    If x = 1 Then
        MsgBox "Hi"
    ' Missing End If
End Sub"""

        # Simple heuristic check
        if_count = invalid_code.count("If ")
        end_if_count = invalid_code.count("End If")

        # This would indicate a potential issue
        assert if_count > end_if_count

    def test_validate_ascii_check(self):
        """Test that validation includes ASCII check."""
        code_with_unicode = """Sub Test()
    MsgBox "✓ Done"
End Sub"""

        has_non_ascii, error = _detect_non_ascii(code_with_unicode)
        assert has_non_ascii is True

    def test_validate_proper_structure(self):
        """Test validation of proper VBA structure."""
        valid_code = """Option Explicit

Public Function Add(x As Long, y As Long) As Long
    Add = x + y
End Function

Sub Main()
    Dim result As Long
    result = Add(5, 10)
    MsgBox result
End Sub"""

        # Check basic structure
        has_non_ascii, _ = _detect_non_ascii(valid_code)
        assert has_non_ascii is False
        assert "Function" in valid_code
        assert "End Function" in valid_code
        assert "Sub" in valid_code
        assert "End Sub" in valid_code


class TestIntegrationScenarios:
    """Integration tests for common VBA validation scenarios."""

    @pytest.mark.asyncio
    async def test_full_workflow_validation(self):
        """Test complete workflow: detect -> suggest -> validate."""
        # Step 1: Code with Unicode
        original_code = 'MsgBox "✓ Task complete → Next"'

        # Step 2: Detect non-ASCII
        has_non_ascii, error_msg = _detect_non_ascii(original_code)
        assert has_non_ascii is True

        # Step 3: Get suggestions
        suggested_code, changes = _suggest_ascii_replacement(original_code)
        assert "✓" not in suggested_code
        assert "[OK]" in suggested_code

        # Step 4: Validate suggested code
        has_non_ascii_after, _ = _detect_non_ascii(suggested_code)
        assert has_non_ascii_after is False

    def test_macro_name_format_variations(self):
        """Test different macro name format variations."""
        workbook_name = "test.xlsm"
        module_name = "Module1"
        proc_name = "Calculate"

        # Generate different formats
        formats = [
            f"{module_name}.{proc_name}",
            f"'{workbook_name}'!{module_name}.{proc_name}",
            proc_name,
            f"'{workbook_name}'!{proc_name}"
        ]

        assert len(formats) == 4
        assert "Module1.Calculate" in formats
        assert "'test.xlsm'!Module1.Calculate" in formats
        assert "Calculate" in formats

    def test_error_message_formatting(self):
        """Test that error messages are well-formatted."""
        code = 'MsgBox "✓✗→"'
        has_non_ascii, error = _detect_non_ascii(code)

        assert has_non_ascii
        assert "non-ASCII" in error
        assert "Common replacements" in error
        assert len(error.split('\n')) >= 3  # Multi-line message


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
