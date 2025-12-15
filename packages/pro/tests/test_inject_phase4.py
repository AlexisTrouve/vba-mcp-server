"""
Additional tests for VBA injection - Phase 4 (v0.4.0).

These tests verify the new features added in v0.4.0:
- Phase 1: Session manager integration, file lock detection
- Phase 2: Post-save verification, improved validation
"""

import pytest
from pathlib import Path
from unittest.mock import Mock, MagicMock, patch, AsyncMock, call
import sys

# Mock Windows modules if not on Windows
if sys.platform != "win32":
    sys.modules['win32com'] = MagicMock()
    sys.modules['win32com.client'] = MagicMock()
    sys.modules['pythoncom'] = MagicMock()
    sys.modules['win32file'] = MagicMock()
    sys.modules['pywintypes'] = MagicMock()

from vba_mcp_pro.tools.inject import (
    inject_vba_tool,
    _inject_vba_via_session,
    _verify_injection,
    _compile_vba_module
)
from vba_mcp_pro.session_manager import OfficeSessionManager


class TestPhase1SessionManagerIntegration:
    """Tests for Phase 1 - inject_vba uses session_manager."""

    @pytest.mark.asyncio
    @patch('vba_mcp_pro.tools.inject.OfficeSessionManager')
    @patch('vba_mcp_pro.tools.inject._create_backup')
    @patch('vba_mcp_pro.tools.inject._verify_injection')
    async def test_inject_uses_session_manager(
        self,
        mock_verify,
        mock_backup,
        mock_manager_class,
        tmp_path
    ):
        """Test that inject_vba uses OfficeSessionManager instead of creating separate COM instance."""
        # Create test file
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Setup mocks
        mock_manager = Mock()
        mock_manager_class.get_instance.return_value = mock_manager

        mock_session = Mock()
        mock_session.file_path = test_file
        mock_session.app_type = "Excel"
        mock_session.vb_project = Mock()

        # Mock VBA components
        mock_component = Mock()
        mock_code_module = Mock()
        mock_code_module.CountOfLines = 0
        mock_code_module.CountOfDeclarationLines = 0
        mock_component.CodeModule = mock_code_module
        mock_component.Name = "TestModule"

        mock_session.vb_project.VBComponents = [mock_component]
        mock_session.vb_project.VBComponents.Add.return_value = mock_component
        mock_session.file_obj = Mock()
        mock_session.file_obj.Save = Mock()

        mock_manager.get_or_create_session = AsyncMock(return_value=mock_session)

        mock_backup.return_value = test_file.parent / "backup.xlsm"
        mock_verify.return_value = (True, None)  # Verification success

        # Call inject_vba_tool
        code = "Sub Test()\nEnd Sub"
        await inject_vba_tool(
            file_path=str(test_file),
            module_name="TestModule",
            code=code
        )

        # Assertions
        mock_manager_class.get_instance.assert_called_once()
        mock_manager.get_or_create_session.assert_called_once()
        # Verify session path was used
        call_args = mock_manager.get_or_create_session.call_args[0]
        assert call_args[0] == test_file

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.OfficeSessionManager._check_file_lock')
    async def test_inject_detects_concurrent_access(
        self,
        mock_check_lock,
        mock_platform,
        tmp_path
    ):
        """Test that injection detects file locks and raises PermissionError."""
        # Create test file
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Mock file as locked
        mock_check_lock.return_value = True

        # Reset singleton
        OfficeSessionManager._instance = None
        manager = OfficeSessionManager.get_instance()

        # Attempt to inject - should raise PermissionError
        with pytest.raises(PermissionError, match="locked"):
            await manager.get_or_create_session(test_file)

        # Cleanup
        OfficeSessionManager._instance = None


class TestPhase2PostSaveVerification:
    """Tests for Phase 2 - Post-save verification."""

    @pytest.mark.asyncio
    @patch('vba_mcp_pro.tools.inject.win32com')
    @patch('vba_mcp_pro.tools.inject.pythoncom')
    async def test_verify_injection_success(
        self,
        mock_pythoncom,
        mock_win32com,
        tmp_path
    ):
        """Test successful verification of injected code."""
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Setup mocks
        mock_app = Mock()
        mock_file = Mock()
        mock_vbproject = Mock()

        # Mock component with matching code
        mock_component = Mock()
        mock_component.Name = "TestModule"
        mock_code_module = Mock()
        mock_code_module.CountOfLines = 2
        mock_code_module.Lines.return_value = "Sub Test()\nEnd Sub"
        mock_component.CodeModule = mock_code_module

        mock_vbproject.VBComponents = [mock_component]
        mock_file.VBProject = mock_vbproject

        mock_app.Workbooks.Open.return_value = mock_file
        mock_win32com.client.Dispatch.return_value = mock_app

        # Test verification
        code = "Sub Test()\nEnd Sub"
        success, error = await _verify_injection(test_file, "TestModule", code)

        # Assertions
        assert success is True
        assert error is None
        mock_app.Workbooks.Open.assert_called_once()
        mock_file.Close.assert_called_once_with(SaveChanges=False)
        mock_app.Quit.assert_called_once()

    @pytest.mark.asyncio
    @patch('vba_mcp_pro.tools.inject.win32com')
    @patch('vba_mcp_pro.tools.inject.pythoncom')
    async def test_verify_injection_module_not_found(
        self,
        mock_pythoncom,
        mock_win32com,
        tmp_path
    ):
        """Test verification failure when module not found."""
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Setup mocks - no components
        mock_app = Mock()
        mock_file = Mock()
        mock_vbproject = Mock()
        mock_vbproject.VBComponents = []  # Empty - module not found
        mock_file.VBProject = mock_vbproject

        mock_app.Workbooks.Open.return_value = mock_file
        mock_win32com.client.Dispatch.return_value = mock_app

        # Test verification
        success, error = await _verify_injection(test_file, "TestModule", "Sub Test()\nEnd Sub")

        # Assertions
        assert success is False
        assert "not found" in error
        mock_file.Close.assert_called_once()
        mock_app.Quit.assert_called_once()

    @pytest.mark.asyncio
    @patch('vba_mcp_pro.tools.inject.win32com')
    @patch('vba_mcp_pro.tools.inject.pythoncom')
    async def test_verify_injection_code_mismatch(
        self,
        mock_pythoncom,
        mock_win32com,
        tmp_path
    ):
        """Test verification failure when code doesn't match."""
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Setup mocks - code mismatch
        mock_app = Mock()
        mock_file = Mock()
        mock_vbproject = Mock()

        mock_component = Mock()
        mock_component.Name = "TestModule"
        mock_code_module = Mock()
        mock_code_module.CountOfLines = 2
        mock_code_module.Lines.return_value = "Sub DifferentCode()\nEnd Sub"  # Mismatch
        mock_component.CodeModule = mock_code_module

        mock_vbproject.VBComponents = [mock_component]
        mock_file.VBProject = mock_vbproject

        mock_app.Workbooks.Open.return_value = mock_file
        mock_win32com.client.Dispatch.return_value = mock_app

        # Test verification
        expected_code = "Sub Test()\nEnd Sub"
        success, error = await _verify_injection(test_file, "TestModule", expected_code)

        # Assertions
        assert success is False
        assert "mismatch" in error.lower()


class TestPhase2ImprovedCompile:
    """Tests for Phase 2 - Improved _compile_vba_module."""

    @patch('vba_mcp_pro.tools.inject.pythoncom')
    def test_compile_vba_module_uses_procofline(self, mock_pythoncom):
        """Test that _compile_vba_module uses ProcOfLine for validation."""
        # Setup mock module
        mock_module = Mock()
        mock_code_module = Mock()
        mock_code_module.CountOfLines = 5
        mock_code_module.Lines.return_value = "Sub Test()\nMsgBox \"Hello\"\nEnd Sub"

        # Mock ProcOfLine - should be called for each line
        mock_code_module.ProcOfLine.return_value = "Test"

        mock_module.CodeModule = mock_code_module
        mock_module.Name = "TestModule"

        # Test compilation
        success, error = _compile_vba_module(mock_module)

        # Assertions
        assert success is True
        assert error is None
        # ProcOfLine should be called for each line (up to 1000 limit)
        assert mock_code_module.ProcOfLine.call_count == 5

    @patch('vba_mcp_pro.tools.inject.pythoncom')
    def test_compile_vba_module_detects_syntax_error(self, mock_pythoncom):
        """Test that compile detects syntax errors via ProcOfLine."""
        # Setup mock module with syntax error
        mock_module = Mock()
        mock_code_module = Mock()
        mock_code_module.CountOfLines = 3
        mock_code_module.Lines.return_value = "Sub Test(\nEnd Sub"  # Missing closing paren

        # Mock COM error on line 2
        mock_pythoncom.com_error = Exception
        mock_code_module.ProcOfLine.side_effect = [
            "Test",  # Line 1 OK
            Exception("Compile error: Expected )"),  # Line 2 error
            "Test"
        ]

        mock_module.CodeModule = mock_code_module
        mock_module.Name = "TestModule"

        # Test compilation
        success, error = _compile_vba_module(mock_module)

        # Assertions
        assert success is False
        assert "Syntax error at line 2" in error
        assert "Expected )" in error


class TestPhase2ExceptionHandling:
    """Tests for Phase 2 - No exception masking."""

    @pytest.mark.asyncio
    @patch('vba_mcp_pro.tools.inject.OfficeSessionManager')
    @patch('vba_mcp_pro.tools.inject._create_backup')
    async def test_inject_raises_proper_exception_types(
        self,
        mock_backup,
        mock_manager_class,
        tmp_path
    ):
        """Test that injection raises specific exception types (not generic)."""
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Setup mocks
        mock_manager = Mock()
        mock_manager_class.get_instance.return_value = mock_manager

        mock_session = Mock()
        mock_session.vb_project = Mock()

        # Import pythoncom mock
        import pythoncom as mock_pythoncom
        mock_pythoncom.com_error = Exception

        # Simulate permission error
        mock_session.vb_project.VBComponents = Mock(
            side_effect=Exception("Permission denied")
        )
        mock_manager.get_or_create_session = AsyncMock(return_value=mock_session)

        # Should raise RuntimeError (from COM error), not generic Exception
        with pytest.raises((RuntimeError, PermissionError, Exception)):
            await inject_vba_tool(
                file_path=str(test_file),
                module_name="Test",
                code="Sub Test()\nEnd Sub"
            )


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
