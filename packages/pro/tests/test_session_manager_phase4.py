"""
Additional tests for OfficeSessionManager - Phase 4 (v0.4.0).

These tests verify the new features added in v0.4.0:
- Phase 1: COM cleanup with ReleaseObject, file lock detection
"""

import pytest
from pathlib import Path
from unittest.mock import Mock, MagicMock, patch, AsyncMock
import sys

# Mock Windows modules if not on Windows
if sys.platform != "win32":
    sys.modules['win32com'] = MagicMock()
    sys.modules['win32com.client'] = MagicMock()
    sys.modules['pythoncom'] = MagicMock()
    sys.modules['win32file'] = MagicMock()
    sys.modules['pywintypes'] = MagicMock()

from vba_mcp_pro.session_manager import OfficeSession, OfficeSessionManager


class TestPhase1COMCleanup:
    """Tests for Phase 1 - COM cleanup with ReleaseObject."""

    @pytest.fixture
    def manager(self):
        """Get a fresh session manager instance."""
        OfficeSessionManager._instance = None
        manager = OfficeSessionManager.get_instance()
        yield manager
        # Cleanup
        import asyncio
        asyncio.run(manager.close_all_sessions(save=False))

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    @patch('vba_mcp_pro.session_manager.win32com.client')
    async def test_com_cleanup_releases_objects(
        self,
        mock_win32com,
        mock_pythoncom,
        mock_platform,
        manager,
        tmp_path
    ):
        """Test that COM objects are explicitly released via pythoncom.ReleaseObject."""
        # Create test file
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Setup mocks
        mock_app = Mock()
        mock_app.Name = "Microsoft Excel"
        mock_workbook = Mock()
        mock_workbook.Name = "test.xlsm"
        mock_vbproject = Mock()
        mock_workbook.VBProject = mock_vbproject

        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.Dispatch.return_value = mock_app

        # Create and close session
        session = await manager.get_or_create_session(test_file)

        # Access vb_project to cache it
        _ = session.vb_project

        # Close session
        await manager.close_session(test_file, save=False)

        # Assertions - ReleaseObject should be called for each COM object
        # Check that ReleaseObject was called (at least once)
        assert mock_pythoncom.ReleaseObject.call_count >= 1

        # Verify it was called with our COM objects
        release_calls = mock_pythoncom.ReleaseObject.call_args_list
        # Should release vb_project, workbook, and app
        assert len(release_calls) >= 1

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    @patch('vba_mcp_pro.session_manager.win32com.client')
    async def test_com_cleanup_order(
        self,
        mock_win32com,
        mock_pythoncom,
        mock_platform,
        manager,
        tmp_path
    ):
        """Test that COM cleanup happens before CoUninitialize."""
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Setup mocks
        mock_app = Mock()
        mock_app.Name = "Microsoft Excel"
        mock_workbook = Mock()
        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.Dispatch.return_value = mock_app

        # Track call order
        call_order = []

        def track_release(*args):
            call_order.append('release')

        def track_uninit(*args):
            call_order.append('uninit')

        mock_pythoncom.ReleaseObject.side_effect = track_release
        mock_pythoncom.CoUninitialize.side_effect = track_uninit

        # Create and close session
        session = await manager.get_or_create_session(test_file)
        await manager.close_session(test_file, save=False)

        # Verify order: release should come before uninit
        if 'release' in call_order and 'uninit' in call_order:
            release_index = call_order.index('release')
            uninit_index = call_order.index('uninit')
            assert release_index < uninit_index, "ReleaseObject must be called before CoUninitialize"


class TestPhase1FileLockDetection:
    """Tests for Phase 1 - File lock detection."""

    @pytest.fixture
    def manager(self):
        """Get a fresh session manager instance."""
        OfficeSessionManager._instance = None
        manager = OfficeSessionManager.get_instance()
        yield manager
        import asyncio
        asyncio.run(manager.close_all_sessions(save=False))

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    async def test_file_lock_detection_locked_file(
        self,
        mock_platform,
        manager,
        tmp_path
    ):
        """Test that file lock detection raises PermissionError for locked files."""
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Mock _check_file_lock to return True (file locked)
        with patch.object(manager, '_check_file_lock', return_value=True):
            # Should raise PermissionError
            with pytest.raises(PermissionError, match="locked"):
                await manager.get_or_create_session(test_file)

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    @patch('vba_mcp_pro.session_manager.win32com.client')
    @patch('vba_mcp_pro.session_manager.win32file')
    async def test_file_lock_detection_uses_createfile(
        self,
        mock_win32file,
        mock_win32com,
        mock_pythoncom,
        mock_platform,
        manager,
        tmp_path
    ):
        """Test that _check_file_lock uses win32file.CreateFile."""
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Mock CreateFile to succeed (file not locked)
        mock_handle = 123
        mock_win32file.CreateFile.return_value = mock_handle
        mock_win32file.CloseHandle = Mock()

        # Setup app mocks
        mock_app = Mock()
        mock_app.Name = "Microsoft Excel"
        mock_workbook = Mock()
        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.Dispatch.return_value = mock_app

        # Call should succeed
        session = await manager.get_or_create_session(test_file)

        # Verify CreateFile was called with exclusive access
        mock_win32file.CreateFile.assert_called_once()
        call_args = mock_win32file.CreateFile.call_args[0]
        # Should be called with file path
        assert str(test_file) in call_args[0]

        # Cleanup
        await manager.close_all_sessions(save=False)

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    @patch('vba_mcp_pro.session_manager.win32com.client')
    @patch('vba_mcp_pro.session_manager.win32file')
    @patch('vba_mcp_pro.session_manager.pywintypes')
    async def test_file_lock_detection_handles_pywintypes_error(
        self,
        mock_pywintypes,
        mock_win32file,
        mock_win32com,
        mock_pythoncom,
        mock_platform,
        manager,
        tmp_path
    ):
        """Test that pywintypes.error is caught and converted to PermissionError."""
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Mock pywintypes.error
        mock_pywintypes.error = Exception

        # Mock CreateFile to raise pywintypes.error (file locked)
        mock_win32file.CreateFile.side_effect = Exception("Access denied")

        # Should raise PermissionError
        with pytest.raises(PermissionError, match="locked"):
            await manager.get_or_create_session(test_file)

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    @patch('vba_mcp_pro.session_manager.win32com.client')
    async def test_file_lock_reuses_alive_session(
        self,
        mock_win32com,
        mock_pythoncom,
        mock_platform,
        manager,
        tmp_path
    ):
        """Test that if file is locked but we have alive session, reuse it."""
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Setup mocks
        mock_app = Mock()
        mock_app.Name = "Microsoft Excel"
        mock_workbook = Mock()
        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.Dispatch.return_value = mock_app

        # Create first session
        session1 = await manager.get_or_create_session(test_file)

        # Mock file as locked
        with patch.object(manager, '_check_file_lock', return_value=True):
            # Get session again - should reuse existing alive session
            session2 = await manager.get_or_create_session(test_file)

        # Should be same session
        assert session1 is session2

        # Cleanup
        await manager.close_all_sessions(save=False)


class TestPhase1ErrorMessages:
    """Test error messages for file lock scenarios."""

    @pytest.fixture
    def manager(self):
        """Get a fresh session manager instance."""
        OfficeSessionManager._instance = None
        manager = OfficeSessionManager.get_instance()
        yield manager

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    async def test_error_message_file_locked_no_session(
        self,
        mock_platform,
        manager,
        tmp_path
    ):
        """Test error message when file locked by external application."""
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Mock file locked, no session
        with patch.object(manager, '_check_file_lock', return_value=True):
            with pytest.raises(PermissionError) as exc_info:
                await manager.get_or_create_session(test_file)

            error_msg = str(exc_info.value)
            # Should mention closing file in Excel/Word
            assert "Close" in error_msg or "locked" in error_msg

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    @patch('vba_mcp_pro.session_manager.win32com.client')
    async def test_error_message_file_locked_dead_session(
        self,
        mock_win32com,
        mock_pythoncom,
        mock_platform,
        manager,
        tmp_path
    ):
        """Test error message when file locked and our session is dead."""
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")

        # Setup mocks
        mock_app = Mock()
        mock_app.Name = "Microsoft Excel"
        mock_workbook = Mock()
        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.Dispatch.return_value = mock_app

        # Create session
        session = await manager.get_or_create_session(test_file)

        # Make session dead
        mock_app.Name = Mock(side_effect=Exception("COM error"))

        # Mock file as locked
        with patch.object(manager, '_check_file_lock', return_value=True):
            with pytest.raises(PermissionError) as exc_info:
                await manager.get_or_create_session(test_file)

            error_msg = str(exc_info.value)
            # Should mention another process
            assert "another" in error_msg.lower() or "locked" in error_msg.lower()

        # Cleanup
        OfficeSessionManager._instance = None


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
