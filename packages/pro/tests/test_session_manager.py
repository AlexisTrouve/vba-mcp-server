"""
Unit tests for OfficeSessionManager.
"""

import asyncio
import pytest
from datetime import datetime, timedelta
from pathlib import Path
from unittest.mock import Mock, MagicMock, patch, AsyncMock
import sys

# Mock Windows modules if not on Windows
if sys.platform != "win32":
    sys.modules['win32com'] = MagicMock()
    sys.modules['win32com.client'] = MagicMock()
    sys.modules['pythoncom'] = MagicMock()

from vba_mcp_pro.session_manager import OfficeSession, OfficeSessionManager


class TestOfficeSession:
    """Tests for OfficeSession class."""

    def test_session_creation(self):
        """Test creating an OfficeSession."""
        mock_app = Mock()
        mock_app.Name = "Microsoft Excel"
        mock_file = Mock()
        mock_file.Name = "test.xlsm"
        file_path = Path("/test/test.xlsm")

        session = OfficeSession(
            app=mock_app,
            file_obj=mock_file,
            file_path=file_path,
            app_type="Excel",
            read_only=False
        )

        assert session.app == mock_app
        assert session.file_obj == mock_file
        assert session.file_path == file_path
        assert session.app_type == "Excel"
        assert session.read_only is False
        assert isinstance(session.opened_at, datetime)
        assert isinstance(session.last_accessed, datetime)

    def test_is_alive_true(self):
        """Test is_alive returns True when COM objects valid."""
        mock_app = Mock()
        mock_app.Name = "Microsoft Excel"
        mock_file = Mock()
        mock_file.Name = "test.xlsm"

        session = OfficeSession(
            app=mock_app,
            file_obj=mock_file,
            file_path=Path("/test/test.xlsm"),
            app_type="Excel"
        )

        assert session.is_alive() is True

    def test_is_alive_false_on_exception(self):
        """Test is_alive returns False when COM objects dead."""
        mock_app = Mock()
        mock_app.Name = Mock(side_effect=Exception("COM error"))
        mock_file = Mock()

        session = OfficeSession(
            app=mock_app,
            file_obj=mock_file,
            file_path=Path("/test/test.xlsm"),
            app_type="Excel"
        )

        assert session.is_alive() is False

    def test_refresh_last_accessed(self):
        """Test refreshing last accessed timestamp."""
        session = OfficeSession(
            app=Mock(),
            file_obj=Mock(),
            file_path=Path("/test/test.xlsm"),
            app_type="Excel"
        )

        original_time = session.last_accessed
        asyncio.run(asyncio.sleep(0.01))  # Small delay
        session.refresh_last_accessed()

        assert session.last_accessed > original_time

    def test_vb_project_excel(self):
        """Test accessing VBProject for Excel."""
        mock_file = Mock()
        mock_vbproject = Mock()
        mock_file.VBProject = mock_vbproject

        session = OfficeSession(
            app=Mock(),
            file_obj=mock_file,
            file_path=Path("/test/test.xlsm"),
            app_type="Excel"
        )

        assert session.vb_project == mock_vbproject
        # Test caching
        assert session.vb_project == mock_vbproject
        mock_file.VBProject  # Should only be called once due to caching


class TestOfficeSessionManager:
    """Tests for OfficeSessionManager class."""

    @pytest.fixture
    def manager(self):
        """Get a fresh session manager instance."""
        # Reset singleton
        OfficeSessionManager._instance = None
        manager = OfficeSessionManager.get_instance()
        yield manager
        # Cleanup
        asyncio.run(manager.close_all_sessions(save=False))

    @pytest.fixture
    def temp_file(self, tmp_path):
        """Create a temporary file for testing."""
        file_path = tmp_path / "test.xlsm"
        file_path.write_text("test")
        return file_path

    def test_singleton_pattern(self):
        """Test that OfficeSessionManager is a singleton."""
        manager1 = OfficeSessionManager.get_instance()
        manager2 = OfficeSessionManager.get_instance()
        assert manager1 is manager2

    @pytest.mark.asyncio
    async def test_create_session_file_not_found(self, manager):
        """Test error when file doesn't exist."""
        with pytest.raises(FileNotFoundError):
            await manager.get_or_create_session(Path("/nonexistent/file.xlsm"))

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Linux')
    async def test_create_session_non_windows(self, mock_platform, manager, temp_file):
        """Test error when not on Windows."""
        with pytest.raises(RuntimeError, match="Windows"):
            await manager.get_or_create_session(temp_file)

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    @patch('vba_mcp_pro.session_manager.win32com.client')
    async def test_create_session_excel(
        self,
        mock_win32com,
        mock_pythoncom,
        mock_platform,
        manager,
        temp_file
    ):
        """Test creating an Excel session."""
        # Setup mocks
        mock_app = Mock()
        mock_app.Name = "Microsoft Excel"
        mock_app.Visible = False
        mock_app.DisplayAlerts = False
        mock_workbook = Mock()
        mock_workbook.Name = "test.xlsm"
        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.Dispatch.return_value = mock_app

        # Create session
        session = await manager.get_or_create_session(temp_file)

        # Assertions
        assert session.app_type == "Excel"
        assert session.file_path.name == "test.xlsm"
        assert not session.read_only
        mock_win32com.Dispatch.assert_called_once_with("Excel.Application")
        mock_app.Workbooks.Open.assert_called_once()
        assert mock_app.Visible is True  # Should be set to True

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    @patch('vba_mcp_pro.session_manager.win32com.client')
    async def test_reuse_existing_session(
        self,
        mock_win32com,
        mock_pythoncom,
        mock_platform,
        manager,
        temp_file
    ):
        """Test reusing an existing session."""
        # Setup mocks
        mock_app = Mock()
        mock_app.Name = "Microsoft Excel"
        mock_workbook = Mock()
        mock_workbook.Name = "test.xlsm"
        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.Dispatch.return_value = mock_app

        # Create first session
        session1 = await manager.get_or_create_session(temp_file)

        # Get same session again
        session2 = await manager.get_or_create_session(temp_file)

        # Should be same object
        assert session1 is session2
        # Dispatch should only be called once
        assert mock_win32com.Dispatch.call_count == 1

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    @patch('vba_mcp_pro.session_manager.win32com.client')
    async def test_close_session(
        self,
        mock_win32com,
        mock_pythoncom,
        mock_platform,
        manager,
        temp_file
    ):
        """Test closing a session."""
        # Setup mocks
        mock_app = Mock()
        mock_app.Name = "Microsoft Excel"
        mock_workbook = Mock()
        mock_workbook.Name = "test.xlsm"
        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.Dispatch.return_value = mock_app

        # Create session
        session = await manager.get_or_create_session(temp_file)

        # Close session
        await manager.close_session(temp_file, save=True)

        # Verify cleanup
        mock_workbook.Save.assert_called_once()
        mock_workbook.Close.assert_called_once()
        mock_app.Quit.assert_called_once()
        mock_pythoncom.CoUninitialize.assert_called()

        # Session should be removed
        assert str(temp_file.resolve()) not in manager._sessions

    @pytest.mark.asyncio
    async def test_close_nonexistent_session(self, manager, temp_file):
        """Test error when closing non-existent session."""
        with pytest.raises(ValueError, match="No open session"):
            await manager.close_session(temp_file)

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    @patch('vba_mcp_pro.session_manager.win32com.client')
    async def test_close_all_sessions(
        self,
        mock_win32com,
        mock_pythoncom,
        mock_platform,
        manager,
        tmp_path
    ):
        """Test closing all sessions."""
        # Create multiple temp files
        file1 = tmp_path / "test1.xlsm"
        file1.write_text("test")
        file2 = tmp_path / "test2.xlsm"
        file2.write_text("test")

        # Setup mocks
        mock_app = Mock()
        mock_app.Name = "Microsoft Excel"
        mock_workbook = Mock()
        mock_workbook.Name = "test.xlsm"
        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.Dispatch.return_value = mock_app

        # Create multiple sessions
        await manager.get_or_create_session(file1)
        await manager.get_or_create_session(file2)

        assert len(manager._sessions) == 2

        # Close all
        await manager.close_all_sessions(save=False)

        # All should be closed
        assert len(manager._sessions) == 0
        assert mock_app.Quit.call_count == 2

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    @patch('vba_mcp_pro.session_manager.win32com.client')
    async def test_cleanup_stale_sessions(
        self,
        mock_win32com,
        mock_pythoncom,
        mock_platform,
        manager,
        temp_file
    ):
        """Test automatic cleanup of stale sessions."""
        # Setup mocks
        mock_app = Mock()
        mock_app.Name = "Microsoft Excel"
        mock_workbook = Mock()
        mock_workbook.Name = "test.xlsm"
        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.Dispatch.return_value = mock_app

        # Set very short timeout for testing
        manager.SESSION_TIMEOUT = 1  # 1 second

        # Create session
        session = await manager.get_or_create_session(temp_file)

        # Manually age the session
        session.last_accessed = datetime.now() - timedelta(seconds=2)

        # Run cleanup manually (instead of waiting for background task)
        async with manager._lock:
            now = datetime.now()
            stale_keys = []
            for file_key, sess in manager._sessions.items():
                age = now - sess.last_accessed
                if age.total_seconds() > manager.SESSION_TIMEOUT:
                    stale_keys.append(file_key)

            for file_key in stale_keys:
                sess = manager._sessions[file_key]
                await manager._close_session_internal(sess, save=True)
                del manager._sessions[file_key]

        # Session should be cleaned up
        assert len(manager._sessions) == 0
        mock_app.Quit.assert_called_once()

    @pytest.mark.asyncio
    @patch('platform.system', return_value='Windows')
    @patch('vba_mcp_pro.session_manager.pythoncom')
    @patch('vba_mcp_pro.session_manager.win32com.client')
    async def test_detect_dead_session(
        self,
        mock_win32com,
        mock_pythoncom,
        mock_platform,
        manager,
        temp_file
    ):
        """Test detection and recreation of dead sessions."""
        # Setup mocks
        mock_app = Mock()
        mock_app.Name = "Microsoft Excel"
        mock_workbook = Mock()
        mock_workbook.Name = "test.xlsm"
        mock_app.Workbooks.Open.return_value = mock_workbook
        mock_win32com.Dispatch.return_value = mock_app

        # Create session
        session1 = await manager.get_or_create_session(temp_file)

        # Make session appear dead
        mock_app.Name = Mock(side_effect=Exception("COM error"))

        # Get session again - should detect dead and recreate
        mock_app.Name = "Microsoft Excel"  # Fix for new session
        session2 = await manager.get_or_create_session(temp_file)

        # Should be different session
        assert session1 is not session2
        # Dispatch should be called twice (original + recreation)
        assert mock_win32com.Dispatch.call_count == 2

    def test_list_sessions(self, manager):
        """Test listing sessions."""
        # No sessions initially
        sessions = manager.list_sessions()
        assert sessions == []

        # Add mock session manually
        mock_session = Mock()
        mock_session.file_path = Path("/test/test.xlsm")
        mock_session.app_type = "Excel"
        mock_session.read_only = False
        mock_session.opened_at = datetime.now()
        mock_session.last_accessed = datetime.now()

        manager._sessions["/test/test.xlsm"] = mock_session

        # List sessions
        sessions = manager.list_sessions()
        assert len(sessions) == 1
        assert sessions[0]["file_name"] == "test.xlsm"
        assert sessions[0]["app_type"] == "Excel"
        assert sessions[0]["read_only"] is False

    @pytest.mark.asyncio
    async def test_cleanup_task_lifecycle(self, manager):
        """Test starting and stopping cleanup task."""
        # Start cleanup task
        manager.start_cleanup_task()
        assert manager._cleanup_task is not None
        assert not manager._cleanup_task.done()

        # Stop cleanup task
        await manager.stop_cleanup_task()
        assert manager._cleanup_task.done()


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
