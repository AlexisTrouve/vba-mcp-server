"""Tests for WSL compatibility with graceful property handling."""

import pytest
from unittest.mock import Mock, patch, PropertyMock
import logging


def test_configure_excel_app_all_succeed():
    """Test _configure_excel_app when all properties succeed."""
    from vba_mcp_pro.tools.inject import _configure_excel_app

    mock_app = Mock()
    status = _configure_excel_app(mock_app, visible=False, display_alerts=False)

    assert status['visible'] is True
    assert status['display_alerts'] is True
    assert status['screen_updating'] is True

    assert mock_app.Visible == False
    assert mock_app.DisplayAlerts == False
    assert mock_app.ScreenUpdating == False


def test_configure_excel_app_visible_fails():
    """Test _configure_excel_app when Visible property fails (WSL scenario)."""
    from vba_mcp_pro.tools.inject import _configure_excel_app

    mock_app = Mock()
    # Simulate WSL error when setting Visible
    type(mock_app).Visible = PropertyMock(side_effect=Exception("Property 'Excel.Application.Visible' can not be set."))

    # Should NOT raise exception - should log warning and continue
    status = _configure_excel_app(mock_app, visible=False, display_alerts=False)

    # Visible failed, but others succeeded
    assert status['visible'] is False
    assert status['display_alerts'] is True  # Should still succeed
    assert status['screen_updating'] is True


def test_configure_excel_app_all_fail():
    """Test _configure_excel_app when all properties fail."""
    from vba_mcp_pro.tools.inject import _configure_excel_app

    mock_app = Mock()
    # All properties raise exceptions
    type(mock_app).Visible = PropertyMock(side_effect=Exception("Cannot set Visible"))
    type(mock_app).DisplayAlerts = PropertyMock(side_effect=Exception("Cannot set DisplayAlerts"))
    type(mock_app).ScreenUpdating = PropertyMock(side_effect=Exception("Cannot set ScreenUpdating"))

    # Should NOT raise exception
    status = _configure_excel_app(mock_app, visible=False, display_alerts=False)

    # All failed gracefully
    assert status['visible'] is False
    assert status['display_alerts'] is False
    assert status['screen_updating'] is False


def test_wsl_detection_in_wsl(tmp_path):
    """Test WSL detection when running in WSL."""
    from vba_mcp_pro.server import _check_wsl_environment

    # Mock /proc/version to simulate WSL
    proc_version = tmp_path / "version"
    proc_version.write_text("Linux version 5.10.16.3-microsoft-standard-WSL2")

    with patch("platform.system", return_value="Linux"):
        with patch("builtins.open", return_value=open(proc_version)):
            is_wsl = _check_wsl_environment()
            assert is_wsl is True


def test_wsl_detection_on_windows():
    """Test WSL detection when running on native Windows."""
    from vba_mcp_pro.server import _check_wsl_environment

    with patch("platform.system", return_value="Windows"):
        is_wsl = _check_wsl_environment()
        assert is_wsl is False


def test_wsl_detection_on_linux_not_wsl():
    """Test WSL detection on native Linux (not WSL)."""
    from vba_mcp_pro.server import _check_wsl_environment

    with patch("platform.system", return_value="Linux"):
        with patch("builtins.open", side_effect=FileNotFoundError):
            is_wsl = _check_wsl_environment()
            assert is_wsl is False


@patch('win32com.client.Dispatch')
def test_verify_injection_with_visible_failure(mock_dispatch):
    """Test _verify_injection continues when Visible fails."""
    from vba_mcp_pro.tools.inject import _verify_injection
    from pathlib import Path
    import asyncio

    # Mock Excel application
    mock_app = Mock()
    mock_workbook = Mock()
    mock_vb_project = Mock()
    mock_component = Mock()
    mock_code_module = Mock()

    # Simulate Visible property failure (WSL scenario)
    type(mock_app).Visible = PropertyMock(
        side_effect=Exception("Property 'Excel.Application.Visible' can not be set.")
    )

    mock_app.Workbooks.Open.return_value = mock_workbook
    mock_workbook.VBProject = mock_vb_project
    mock_component.Name = "TestModule"
    mock_code_module.CountOfLines = 2
    mock_code_module.Lines.return_value = "Sub Test()\nEnd Sub"
    mock_component.CodeModule = mock_code_module
    mock_vb_project.VBComponents = [mock_component]

    mock_dispatch.return_value = mock_app

    # Should not raise exception despite Visible failure
    success, error = asyncio.run(_verify_injection(
        Path("test.xlsm"),
        "TestModule",
        "Sub Test()\nEnd Sub"
    ))

    # Should succeed despite Visible error
    assert success is True
    assert error is None


def test_configure_excel_app_logs_warnings(caplog):
    """Test that _configure_excel_app logs appropriate warnings."""
    from vba_mcp_pro.tools.inject import _configure_excel_app

    mock_app = Mock()
    type(mock_app).Visible = PropertyMock(side_effect=Exception("Cannot set Visible"))

    with caplog.at_level(logging.WARNING):
        status = _configure_excel_app(mock_app, visible=False, display_alerts=False)

    # Check that warning was logged
    assert any("Could not set Excel.Visible=False" in record.message for record in caplog.records)
    assert status['visible'] is False


def test_configure_excel_app_visible_true():
    """Test _configure_excel_app with visible=True."""
    from vba_mcp_pro.tools.inject import _configure_excel_app

    mock_app = Mock()
    status = _configure_excel_app(mock_app, visible=True, display_alerts=True)

    assert status['visible'] is True
    assert status['display_alerts'] is True
    assert mock_app.Visible == True
    assert mock_app.DisplayAlerts == True


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
