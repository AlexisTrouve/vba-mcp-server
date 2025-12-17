"""Tests for VBA injection tool."""
import pytest
from pathlib import Path
import platform
import sys

from vba_mcp_pro.tools.inject import inject_vba_tool, _create_backup


class TestVBAInjection:
    """Test suite for VBA injection."""

    @pytest.mark.asyncio
    async def test_inject_file_not_found(self):
        """Test injection with non-existent file."""
        with pytest.raises(FileNotFoundError):
            await inject_vba_tool(
                file_path="/does/not/exist.xlsm",
                module_name="TestModule",
                code="Sub Test()\nEnd Sub"
            )

    @pytest.mark.asyncio
    async def test_inject_unsupported_platform(self, monkeypatch, sample_xlsm):
        """Test that injection fails on non-Windows platforms."""
        # Mock platform.system to return non-Windows
        monkeypatch.setattr(platform, "system", lambda: "Linux")

        with pytest.raises(RuntimeError, match="only supported on Windows"):
            await inject_vba_tool(
                file_path=str(sample_xlsm),
                module_name="TestModule",
                code="Sub Test()\nEnd Sub"
            )

    @pytest.mark.asyncio
    @pytest.mark.windows_only
    @pytest.mark.slow
    async def test_inject_vba_success(self, sample_xlsm_copy):
        """Test successful VBA injection (Windows + Office required)."""
        # Skip if not on Windows
        if platform.system() != "Windows":
            pytest.skip("Test requires Windows + Office")

        # Skip if pywin32 not available
        try:
            import win32com.client
        except ImportError:
            pytest.skip("Test requires pywin32")

        # Skip if Office not available
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Quit()
        except:
            pytest.skip("Test requires Microsoft Office")

        # New VBA code to inject
        new_code = """
Option Explicit

Public Function NewFunction() As String
    NewFunction = "Injected!"
End Function
"""

        result = await inject_vba_tool(
            file_path=str(sample_xlsm_copy),
            module_name="InjectedModule",
            code=new_code,
            create_backup=True
        )

        assert "VBA Injection Successful" in result
        assert "InjectedModule" in result
        assert "created" in result.lower() or "updated" in result.lower()

        # Verify backup was created
        backup_dir = sample_xlsm_copy.parent / ".vba_backups"
        assert backup_dir.exists()

    @pytest.mark.asyncio
    @pytest.mark.windows_only
    @pytest.mark.slow
    async def test_inject_update_existing_module(self, sample_xlsm_copy):
        """Test updating an existing VBA module."""
        if platform.system() != "Windows":
            pytest.skip("Test requires Windows + Office")

        try:
            import win32com.client
        except ImportError:
            pytest.skip("Test requires pywin32")

        # Updated code for existing module
        updated_code = """
Option Explicit

Public Function HelloWorld() As String
    HelloWorld = "Updated Hello!"
End Function
"""

        result = await inject_vba_tool(
            file_path=str(sample_xlsm_copy),
            module_name="TestModule",
            code=updated_code,
            create_backup=False  # Don't create backup for this test
        )

        assert "VBA Injection Successful" in result
        assert "TestModule" in result

    @pytest.mark.asyncio
    async def test_inject_unsupported_file_type(self, tmp_path):
        """Test injection with unsupported file type."""
        # Create a .txt file
        test_file = tmp_path / "test.txt"
        test_file.write_text("dummy content")

        # Mock Windows platform
        if platform.system() == "Windows":
            with pytest.raises(ValueError, match="Unsupported file type"):
                await inject_vba_tool(
                    file_path=str(test_file),
                    module_name="Test",
                    code="Sub Test()\nEnd Sub"
                )

    def test_create_backup_function(self, sample_xlsm_copy):
        """Test the backup creation helper function."""
        backup_path = _create_backup(sample_xlsm_copy)

        assert backup_path.exists()
        assert backup_path.parent.name == ".vba_backups"
        assert "backup" in backup_path.name
        assert backup_path.suffix == sample_xlsm_copy.suffix

    @pytest.mark.asyncio
    async def test_inject_without_pywin32(self, sample_xlsm):
        """Test that appropriate error is raised when pywin32 is missing."""
        # Only test if on Windows
        if platform.system() != "Windows":
            pytest.skip("Test only relevant on Windows")

        # Skip this test since pywin32 is installed
        # This test would require more complex mocking to work properly
        pytest.skip("pywin32 mocking test skipped - requires advanced import mocking")
