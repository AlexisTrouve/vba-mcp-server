"""
Integration tests for full VBA MCP workflow (v0.4.0).

These tests verify end-to-end functionality combining:
- inject_vba with session_manager
- Post-save verification
- run_macro with AutomationSecurity
- No file corruption on multiple injections

**IMPORTANT**: These tests require:
- Windows OS
- Microsoft Office installed
- pywin32 package
- Trust access to VBA project enabled in Office

Run with: pytest test_integration_full_workflow.py -v --slow
"""

import pytest
import platform
import shutil
from pathlib import Path
from datetime import datetime

# Skip all tests if not on Windows
pytestmark = pytest.mark.skipif(
    platform.system() != "Windows",
    reason="Integration tests require Windows + Office"
)


class TestFullWorkflowIntegration:
    """End-to-end workflow tests."""

    @pytest.fixture
    def sample_xlsm_copy(self, tmp_path):
        """Create a temporary copy of test Excel file."""
        # Create a minimal Excel file for testing
        test_file = tmp_path / f"test_workflow_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsm"

        # Try to copy from test.xlsm if it exists
        source_file = Path(__file__).parent.parent.parent.parent / "test.xlsm"
        if source_file.exists():
            shutil.copy2(source_file, test_file)
        else:
            # Create empty file as fallback
            test_file.write_text("test")

        yield test_file

        # Cleanup
        if test_file.exists():
            try:
                test_file.unlink()
            except:
                pass

    @pytest.mark.asyncio
    @pytest.mark.integration
    @pytest.mark.slow
    @pytest.mark.windows_only
    async def test_inject_and_run_macro_full_workflow(self, sample_xlsm_copy):
        """
        Test complete workflow: inject VBA → verify → run macro → verify no corruption.

        This is the critical end-to-end test for v0.4.0 fixes.
        """
        # Skip if not on Windows
        if platform.system() != "Windows":
            pytest.skip("Test requires Windows + Office")

        # Skip if pywin32 not available
        try:
            import win32com.client
            import pythoncom
        except ImportError:
            pytest.skip("Test requires pywin32")

        # Skip if Office not available
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Quit()
        except:
            pytest.skip("Test requires Microsoft Office")

        from vba_mcp_pro.tools.inject import inject_vba_tool
        from vba_mcp_pro.tools.office_automation import run_macro_tool
        from vba_mcp_pro.session_manager import OfficeSessionManager

        # Clean sessions before test
        manager = OfficeSessionManager.get_instance()
        await manager.close_all_sessions(save=False)

        # Step 1: Inject a test macro
        test_code = """
Option Explicit

Public Function TestFunction() As String
    TestFunction = "Integration test OK"
End Function

Public Sub TestSub()
    MsgBox "Integration test executed"
End Sub
"""

        result = await inject_vba_tool(
            file_path=str(sample_xlsm_copy),
            module_name="IntegrationTest",
            code=test_code,
            create_backup=True
        )

        # Verify injection success
        assert "VBA Injection Successful" in result
        assert "IntegrationTest" in result
        assert "Verified: Yes" in result  # Post-save verification passed

        # Step 2: Verify module exists by reading file
        pythoncom.CoInitialize()
        try:
            app = win32com.client.Dispatch("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False

            workbook = app.Workbooks.Open(str(sample_xlsm_copy), ReadOnly=True)
            vbproject = workbook.VBProject

            # Check module exists
            module_found = False
            for component in vbproject.VBComponents:
                if component.Name == "IntegrationTest":
                    module_found = True
                    code_module = component.CodeModule
                    assert code_module.CountOfLines > 0, "Module is empty"
                    break

            assert module_found, "Module 'IntegrationTest' not found after injection"

            workbook.Close(SaveChanges=False)
            app.Quit()
        finally:
            pythoncom.CoUninitialize()

        # Step 3: Run the injected macro
        result = await run_macro_tool(
            file_path=str(sample_xlsm_copy),
            macro_name="IntegrationTest.TestFunction",
            enable_macros=True
        )

        # Verify macro execution
        assert "Macro Executed Successfully" in result
        assert "Integration test OK" in result or "Return value" in result

        # Step 4: Verify file is not corrupted
        # Try to reopen and read the module again
        pythoncom.CoInitialize()
        try:
            app = win32com.client.Dispatch("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False

            workbook = app.Workbooks.Open(str(sample_xlsm_copy), ReadOnly=True)
            vbproject = workbook.VBProject

            # Module should still exist
            module_found = False
            for component in vbproject.VBComponents:
                if component.Name == "IntegrationTest":
                    module_found = True
                    break

            assert module_found, "Module corrupted or lost after execution"

            workbook.Close(SaveChanges=False)
            app.Quit()
        finally:
            pythoncom.CoUninitialize()

        # Step 5: Close all sessions (cleanup)
        await manager.close_all_sessions(save=False)

    @pytest.mark.asyncio
    @pytest.mark.integration
    @pytest.mark.slow
    @pytest.mark.windows_only
    async def test_multiple_injections_no_corruption(self, sample_xlsm_copy):
        """
        Test that 10 consecutive injections don't corrupt the file.

        This tests the fix for Problem #3 (file corruption from multiple injections).
        """
        if platform.system() != "Windows":
            pytest.skip("Test requires Windows + Office")

        try:
            import win32com.client
            import pythoncom
        except ImportError:
            pytest.skip("Test requires pywin32")

        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Quit()
        except:
            pytest.skip("Test requires Microsoft Office")

        from vba_mcp_pro.tools.inject import inject_vba_tool
        from vba_mcp_pro.session_manager import OfficeSessionManager

        # Clean sessions
        manager = OfficeSessionManager.get_instance()
        await manager.close_all_sessions(save=False)

        # Inject 10 different modules
        injected_modules = []
        for i in range(10):
            module_name = f"Module{i+1}"
            code = f"""
Option Explicit

Public Function Test{i+1}() As Long
    Test{i+1} = {i+1}
End Function
"""

            result = await inject_vba_tool(
                file_path=str(sample_xlsm_copy),
                module_name=module_name,
                code=code,
                create_backup=(i == 0)  # Only backup first injection
            )

            assert "VBA Injection Successful" in result
            assert module_name in result
            injected_modules.append(module_name)

        # Verify all modules exist
        pythoncom.CoInitialize()
        try:
            app = win32com.client.Dispatch("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False

            workbook = app.Workbooks.Open(str(sample_xlsm_copy), ReadOnly=True)
            vbproject = workbook.VBProject

            # Count modules
            found_modules = []
            for component in vbproject.VBComponents:
                if component.Name in injected_modules:
                    found_modules.append(component.Name)

            # All 10 modules should exist
            assert len(found_modules) == 10, f"Expected 10 modules, found {len(found_modules)}"

            # File should be openable (not corrupted)
            assert workbook is not None

            workbook.Close(SaveChanges=False)
            app.Quit()
        finally:
            pythoncom.CoUninitialize()

        # Cleanup
        await manager.close_all_sessions(save=False)

    @pytest.mark.asyncio
    @pytest.mark.integration
    @pytest.mark.slow
    @pytest.mark.windows_only
    async def test_no_zombie_processes(self, sample_xlsm_copy):
        """
        Test that no zombie Excel processes are left after operations.

        This verifies proper COM cleanup (Phase 1.2).
        """
        if platform.system() != "Windows":
            pytest.skip("Test requires Windows + Office")

        try:
            import win32com.client
            import psutil
        except ImportError:
            pytest.skip("Test requires pywin32 and psutil")

        from vba_mcp_pro.tools.inject import inject_vba_tool
        from vba_mcp_pro.session_manager import OfficeSessionManager

        # Count Excel processes before
        def count_excel_processes():
            count = 0
            for proc in psutil.process_iter(['name']):
                try:
                    if proc.info['name'].lower() in ['excel.exe', 'winword.exe']:
                        count += 1
                except:
                    pass
            return count

        processes_before = count_excel_processes()

        # Perform operations
        manager = OfficeSessionManager.get_instance()
        await manager.close_all_sessions(save=False)

        code = "Sub Test()\nEnd Sub"
        await inject_vba_tool(
            file_path=str(sample_xlsm_copy),
            module_name="TestZombie",
            code=code
        )

        # Close all sessions
        await manager.close_all_sessions(save=False)

        # Wait a bit for cleanup
        import asyncio
        await asyncio.sleep(2)

        # Count processes after
        processes_after = count_excel_processes()

        # Should not have more processes than before
        # (Allow same or fewer, as other tests might have left processes)
        assert processes_after <= processes_before + 1, (
            f"Possible zombie processes: {processes_after - processes_before} new Excel processes"
        )


if __name__ == "__main__":
    pytest.main([__file__, "-v", "--tb=short"])
