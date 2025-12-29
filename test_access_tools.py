"""
Test Access tools for VBA MCP Server v0.6.0
"""
import sys
import os
import asyncio
import traceback

# Add paths
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "packages/core/src"))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "packages/pro/src"))

from vba_mcp_pro.tools.office_automation import (
    list_access_tables_tool,
    list_access_queries_tool,
    run_access_query_tool,
    get_worksheet_data_tool,
    set_worksheet_data_tool,
)
from vba_mcp_pro.tools.inject import inject_vba_tool

DB_PATH = r"C:\Users\alexi\Documents\projects\vba-mcp-demo\sample-files\demo-database.accdb"

async def test_list_access_tables():
    print("\n" + "="*60)
    print("TEST 1: list_access_tables_tool")
    print("="*60)

    try:
        result = await list_access_tables_tool(DB_PATH)
        print(result[:500] if len(result) > 500 else result)

        # Check expected content
        if "Employees" in result and "Projects" in result:
            print("\n[OK] Found expected tables")
            return True
        else:
            print("\n[WARN] Missing expected tables")
            return False
    except Exception as e:
        print(f"[ERROR] {e}")
        traceback.print_exc()
        return False

async def test_list_access_queries():
    print("\n" + "="*60)
    print("TEST 2: list_access_queries_tool")
    print("="*60)

    try:
        result = await list_access_queries_tool(DB_PATH)
        print(result[:500] if len(result) > 500 else result)

        if "qryITEmployees" in result:
            print("\n[OK] Found expected queries")
            return True
        else:
            print("\n[WARN] Missing expected queries")
            return False
    except Exception as e:
        print(f"[ERROR] {e}")
        traceback.print_exc()
        return False

async def test_run_access_query():
    print("\n" + "="*60)
    print("TEST 3: run_access_query_tool (saved query)")
    print("="*60)

    try:
        result = await run_access_query_tool(DB_PATH, query_name="qryITEmployees")
        print(result[:800] if len(result) > 800 else result)

        if "IT" in result and "rows" in result:
            print("\n[OK] Query executed successfully")
            return True
        else:
            print("\n[WARN] Unexpected result")
            return False
    except Exception as e:
        print(f"[ERROR] {e}")
        traceback.print_exc()
        return False

async def test_run_access_query_sql():
    print("\n" + "="*60)
    print("TEST 4: run_access_query_tool (custom SQL)")
    print("="*60)

    sql = "SELECT FirstName, LastName, Salary FROM Employees WHERE Salary > 50000"
    try:
        result = await run_access_query_tool(DB_PATH, sql=sql)
        print(result[:800] if len(result) > 800 else result)

        if "rows" in result:
            print("\n[OK] SQL query executed")
            return True
        else:
            print("\n[WARN] Unexpected result")
            return False
    except Exception as e:
        print(f"[ERROR] {e}")
        traceback.print_exc()
        return False

async def test_get_worksheet_data():
    print("\n" + "="*60)
    print("TEST 5: get_worksheet_data_tool (Access mode)")
    print("="*60)

    try:
        result = await get_worksheet_data_tool(
            DB_PATH,
            sheet_name="Employees",
            where_clause="Department = 'Finance'",
            order_by="Salary DESC"
        )
        print(result[:800] if len(result) > 800 else result)

        if "Finance" in result or "rows" in result or "data" in result:
            print("\n[OK] Data retrieved")
            return True
        else:
            print("\n[WARN] Unexpected result")
            return False
    except Exception as e:
        print(f"[ERROR] {e}")
        traceback.print_exc()
        return False

async def test_set_worksheet_data():
    print("\n" + "="*60)
    print("TEST 6: set_worksheet_data_tool (Access append)")
    print("="*60)

    # Data must be 2D array [[row1], [row2], ...]
    # Columns: FirstName, LastName, Department, Salary, HireDate (skip EmployeeID = AutoNumber)
    new_data = [
        ["TestUser", "Temporary", "QA", 40000, None]
    ]
    columns = ["FirstName", "LastName", "Department", "Salary", "HireDate"]

    try:
        result = await set_worksheet_data_tool(
            DB_PATH,
            sheet_name="Employees",
            data=new_data,
            columns=columns,
            mode="append"
        )
        print(result[:500] if len(result) > 500 else result)

        # Check if insert succeeded (cleanup will be done manually later)
        if "Records inserted: 1" in result or "success" in result.lower():
            print("\n[OK] Data inserted successfully")
            print("[INFO] Cleanup skipped (run_access_query doesn't support DELETE)")
            return True
        else:
            print("\n[WARN] Check result")
            return False
    except Exception as e:
        print(f"[ERROR] {e}")
        traceback.print_exc()
        return False

async def test_extract_vba():
    """Test VBA extraction via session manager (oletools doesn't support Access)"""
    print("\n" + "="*60)
    print("TEST 7: extract_vba via session manager (Access)")
    print("="*60)

    try:
        from pathlib import Path
        from vba_mcp_pro.session_manager import OfficeSessionManager

        manager = OfficeSessionManager.get_instance()
        path = Path(DB_PATH).resolve()

        # Get existing session (created by previous tests)
        session = await manager.get_or_create_session(path, read_only=True)
        session.refresh_last_accessed()

        vb_project = session.vb_project
        modules_found = []

        for component in vb_project.VBComponents:
            modules_found.append(component.Name)
            if component.Name == "DemoModule":
                code_module = component.CodeModule
                if code_module.CountOfLines > 0:
                    code = code_module.Lines(1, code_module.CountOfLines)
                    print(f"Module: DemoModule")
                    print(f"Lines: {code_module.CountOfLines}")
                    print(f"Code preview:\n{code[:300]}...")

        if "DemoModule" in modules_found:
            print(f"\n[OK] Found modules: {modules_found}")
            return True
        else:
            print(f"\n[WARN] DemoModule not found. Available: {modules_found}")
            return False

    except Exception as e:
        print(f"[ERROR] {e}")
        traceback.print_exc()
        return False

async def test_inject_vba():
    print("\n" + "="*60)
    print("TEST 8: inject_vba_tool (Access)")
    print("="*60)

    new_code = '''Sub TestInjected()
    MsgBox "Injected from MCP!", vbInformation, "Test"
End Sub
'''

    try:
        result = await inject_vba_tool(DB_PATH, module_name="InjectedModule", code=new_code)
        print(result[:500] if len(result) > 500 else result)

        if "success" in result.lower() or "injected" in result.lower():
            print("\n[OK] VBA injected")
            return True
        else:
            print("\n[WARN] Check result")
            return False
    except Exception as e:
        print(f"[ERROR] {e}")
        traceback.print_exc()
        return False

async def main():
    print("="*60)
    print("VBA MCP Server v0.6.0 - Access Tools Test Suite")
    print("="*60)
    print(f"Database: {DB_PATH}")

    if not os.path.exists(DB_PATH):
        print(f"[ERROR] Database not found: {DB_PATH}")
        return

    results = []

    # Run all tests
    results.append(("list_access_tables", await test_list_access_tables()))
    results.append(("list_access_queries", await test_list_access_queries()))
    results.append(("run_access_query (saved)", await test_run_access_query()))
    results.append(("run_access_query (SQL)", await test_run_access_query_sql()))
    results.append(("get_worksheet_data", await test_get_worksheet_data()))
    results.append(("set_worksheet_data", await test_set_worksheet_data()))
    results.append(("extract_vba", await test_extract_vba()))
    results.append(("inject_vba", await test_inject_vba()))

    # Summary
    print("\n" + "="*60)
    print("TEST SUMMARY")
    print("="*60)

    passed = sum(1 for _, r in results if r)
    total = len(results)

    for name, result in results:
        status = "[PASS]" if result else "[FAIL]"
        print(f"  {status} {name}")

    print("-"*60)
    print(f"Results: {passed}/{total} tests passed ({100*passed//total}%)")
    print("="*60)

if __name__ == "__main__":
    asyncio.run(main())
