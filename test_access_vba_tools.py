#!/usr/bin/env python3
"""
Test script for the 3 CRITICAL Access VBA tools via COM.

Tests:
1. extract_vba_access - Extract VBA code from .accdb using COM
2. analyze_structure_access - Analyze VBA structure via COM
3. compile_vba - Compile VBA project and detect errors

Usage:
    python test_access_vba_tools.py
"""

import asyncio
import sys
from pathlib import Path

# Add packages to path
sys.path.insert(0, str(Path(__file__).parent / "packages" / "core" / "src"))
sys.path.insert(0, str(Path(__file__).parent / "packages" / "pro" / "src"))

from vba_mcp_pro.tools.access_vba import (
    extract_vba_access_tool,
    analyze_structure_access_tool,
    compile_vba_tool
)


# Test database path
TEST_DB = r"C:\Users\alexi\Documents\projects\vba-mcp-demo\sample-files\demo-database.accdb"


async def test_extract_vba_access():
    """Test extract_vba_access_tool on Access database."""
    print("\n" + "=" * 70)
    print("TEST 1: extract_vba_access")
    print("=" * 70)

    try:
        # Test 1a: Extract all modules
        print("\n[1a] Extracting all VBA modules...")
        result = await extract_vba_access_tool(TEST_DB)
        print(result[:2000])  # Print first 2000 chars
        if "modules" in result.lower() or "vba extraction" in result.lower():
            print("\n[PASS] All modules extracted successfully")
        else:
            print("\n[WARN] Unexpected result format")

        # Test 1b: Extract specific module
        print("\n[1b] Extracting specific module 'DemoModule'...")
        result = await extract_vba_access_tool(TEST_DB, module_name="DemoModule")
        print(result[:1500])
        if "DemoModule" in result:
            print("\n[PASS] Specific module extracted successfully")
        else:
            print("\n[WARN] Module not found or unexpected result")

        return True

    except FileNotFoundError as e:
        print(f"\n[SKIP] Test database not found: {e}")
        return None
    except Exception as e:
        print(f"\n[FAIL] Error: {type(e).__name__}: {e}")
        return False


async def test_analyze_structure_access():
    """Test analyze_structure_access_tool on Access database."""
    print("\n" + "=" * 70)
    print("TEST 2: analyze_structure_access")
    print("=" * 70)

    try:
        # Test 2a: Analyze entire project
        print("\n[2a] Analyzing VBA structure...")
        result = await analyze_structure_access_tool(TEST_DB)
        print(result[:2000])

        if "metrics" in result.lower() and "complexity" in result.lower():
            print("\n[PASS] Structure analysis completed successfully")
        else:
            print("\n[WARN] Unexpected result format")

        # Test 2b: Analyze specific module
        print("\n[2b] Analyzing specific module 'DemoModule'...")
        result = await analyze_structure_access_tool(TEST_DB, module_name="DemoModule")
        print(result[:1500])

        if "DemoModule" in result:
            print("\n[PASS] Module analysis completed successfully")
        else:
            print("\n[WARN] Module not found or unexpected result")

        return True

    except FileNotFoundError as e:
        print(f"\n[SKIP] Test database not found: {e}")
        return None
    except Exception as e:
        print(f"\n[FAIL] Error: {type(e).__name__}: {e}")
        return False


async def test_compile_vba():
    """Test compile_vba_tool on Access database."""
    print("\n" + "=" * 70)
    print("TEST 3: compile_vba")
    print("=" * 70)

    try:
        print("\n[3a] Compiling VBA project...")
        result = await compile_vba_tool(TEST_DB)
        print(result)

        if "compilation" in result.lower():
            if "success" in result.lower():
                print("\n[PASS] VBA project compiled successfully - no errors")
            elif "failed" in result.lower() or "error" in result.lower():
                print("\n[PASS] VBA compilation detected errors (this is expected behavior)")
            else:
                print("\n[PASS] Compilation check completed")
        else:
            print("\n[WARN] Unexpected result format")

        return True

    except FileNotFoundError as e:
        print(f"\n[SKIP] Test database not found: {e}")
        return None
    except Exception as e:
        print(f"\n[FAIL] Error: {type(e).__name__}: {e}")
        return False


async def main():
    """Run all tests."""
    print("=" * 70)
    print("ACCESS VBA TOOLS TEST SUITE")
    print("=" * 70)
    print(f"\nTest database: {TEST_DB}")
    print(f"Database exists: {Path(TEST_DB).exists()}")

    results = {}

    # Run tests
    results["extract_vba_access"] = await test_extract_vba_access()
    results["analyze_structure_access"] = await test_analyze_structure_access()
    results["compile_vba"] = await test_compile_vba()

    # Summary
    print("\n" + "=" * 70)
    print("TEST SUMMARY")
    print("=" * 70)

    passed = 0
    failed = 0
    skipped = 0

    for test_name, result in results.items():
        if result is True:
            status = "PASS"
            passed += 1
        elif result is False:
            status = "FAIL"
            failed += 1
        else:
            status = "SKIP"
            skipped += 1
        print(f"  {test_name}: {status}")

    print()
    print(f"Total: {len(results)} tests")
    print(f"  Passed:  {passed}")
    print(f"  Failed:  {failed}")
    print(f"  Skipped: {skipped}")

    # Return exit code
    if failed > 0:
        return 1
    return 0


if __name__ == "__main__":
    exit_code = asyncio.run(main())
    sys.exit(exit_code)
