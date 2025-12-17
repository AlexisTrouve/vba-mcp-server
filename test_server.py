#!/usr/bin/env python3
"""
Quick test to verify VBA MCP Pro server can list tools.
Updated for v0.3.0 with Excel Tables support.
"""

import asyncio
import sys
from pathlib import Path

# Add packages to path
repo_root = Path(__file__).parent
sys.path.insert(0, str(repo_root / "packages" / "core" / "src"))
sys.path.insert(0, str(repo_root / "packages" / "lite" / "src"))
sys.path.insert(0, str(repo_root / "packages" / "pro" / "src"))

from vba_mcp_pro.server import app


async def test_list_tools():
    """Test that server can list all tools."""
    print("Testing VBA MCP Pro Server v0.3.0...")
    print("=" * 60)

    try:
        # Get tool list
        tools = await app.list_tools()

        print(f"\n[OK] Server loaded successfully!")
        print(f"[OK] Found {len(tools)} tools\n")

        # Expected tool count
        EXPECTED_TOOL_COUNT = 21

        # Group tools by category
        categories = {
            "Core/Lite (Read-only)": ["extract_vba", "list_modules", "analyze_code"],
            "Pro - Injection/Validation": ["inject_vba", "validate_vba_code", "list_macros"],
            "Pro - Office Automation": ["open_in_office", "run_macro", "get_worksheet_data",
                                         "set_worksheet_data", "close_office_file", "list_open_files"],
            "Pro - Excel Tables (NEW v0.3.0)": ["list_tables", "insert_rows", "delete_rows",
                                                 "insert_columns", "delete_columns", "create_table"],
            "Pro - Backup/Refactor": ["backup", "refactor"]
        }

        tool_names = [t.name for t in tools]
        all_expected = []

        for category, expected in categories.items():
            all_expected.extend(expected)
            print(f"{category}:")
            found = [name for name in expected if name in tool_names]
            missing = [name for name in expected if name not in tool_names]

            for name in found:
                print(f"  ✓ {name}")
            for name in missing:
                print(f"  ✗ {name} (MISSING)")
            print()

        # Check for unexpected tools
        unexpected = set(tool_names) - set(all_expected)
        if unexpected:
            print("UNEXPECTED TOOLS:")
            for name in unexpected:
                print(f"  ? {name}")
            print()

        print("=" * 60)

        if len(tools) == EXPECTED_TOOL_COUNT:
            print(f"✓ SUCCESS: Server has exactly {EXPECTED_TOOL_COUNT} tools as expected!")
        else:
            print(f"⚠ WARNING: Expected {EXPECTED_TOOL_COUNT} tools, found {len(tools)}")

        print("\nNext steps:")
        print("1. Restart Claude Desktop to load the updated server")
        print("2. Test with prompts from QUICK_TEST_PROMPTS.md")
        print("3. Try the new Excel Tables features!")

        return True

    except Exception as e:
        print(f"\n[ERROR] Failed to load server: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    success = asyncio.run(test_list_tools())
    sys.exit(0 if success else 1)
