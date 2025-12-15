# VBA MCP Server Pro

Professional MCP server for VBA extraction, analysis, and modification.

**Version 0.4.0** - Production Ready 🚀

## What's New in v0.4.0

**Major Reliability Improvements** - Three critical problems resolved:

✅ **Verified Injections** - Post-save verification ensures modules actually exist after injection (no more false positives)

✅ **Automated Macro Execution** - Injected macros now execute without manual security prompts via `AutomationSecurityContext`

✅ **No File Corruption** - Session-based injection eliminates corruption from multiple injections

**Score:** 6/10 → **9/10 (Production Ready)**

[See detailed changelog](CHANGELOG.md#040---2025-12-15)

## Features

### Lite Features (included)
- Extract VBA code from Office files
- List VBA modules
- Analyze code structure and complexity

### Pro Features
- **Inject VBA** - Modify and inject VBA code back into Office files (with validation, verification & rollback)
- **Run Macros** - Execute VBA macros programmatically with automatic security handling
- **Refactor** - AI-powered refactoring suggestions
- **Backup Management** - Create, restore, and manage backups
- **Office Automation** - Open files, read/write data interactively
- **Excel Tables** - Full support for Excel Tables (ListObjects) with row/column operations
- **VBA Validation** - Validate VBA code before injection, detect syntax errors
- **Session Management** - Efficient file session reuse with automatic cleanup

## Installation

```bash
pip install vba-mcp-server-pro
```

On Windows, for full VBA injection support:
```bash
pip install vba-mcp-server-pro[windows]
```

## Platform Requirements

### Windows (Native) - RECOMMENDED ✅
- Full support for all features
- VBA injection, validation, and execution
- No limitations

### Windows (WSL) - LIMITED SUPPORT ⚠️
- **Known Issue:** `Excel.Application.Visible` property cannot be set in WSL
- **Impact:** Excel windows may be visible during automation (Visible=False fails)
- **Workaround:** VBA MCP gracefully handles this - operations continue despite error
- **Recommendation:** For best experience, run on native Windows Python

**WSL Setup:**
```bash
# Option 1: Run Windows Python from WSL
/mnt/c/Python311/python.exe -m vba_mcp_pro.server

# Option 2: Install on Windows and connect from WSL
# (See KNOWN_ISSUES.md for full WSL guide)
```

### macOS / Linux - NOT SUPPORTED ❌
- pywin32 and COM automation are Windows-only
- Use Lite package for read-only VBA extraction

## Configuration

```json
{
  "mcpServers": {
    "vba-pro": {
      "command": "vba-mcp-server-pro"
    }
  }
}
```

## Usage

### Excel Table Operations (NEW in v0.3.0)
"List all tables in sales.xlsx"
"Insert 5 rows at position 10 in the Sales table"
"Delete columns ['TempData', 'Notes'] from Analysis table"
"Convert range A1:D100 to a table named SalesData"
"Get data from columns ['Name', 'Revenue'] in Customers table"

### Office Automation
"Open budget.xlsm in Excel"
"Run the GenerateReport macro in budget.xlsm"
"Run the Calculate macro with enable_macros set to true"  # NEW: automatic security handling
"Get data from Sheet1, range A1:D10"
"Write [[100, 200], [300, 400]] to Sheet2 starting at B5"
"Close budget.xlsm and save changes"

### VBA Validation & Injection (with Verification)
"Validate this VBA code before injecting: Sub Test()..."
"List all macros in budget.xlsm"
"Update the CalculateTotal function in budget.xlsm with this new code..."  # Auto-verified

**Note:** All injections now include post-save verification. Output includes `"verified": True` when successful.

### Refactoring
"Suggest refactoring for my Excel macros"

### Backup Management
"Create a backup of budget.xlsm before I modify it"
"Restore budget.xlsm from yesterday's backup"

## License

Commercial License - Contact alexistrouve.pro@gmail.com for licensing.

## Support

Email: alexistrouve.pro@gmail.com
