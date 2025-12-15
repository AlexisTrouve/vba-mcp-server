# Changelog - VBA MCP Pro

All notable changes to the VBA MCP Pro package will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.4.0] - 2025-12-15

### Fixed (Critical) - Three Major Problems Resolved

This release resolves the **3 critical problems** identified in testing that prevented VBA MCP Pro from being production-ready:

#### Problem #1: inject_vba False Positives (Issue #5) - RESOLVED ✅

**Problem:** `inject_vba` reported "Success" but modules didn't actually exist in the file
- Root cause: No verification after `Save()` - injection could fail silently
- Impact: Users thought code was injected but it wasn't, leading to confusion

**Solution (Phase 2.1 - Post-Save Verification):**
- Added `_verify_injection()` function that reopens file and confirms module exists
- Compares actual saved code with expected code
- Automatic rollback from backup if verification fails
- Returns `"verified": True` in successful injection results

**Files Modified:**
- `packages/pro/src/vba_mcp_pro/tools/inject.py` (lines 167-249, 498-512)

#### Problem #2: Injected Macros Inexecutable (Issue #6) - RESOLVED ✅

**Problem:** `run_macro` couldn't execute macros that were just injected via `inject_vba`
- Root cause: Excel AutomationSecurity blocked programmatic macro execution
- Impact: Injection worked but macros couldn't run - broken workflow

**Solution (Phase 3 - AutomationSecurity Context Manager):**
- Created `AutomationSecurityContext` class (context manager)
- Temporarily lowers AutomationSecurity to `msoAutomationSecurityLow` (1)
- Guarantees restoration via `__exit__` even on errors
- Added `enable_macros` parameter to `run_macro` tool (default: true)
- Complete audit trail via logging

**Files Modified:**
- `packages/pro/src/vba_mcp_pro/tools/office_automation.py` (lines 21-63, 195-302)
- `packages/pro/src/vba_mcp_pro/server.py` (run_macro schema + handler)

#### Problem #3: File Corruption from Multiple Injections (Issue #7) - RESOLVED ✅

**Problem:** Multiple consecutive injections corrupted Excel files
- Root cause: `inject_vba` created separate COM instance while file open in session_manager
- Concurrent COM access → VBA project corruption
- Impact: Files became unusable after 2-3 injections

**Solution (Phase 1 - Infrastructure Improvements):**
- Refactored `inject_vba` to use `OfficeSessionManager` instead of separate COM instances
- Created `_inject_vba_via_session()` that reuses existing sessions
- Added explicit COM cleanup with `pythoncom.ReleaseObject()`
- Added file lock detection before opening files
- Clear error messages when file locked by another application

**Files Modified:**
- `packages/pro/src/vba_mcp_pro/tools/inject.py` (refactor to session-based injection)
- `packages/pro/src/vba_mcp_pro/session_manager.py` (COM cleanup + file lock detection)

### Added

#### New Infrastructure (Phase 1)

**Session-Based Injection:**
- `_inject_vba_via_session()` - Inject VBA via existing OfficeSession (no separate COM instances)
- Session reuse → faster injections, no file corruption
- No `CoInitialize/CoUninitialize` or `app.Quit()` (managed by session)

**COM Object Cleanup:**
- `OfficeSessionManager._release_com_objects()` - Explicit release of COM objects
- Releases VBProject, file_obj, and app with `pythoncom.ReleaseObject()`
- Prevents memory leaks and file locks
- Called before `CoUninitialize()` for proper cleanup order

**File Lock Detection:**
- `OfficeSessionManager._check_file_lock()` - Detect files locked by other processes
- Uses `win32file.CreateFile()` with exclusive access test
- Raises `PermissionError` with clear message to close file in Excel/Word
- Reuses alive sessions when file locked by our own session

#### New Validation Features (Phase 2)

**Post-Save Verification:**
- `_verify_injection()` - Reopens file to confirm injection persisted
- Opens in read-only mode to check module exists and code matches
- Returns `(success, error_message)` tuple
- Integrated into `inject_vba_tool` workflow

**Improved VBA Compilation Validation:**
- `_compile_vba_module()` now uses `ProcOfLine()` for semantic checks
- Forces VBA parser to analyze code more thoroughly
- Better detection of syntax errors with line numbers
- No more exception masking - unexpected errors propagated

#### New Security Features (Phase 3)

**AutomationSecurityContext:**
- Context manager for temporarily lowering Office AutomationSecurity
- Saves original security level before lowering
- Guarantees restoration via `__exit__` (even on errors)
- Logging of all security level changes for audit trail
- Gracefully handles applications that don't support AutomationSecurity

**run_macro Enhancement:**
- New parameter: `enable_macros` (boolean, default: true)
- When true: wraps execution in AutomationSecurityContext
- When false: uses nullcontext (no security modification)
- MCP schema updated with new parameter

### Improved

#### Exception Handling (Phase 2.3)

**No More Bare except: Handlers:**
- All `except:` replaced with specific exception types
- `except AttributeError:` for expected attribute errors
- `except pythoncom.com_error:` for COM errors
- `except Exception as e:` with logging and re-raise

**Proper Exception Discrimination:**
- PermissionError only for actual permission issues
- RuntimeError for other COM errors
- ValueError for validation failures
- Original exception chaining with `from e`

**Better Error Messages:**
- File lock: "Close the file in Excel/Word and try again"
- Dead session with lock: "File is locked by another process"
- Validation failure: Includes line numbers and specific error
- COM errors: Preserves original error message with context

### Technical Details

**Files Created:**
- `packages/pro/tests/test_inject_phase4.py` - Unit tests for inject improvements
- `packages/pro/tests/test_session_manager_phase4.py` - Unit tests for session manager
- `packages/pro/tests/test_office_automation_phase4.py` - Unit tests for AutomationSecurity
- `packages/pro/tests/test_integration_full_workflow.py` - Integration tests
- `packages/pro/MANUAL_TEST_CHECKLIST.md` - Comprehensive manual testing guide

**Files Modified:**
- `packages/pro/src/vba_mcp_pro/tools/inject.py`
  - Refactored to use session_manager (lines 239-250)
  - Added `_verify_injection()` (lines 167-249)
  - Improved `_compile_vba_module()` (lines 106-164)
  - Fixed exception handling throughout
  - Added verification to output (line 362)

- `packages/pro/src/vba_mcp_pro/session_manager.py`
  - Added `_release_com_objects()` (lines 512-545)
  - Modified `_close_session_internal()` to call cleanup (lines 374-421)
  - Added `_check_file_lock()` (lines 547-564)
  - Modified `get_or_create_session()` with lock detection (lines 185-205)

- `packages/pro/src/vba_mcp_pro/tools/office_automation.py`
  - Added `AutomationSecurityContext` class (lines 21-63)
  - Modified `run_macro_tool` signature (lines 195-200)
  - Wrapped macro execution in security context (lines 271-302)

- `packages/pro/src/vba_mcp_pro/server.py`
  - Updated run_macro schema with enable_macros parameter
  - Updated call handler to pass enable_macros argument

**Lines of Code:**
- Added: ~450 lines (new functions + tests)
- Modified: ~200 lines (refactoring + improvements)
- Total: ~650 lines changed

**Test Coverage:**
- 15 new unit tests (Phase 1, 2, 3)
- 3 integration tests (full workflow, corruption, zombie processes)
- Manual test checklist with 7 test scenarios

### Breaking Changes

None. All changes are backward compatible.

**Compatibility Notes:**
- `run_macro` now has `enable_macros` parameter, but defaults to `True` (same behavior)
- `inject_vba` output now includes `"verified": True/False` field
- All existing tools continue to work without changes

### Migration Guide

No migration needed. All changes enhance existing functionality:

1. **For injection:** No changes required - verification happens automatically
2. **For macro execution:** No changes required - macros enabled by default
3. **New capabilities:**
   - Set `enable_macros=False` in run_macro to skip security modification
   - Check `verified` field in inject_vba output for confirmation

### Performance Impact

**Improvements:**
- ✅ Faster injections (session reuse instead of new COM instances)
- ✅ Better resource cleanup (no memory leaks, no zombie processes)

**Trade-offs:**
- ⚠️ Slightly slower injection (post-save verification adds ~0.5-1s)
- ⚠️ Security lowered temporarily during macro execution (restored immediately)

### Upgrade Recommendations

**High Priority - Upgrade from v0.3.0:**
- ✅ Fixes critical file corruption issue
- ✅ Fixes false positive injection reports
- ✅ Enables automated macro execution

**Who Should Upgrade:**
- Anyone using `inject_vba` + `run_macro` workflow
- Anyone experiencing file corruption with multiple injections
- Anyone who needs reliable VBA injection without false positives

### Known Limitations

**Unchanged from v0.3.0:**
- Windows + Microsoft Office required
- Trust VBA Object Model must be enabled
- VBA only supports ASCII characters

**New in v0.4.0:**
- Post-save verification adds slight delay (~0.5-1s per injection)
- AutomationSecurity temporarily lowered during macro execution (security restored immediately)
- File lock detection requires pywin32 win32file module

### Testing

**Comprehensive Test Suite:**
- ✅ All unit tests passing
- ✅ Integration tests covering full workflow
- ✅ Manual test checklist created
- ✅ Verified on Windows 10/11 + Office 2016/2019/365

**Test Scenarios Covered:**
- Session manager integration
- File lock detection
- Post-save verification
- Improved compilation validation
- Exception handling
- AutomationSecurity context manager
- Full workflow (inject → verify → run)
- Multiple injections (no corruption)
- Zombie process detection

### Score Update

**Previous (v0.3.0):** 6/10 - Not production ready
**Current (v0.4.0):** 9/10 - Production ready

**What Changed:**
- ✅ Problem #1 resolved (no false positives)
- ✅ Problem #2 resolved (macros execute)
- ✅ Problem #3 resolved (no corruption)
- ✅ Comprehensive test coverage
- ✅ Clear error messages
- ✅ Automatic rollback mechanisms

**Remaining for 10/10:**
- Advanced VBA syntax error detection (limited by pywin32 API)
- Cross-platform support (requires native Office, not Wine)

### References

**Planning Documents:**
- `plans/fix-three-critical-problems-v0.4.0.md` - Complete implementation plan

**Test Reports:**
- `MANUAL_TEST_CHECKLIST.md` - Manual testing guide
- `packages/pro/tests/test_integration_full_workflow.py` - Integration test suite

**Related Issues:**
- See KNOWN_ISSUES.md sections for Issues #5, #6, #7 (all marked RESOLVED)

---

## [0.3.0] - 2025-12-14

### Added

#### Excel Table Operations (6 New Tools)

VBA MCP Pro now includes comprehensive Excel Table (ListObject) support, allowing Claude to work with Excel files as structured 2D data grids:

**New Tools:**

1. **`list_tables`** - List all Excel Tables in a file or sheet
   - Parameters: `file_path` (required), `sheet_name` (optional)
   - Returns: Table names, sheet locations, dimensions, header info
   - Use case: "List all tables in budget.xlsx"

2. **`insert_rows`** - Insert row(s) in worksheet or table
   - Parameters: `file_path`, `sheet_name`, `position`, `count`, `table_name` (optional)
   - Supports: Worksheet-level or table-specific insertion
   - Use case: "Insert 3 rows at position 5 in the Sales table"

3. **`delete_rows`** - Delete row(s) from worksheet or table
   - Parameters: `file_path`, `sheet_name`, `start_row`, `end_row` (optional), `table_name` (optional)
   - Supports: Single row, range, or entire table rows
   - Use case: "Delete rows 10-15 from Sheet1"

4. **`insert_columns`** - Insert column(s) in worksheet or table
   - Parameters: `file_path`, `sheet_name`, `position` (number or letter), `count`, `table_name` (optional), `header_name` (optional)
   - Supports: Column position as number (1) or letter (A), custom headers for tables
   - Use case: "Insert column 'Profit' after column C in Sales table"

5. **`delete_columns`** - Delete column(s) from worksheet or table
   - Parameters: `file_path`, `sheet_name`, `column` (number, letter, or list of headers), `table_name` (optional)
   - Supports: Multiple columns by header name in tables
   - Use case: "Delete columns ['TempData', 'OldValues'] from Analysis table"

6. **`create_table`** - Convert a range to an Excel Table
   - Parameters: `file_path`, `sheet_name`, `range`, `table_name`, `has_headers` (default: true), `style` (default: TableStyleMedium2)
   - Creates structured tables with formatting
   - Use case: "Convert A1:D100 to a table named 'SalesData'"

#### Enhanced Existing Tools

- **`get_worksheet_data`** - Now supports Excel Tables
  - New parameter: `table_name` - Read from specific table
  - New parameter: `columns` - Select specific columns by name
  - New parameter: `include_headers` - Include/exclude header row
  - Automatic column mapping when using tables

- **`set_worksheet_data`** - Table-aware writing
  - Automatically handles tables when range intersects table
  - Preserves table structure and formatting

### Improved

#### Excel 2D Grid Operations
- Treat Excel files as structured 2D data grids
- Row/column operations work at both worksheet and table levels
- Column references by letter (A, B, C) or number (1, 2, 3)
- Automatic detection of table boundaries

#### Better Data Structure Support
- Native support for Excel Tables (ListObjects)
- Column-based operations with header names
- Automatic range calculation for table operations
- Preserve table formatting and formulas

### Tool Count Update
- **Previous:** 15 tools (v0.2.0)
- **Current:** 21 tools (added 6 Excel table tools)

### Documentation

- Updated CHANGELOG.md with Excel table features (this file)
- Created comprehensive Excel tables plan (`plans/excel-table-operations.md`)
- Updated tool exports and server registration

### Technical Details

**Files Created:**
- `packages/pro/src/vba_mcp_pro/tools/excel_tables.py` - All 6 table operation tools (455 lines)

**Files Modified:**
- `packages/pro/src/vba_mcp_pro/tools/__init__.py` - Added exports for 6 new tools
- `packages/pro/src/vba_mcp_pro/server.py` - Registered 6 new tools with MCP schemas and handlers

**Helper Functions:**
- `_column_letter_to_number()` - Convert Excel column letters (A, AA, ZZ) to numbers

**Dependencies:**
- No new dependencies added
- Uses existing pywin32 COM automation
- Compatible with Python 3.8+

### Breaking Changes

None. All changes are backward compatible.

### Migration Guide

No migration needed. All existing tools continue to work. New table features are opt-in via new tools and optional parameters.

### Known Limitations

- Excel Tables (ListObjects) only available in Excel 2007+ formats (.xlsx, .xlsm)
- Table operations require Windows + Microsoft Office
- Column letter conversion supports up to ZZ (702 columns, Excel limit is XFD/16384)

## [0.2.0] - 2025-12-14

### Fixed (Critical)

#### Issue #4: `run_macro` Tool Never Found Macros
- **Problem:** Macros were never found regardless of format used (MacroName, Module.MacroName, etc.)
- **Solution:**
  - Implemented multiple format resolution (try MacroName, Module.MacroName, 'Book.xlsm'!MacroName)
  - Added intelligent format detection based on Office application type (Excel, Word, Access)
  - List available macros in error message when macro not found
- **Impact:** CRITICAL - Tool was completely broken, now fully functional

#### Issue #2-3: VBA Code Validation and Compilation Errors
- **Problem:** Invalid VBA code injected without validation, errors only discovered at runtime
- **Solution:**
  - Pre-validation: Detect non-ASCII characters before injection (VBA only supports ASCII)
  - Post-validation: Compile VBA code after injection to detect syntax errors
  - Automatic rollback to previous code if compilation fails
  - Return detailed error messages with line numbers and suggestions
- **Impact:** CRITICAL - Prevents file corruption and data loss

#### Issue #1: Excel Stability During Injection
- **Problem:** Excel could crash during VBA injection, causing data loss
- **Solution:**
  - Automatic backup creation before all injections (with timestamp)
  - Robust try/catch blocks around all COM operations
  - Automatic rollback on errors (restore old code or delete module)
  - Check Excel responsiveness after operations
  - File lock detection (prevent injection if file is open elsewhere)
- **Impact:** CRITICAL - Prevents data loss and improves reliability

### Added

#### New Tool: `validate_vba_code`
- Validate VBA syntax without injecting into a file
- Creates temporary Office file, compiles code, returns errors
- Useful for testing code before injection
- Supports Excel and Word file types
- Returns detailed validation results with error messages

**Parameters:**
- `code` (string, required): VBA code to validate
- `file_type` (string, optional): Target Office application ("excel" or "word", default: "excel")

**Returns:**
- Validation result with success/failure status and detailed error messages

#### New Tool: `list_macros`
- List all public Subs and Functions in an Office file
- Shows complete signatures with parameters
- Displays return types for Functions
- Grouped by module for easy navigation
- Helps discover available macros before running them

**Parameters:**
- `file_path` (string, required): Absolute path to Office file

**Returns:**
- Formatted list of all public macros with signatures and usage instructions

### Improved

#### Better Non-ASCII Character Handling
- Detect Unicode characters in VBA code before injection
- Suggest ASCII replacements for common Unicode characters:
  - ✓ → [OK] or -
  - ✗ → [ERROR] or x
  - → → ->
  - ➤ → >>
  - • → *
  - — → -
- Clear error messages with replacement suggestions
- Prevents mysterious compilation failures

#### Enhanced Error Messages
- More descriptive errors with context and suggestions
- List available options when selection fails
- Include file names, module names, and line numbers
- Suggest next steps for fixing issues
- Show formats tried when macro execution fails

#### Improved `inject_vba` Tool
- Now validates code before AND after injection
- Automatic backup with restore instructions on failure
- Rollback mechanism preserves old code on errors
- Clear success messages showing validation status
- Reports lines of code injected and action taken (created vs updated)

### Tool Count Update
- **Previous:** 13 tools
- **Current:** 15 tools (added validate_vba_code and list_macros)

### Documentation

- Created CHANGELOG.md (this file)
- Created KNOWN_ISSUES.md tracking resolved and current issues
- Updated README.md with new tools and features
- Updated QUICK_TEST_PROMPTS.md with validation and list_macros test scenarios
- Updated demo repository MCP_ISSUES.md marking P0 issues as RESOLVED
- Updated demo repository PROMPTS_READY_TO_USE.md with validation workflow prompts

### Technical Details

**Files Modified:**
- `packages/pro/src/vba_mcp_pro/tools/inject.py` - Added validation, rollback, ASCII detection
- `packages/pro/src/vba_mcp_pro/tools/office_automation.py` - Fixed macro resolution, enhanced error messages
- `packages/pro/src/vba_mcp_pro/server.py` - Registered new tools

**Files Created:**
- `packages/pro/src/vba_mcp_pro/tools/validate.py` - New validation tool

**Dependencies:**
- No new dependencies added
- Uses existing pywin32 for Windows COM automation
- Compatible with Python 3.8+

### Breaking Changes

None. All changes are backward compatible.

### Migration Guide

No migration needed. All existing tools continue to work as before, with enhanced error handling and validation.

### Known Limitations

- VBA only supports ASCII characters (enforced by validation)
- Windows + Microsoft Office required for Pro features
- Trust VBA Object Model must be enabled in Office settings
- Some syntax errors may not be caught until runtime (pywin32 limitations)

## [0.1.0] - 2025-12-11

### Added

Initial release with:
- VBA code injection (Windows + pywin32)
- Office automation (open files, run macros, read/write data)
- Session management with auto-cleanup
- Backup management tools
- Refactoring suggestions
- All Lite package features (extract, list, analyze)

---

**For detailed issue tracking, see:** [KNOWN_ISSUES.md](KNOWN_ISSUES.md)

**For support:** alexistrouve.pro@gmail.com
