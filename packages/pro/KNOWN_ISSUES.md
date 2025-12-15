# Known Issues - VBA MCP Pro

This document tracks resolved and current issues in the VBA MCP Pro package.

---

## Resolved Issues

### Issue #1: Excel Crashes During VBA Injection
**Status:** RESOLVED in v0.2.0 (2025-12-14)
**Priority:** P0 - CRITICAL
**Resolution Date:** 2025-12-14

**Original Problem:**
- Excel could crash during VBA code injection
- No graceful error handling
- Potential data loss when crashes occurred

**Solution Implemented:**
- Automatic backup creation before all injections (with timestamp)
- Robust try/catch blocks around all COM operations
- Automatic rollback on errors (restore old code or delete new module)
- Excel responsiveness check after operations
- File lock detection (prevents injection if file is open elsewhere)

**Verification:**
- Tested with multiple injection scenarios
- Tested with locked files
- Tested with invalid code (triggers rollback)
- No crashes observed in testing

---

### Issue #2: No VBA Code Validation Before Injection
**Status:** RESOLVED in v0.2.0 (2025-12-14)
**Priority:** P0 - CRITICAL
**Resolution Date:** 2025-12-14

**Original Problem:**
- `inject_vba` tool accepted any code without validation
- No syntax checking before injection
- Errors only discovered when manually running macros
- Common issue: Unicode characters in code (e.g., ✓, →, •) not supported by VBA

**Solution Implemented:**
- Pre-validation: Detect non-ASCII characters before injection
- Post-validation: Compile VBA code after injection
- Automatic rollback if compilation fails
- Clear error messages with ASCII replacement suggestions
- New `validate_vba_code` tool for standalone validation

**Verification:**
- Test with Unicode characters: REJECTED with helpful suggestions
- Test with syntax errors: REJECTED after compilation, code rolled back
- Test with valid code: ACCEPTED and compiled successfully

---

### Issue #3: No Compilation Error Reporting
**Status:** RESOLVED in v0.2.0 (2025-12-14)
**Priority:** P0 - CRITICAL
**Resolution Date:** 2025-12-14

**Original Problem:**
- When invalid VBA code was injected, tool reported "Success"
- No indication of syntax errors
- Users discovered problems only when trying to run macros

**Solution Implemented:**
- Implemented `_compile_vba_module()` function
- Detects syntax errors after injection
- Returns detailed error messages with context
- Automatic rollback on compilation failure
- Success message now includes "Validation: Passed ✓" status

**Verification:**
- Test with missing End If: Returns compilation error, rolls back
- Test with invalid variable names: Returns error, rolls back
- Test with valid code: Returns success with validation status

---

### Issue #4: `run_macro` Tool Never Finds Macros
**Status:** RESOLVED in v0.2.0 (2025-12-14)
**Priority:** P0 - BLOQUANT
**Resolution Date:** 2025-12-14

**Original Problem:**
- `run_macro` always returned "Macro not found"
- Even for macros that existed and were visible in Excel
- Tried formats: `MacroName`, `Module.MacroName` - none worked

**Solution Implemented:**
- Try multiple macro name formats automatically:
  - For Excel: `MacroName`, `Module.MacroName`, `'Book.xlsm'!MacroName`, `'Book.xlsm'!Module.MacroName`
  - For Word/Access: `MacroName`, `Module.MacroName`
- Intelligent format selection based on application type
- List available macros in error message when not found
- New `list_macros` tool to discover available macros
- Helper function `_list_available_macros()` for debugging

**Verification:**
- Test with simple macro name: WORKS
- Test with module.macro format: WORKS
- Test with non-existent macro: Lists available macros in error
- Test with macros with parameters: WORKS

---

### Issue #5: inject_vba False Positives
**Status:** RESOLVED in v0.4.0 (2025-12-15)
**Priority:** P0 - CRITICAL
**Resolution Date:** 2025-12-15

**Original Problem:**
- `inject_vba` reported "VBA Injection Successful" but modules didn't actually exist in the file
- No verification after `Save()` - injection could fail silently
- Users thought code was injected but it wasn't, leading to confusion and errors
- Common scenario: Save() succeeds but VBA project doesn't persist the module

**Solution Implemented:**
- Added `_verify_injection()` function (inject.py lines 167-249)
- Reopens file in read-only mode after injection
- Confirms module exists and code matches expected code
- Automatic rollback from backup if verification fails
- Returns `"verified": True` in successful injection results
- Integrated into `inject_vba_tool` workflow (inject.py lines 498-512)

**Technical Details:**
- Opens separate COM instance to avoid session conflicts
- Compares actual saved code with expected code (strips whitespace)
- Returns `(success, error_message)` tuple
- Triggers restore from backup on verification failure
- Adds ~0.5-1s to injection time (acceptable trade-off)

**Verification:**
- Test with valid code: Module exists after verification ✅
- Test with verification failure: File restored from backup ✅
- Test output includes "Verified: Yes" ✅
- Integration test confirms module persists after save ✅

---

### Issue #6: Injected Macros Inexecutable
**Status:** RESOLVED in v0.4.0 (2025-12-15)
**Priority:** P0 - CRITICAL
**Resolution Date:** 2025-12-15

**Original Problem:**
- `run_macro` couldn't execute macros that were just injected via `inject_vba`
- Excel AutomationSecurity setting blocked programmatic macro execution
- Macros found by `list_macros` but execution failed with "cannot run the macro"
- Broken workflow: injection worked but macros couldn't run
- Users had to manually enable macros or lower security settings

**Solution Implemented:**
- Created `AutomationSecurityContext` class (office_automation.py lines 21-63)
- Context manager temporarily lowers AutomationSecurity to `msoAutomationSecurityLow` (1)
- Saves original security level before lowering
- Guarantees restoration via `__exit__` even on errors
- Added `enable_macros` parameter to `run_macro` tool (default: true)
- Complete audit trail via logging for security compliance

**Technical Details:**
- Context manager saves original AutomationSecurity level
- Lowers to level 1 (macros enabled) during macro execution
- Restores original level in `__exit__` (guaranteed by Python)
- Gracefully handles applications that don't support AutomationSecurity
- Logging: "AutomationSecurity: 2 → 1" and "AutomationSecurity restored: 2"
- Can be disabled with `enable_macros=False` parameter

**Security Notes:**
- Security is lowered only during macro execution (< 1 second)
- Immediately restored after execution or on errors
- All changes logged for audit trail
- Backup system still protects against malicious macros
- User has control via `enable_macros` parameter

**Verification:**
- Test with injected macro: Executes without manual intervention ✅
- Test with error during execution: Security still restored ✅
- Test with `enable_macros=False`: No security modification ✅
- Integration test: Full inject → run workflow ✅
- Logs show security lowering and restoration ✅

---

### Issue #7: File Corruption from Multiple Injections
**Status:** RESOLVED in v0.4.0 (2025-12-15)
**Priority:** P0 - CRITICAL
**Resolution Date:** 2025-12-15

**Original Problem:**
- Multiple consecutive injections corrupted Excel files
- VBA project became unreadable after 2-3 injections
- Files couldn't be opened or showed "VBA project corrupted" errors
- Root cause: `inject_vba` created separate COM instance while file was already open in `session_manager`
- Concurrent COM access to same file → VBA project corruption
- File locks caused by zombie COM objects

**Solution Implemented (Phase 1 - Infrastructure):**

**1. Session-Based Injection:**
- Refactored `inject_vba` to use `OfficeSessionManager` (inject.py lines 239-250)
- Created `_inject_vba_via_session()` function (inject.py lines 278-412)
- No more separate COM instances - reuses existing sessions
- No `CoInitialize/CoUninitialize` or `app.Quit()` (managed by session)
- Session reuse → faster injections, no concurrent access

**2. COM Object Cleanup:**
- Added `_release_com_objects()` to OfficeSessionManager (session_manager.py lines 512-545)
- Explicitly releases VBProject, file_obj, and app with `pythoncom.ReleaseObject()`
- Called before `CoUninitialize()` for proper cleanup order
- Prevents memory leaks and file locks
- Eliminates zombie Excel processes

**3. File Lock Detection:**
- Added `_check_file_lock()` method (session_manager.py lines 547-564)
- Uses `win32file.CreateFile()` with exclusive access test
- Detects files locked by other processes or applications
- Raises `PermissionError` with clear message: "Close the file in Excel/Word and try again"
- Reuses alive sessions when file locked by our own session
- Modified `get_or_create_session()` with lock detection (session_manager.py lines 185-205)

**Technical Details:**
- Unified all file access through session_manager
- Eliminated concurrent COM instance creation
- Proper COM object lifetime management
- File lock detection prevents opening locked files
- Clear error messages for lock scenarios

**Verification:**
- Test with 2 consecutive injections: No corruption ✅
- Test with 10 consecutive injections: All modules exist, file opens ✅
- Test with file locked in Excel: Clear error message ✅
- Test with multiple sessions: No conflicts ✅
- Integration test: 10 injections, no corruption ✅
- Task Manager: No zombie Excel processes ✅

---

## Current Issues

**None currently known.**

All P0 (critical) issues have been resolved. Please report new issues at:
- Email: alexistrouve.pro@gmail.com
- GitHub Issues: [Will be added when repository is public]

---

## Current Limitations

These are not bugs, but inherent limitations of VBA and the Windows COM automation approach.

### 1. VBA Only Supports ASCII Characters

**Description:**
VBA does not support Unicode characters in code (only in string literals).

**Impact:**
Code with Unicode characters (✓, →, •, etc.) will fail to compile.

**Mitigation:**
- Server now detects non-ASCII characters before injection
- Provides helpful replacement suggestions (✓ → [OK], → → ->, etc.)
- Rejects code with clear error message

**Example:**
```vba
' ❌ This will be rejected:
MsgBox "✓ Success"

' ✅ This will work:
MsgBox "[OK] Success"
```

---

### 2. Windows + Microsoft Office Required

**Description:**
All Pro features (injection, automation) require:
- Windows operating system
- Microsoft Office installed (Excel, Word, or Access)
- Trust VBA Object Model enabled in Office settings

**Impact:**
Pro features do not work on macOS or Linux.

**Workaround:**
- Use Lite package for cross-platform read-only VBA extraction
- Run Pro package in Windows VM or Remote Desktop if needed

---

### 3. VBA Compilation Detection Limitations

**Description:**
VBA compilation detection is limited by pywin32 capabilities.

**Impact:**
Some syntax errors may not be caught until runtime:
- Logic errors (will compile but fail when run)
- Runtime errors (e.g., division by zero, array bounds)
- Reference errors (missing libraries)

**Mitigation:**
- Server catches basic syntax errors
- Users should still test macros manually after injection
- Use `validate_vba_code` tool before injection for pre-flight checks

---

### 4. Macro Execution Context Issues

**Description:**
Some macros may fail when run via automation even if they work when run manually.

**Failing Scenarios:**
- Macros expecting user interaction (InputBox, MsgBox with user input)
- Macros requiring active selection/context
- Macros depending on specific Excel state (selected sheet, active cell)
- Macros using Application.Caller or similar context-dependent functions

**Mitigation:**
- Design macros to accept parameters instead of prompting users
- Pass required context as function arguments
- Avoid macros that depend on UI state

**Example:**
```vba
' ❌ May fail in automation:
Sub ProcessData()
    Dim rng As Range
    Set rng = Application.Selection  ' Depends on active selection
    ' ... process rng
End Sub

' ✅ Works in automation:
Sub ProcessData(rng As Range)
    ' ... process rng
End Sub
' Call: run_macro("ProcessData", [Range("A1:C10")])
```

---

### 5. Session Timeout

**Description:**
Office files auto-close after 1 hour of inactivity.

**Impact:**
Need to reopen files if session expires.

**Mitigation:**
- Session automatically refreshes on each operation
- 1 hour is usually sufficient for typical workflows
- Files can be manually reopened with `open_in_office` tool

---

### 6. File Locking

**Description:**
Cannot inject VBA if file is already open in Office.

**Impact:**
Must close file before injection.

**Mitigation:**
- Server detects locked files and returns clear error
- Use `close_office_file` tool to close file programmatically
- Or close file manually in Office application

---

## Workarounds and Best Practices

### Best Practice #1: Always Use `validate_vba_code` First

**Recommendation:**
Before injecting code, validate it first:

```
1. Use validate_vba_code to check syntax
2. Review any errors and fix them
3. Only then use inject_vba to inject the corrected code
```

**Benefit:**
Catches errors before modifying the file, avoiding rollbacks and failed injections.

---

### Best Practice #2: Check Available Macros Before Running

**Recommendation:**
Use `list_macros` to see what's available:

```
1. Use list_macros to list all public macros
2. Identify the exact name and signature
3. Use run_macro with the correct name/format
```

**Benefit:**
Avoids "macro not found" errors and shows expected parameters.

---

### Best Practice #3: Keep Backups

**Recommendation:**
Even though `inject_vba` creates automatic backups:
- Keep manual backups of important files
- Use version control for VBA code (extract to .bas files)
- Test on copies before modifying production files

**Benefit:**
Extra safety layer for critical files.

---

### Best Practice #4: Design Automation-Friendly Macros

**Recommendation:**
When writing macros to be run via automation:
- Accept parameters instead of using InputBox
- Return values instead of showing MsgBox
- Avoid dependencies on active selection
- Handle errors gracefully (don't rely on user clicking "OK")

**Example:**
```vba
' ✅ Automation-friendly:
Function CalculateTotal(data As Range) As Double
    On Error GoTo ErrorHandler
    ' ... calculation logic
    CalculateTotal = result
    Exit Function
ErrorHandler:
    CalculateTotal = -1  ' Error code
End Function
```

---

## Reporting New Issues

If you discover a new issue:

1. **Check this document** to see if it's a known limitation
2. **Gather information:**
   - VBA MCP Pro version (`pip show vba-mcp-server-pro`)
   - Windows version
   - Office version (Excel 2016, 2019, 365, etc.)
   - Exact error message
   - Steps to reproduce
3. **Report:**
   - Email: alexistrouve.pro@gmail.com
   - Include all information from step 2
   - Attach sample file if possible (redacted if sensitive)

---

**Last Updated:** 2025-12-15
**Version:** 0.4.0
