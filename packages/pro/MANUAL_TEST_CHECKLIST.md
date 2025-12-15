# Manual Test Checklist - VBA MCP Pro v0.4.0

**Purpose:** Verify that fixes for the 3 critical problems work in real-world usage.

**Prerequisites:**
- Windows OS
- Microsoft Office (Excel/Word) installed
- Claude Desktop with VBA MCP Pro configured
- Sample Excel files (.xlsm) ready

---

## 🎯 Critical Problems Being Tested

1. **Problem #1**: inject_vba false positives (says "Success" but module doesn't exist)
2. **Problem #2**: run_macro can't execute injected macros (blocked by Excel security)
3. **Problem #3**: Multiple injections corrupt files

---

## Test 1: Injection via Session Manager

**Goal:** Verify inject_vba uses session_manager (no corruption)

### Steps:

1. Open a test file with Claude:
   ```
   Open test.xlsm in Excel
   ```

2. Inject first module:
   ```
   In test.xlsm, inject this VBA code into module "TestModule1":

   Sub HelloWorld()
       MsgBox "Test 1 OK"
   End Sub
   ```

3. **Verify output shows:**
   - ✅ "VBA Injection Successful"
   - ✅ "Verified: Yes" (new in v0.4.0)
   - ✅ Module: TestModule1
   - ✅ Backup created

4. Inject second module **without closing file**:
   ```
   In test.xlsm, inject this VBA code into module "TestModule2":

   Sub SecondTest()
       MsgBox "Test 2 OK"
   End Sub
   ```

5. **Verify no errors:**
   - ✅ No "file locked" error (would happen in v0.3.0)
   - ✅ Both injections succeed
   - ✅ "Verified: Yes" for both

6. **Open Excel manually and verify:**
   - Press Alt+F11 in Excel
   - Check VBA Project Explorer
   - ✅ TestModule1 exists
   - ✅ TestModule2 exists
   - ✅ Code is present in both

**Expected Result:** ✅ Both modules injected successfully, no file corruption

---

## Test 2: Post-Save Verification

**Goal:** Detect when modules don't persist (Problem #1)

### Steps:

1. Use a fresh Excel file:
   ```
   Create a new test file called verification_test.xlsm
   ```

2. Inject valid code:
   ```
   In verification_test.xlsm, inject this VBA code into module "ValidModule":

   Sub TestValidation()
       Dim x As Long
       x = 42
   End Sub
   ```

3. **Check output:**
   - ✅ "VBA Injection Successful"
   - ✅ **"Verified: Yes"** ← KEY INDICATOR
   - ✅ Backup created

4. **Open in Excel and verify manually:**
   - Alt+F11
   - Check ValidModule exists
   - ✅ Code is present and correct

5. If verification had failed, output would show:
   - ❌ "Injection verification failed: [reason]"
   - ❌ "File restored from backup"

**Expected Result:** ✅ Verification passes, module actually exists

---

## Test 3: Macro Execution with Security

**Goal:** Execute injected macros without manual security prompts (Problem #2)

### Steps:

1. Inject a test macro:
   ```
   In test.xlsm, inject this VBA code into module "ExecutionTest":

   Function Calculate(a As Long, b As Long) As Long
       Calculate = a + b
   End Function
   ```

2. Wait for "Verified: Yes" confirmation

3. Execute the macro via Claude:
   ```
   Run the macro ExecutionTest.Calculate in test.xlsm with arguments 10 and 20
   ```

4. **Check output:**
   - ✅ "Macro Executed Successfully"
   - ✅ "Return value: 30"
   - ✅ NO security warnings or prompts
   - ✅ Executed without manual intervention

5. **Check logs** (Claude Desktop → Help → View Logs):
   - Search for "AutomationSecurity"
   - ✅ Should see: "AutomationSecurity: 2 → 1" (lowered)
   - ✅ Should see: "AutomationSecurity restored: 2" (restored)

**Expected Result:** ✅ Macro executes automatically with enable_macros=true

**Bonus Test:** Disable macro execution:
```
Run ExecutionTest.Calculate in test.xlsm with enable_macros: false
```
- May fail with security error (expected)

---

## Test 4: Multiple Injections (No Corruption)

**Goal:** Verify 10 consecutive injections don't corrupt file (Problem #3)

### Steps:

1. Create fresh test file:
   ```
   Create a new file multi_inject_test.xlsm
   ```

2. Inject 10 modules consecutively:
   ```
   In multi_inject_test.xlsm, inject 10 modules named Module1 through Module10.
   Each should contain:

   For Module1:
   Sub Test1()
       MsgBox "Module 1"
   End Sub

   For Module2:
   Sub Test2()
       MsgBox "Module 2"
   End Sub

   ... and so on through Module10
   ```

3. **After each injection, verify:**
   - ✅ "Verified: Yes"
   - ✅ No corruption errors
   - ✅ No "file locked" errors

4. **Open Excel manually after all 10:**
   - Alt+F11
   - Count modules in VBA Project Explorer
   - ✅ All 10 modules present
   - ✅ File opens without errors
   - ✅ No VBA project corruption warnings

5. **Close and reopen file:**
   - Close Excel completely
   - Reopen multi_inject_test.xlsm
   - ✅ File opens successfully
   - ✅ All modules still present

**Expected Result:** ✅ All 10 modules exist, file is not corrupted

---

## Test 5: Error Handling

**Goal:** Verify clear error messages

### Test 5A: File Lock Detection

1. Open test.xlsm in Excel manually
2. Try to inject via Claude while file is open:
   ```
   In test.xlsm, inject module "LockedTest" with code:
   Sub Test()
   End Sub
   ```

3. **Expected output:**
   - ❌ Error message containing "locked" or "Close the file"
   - ✅ Clear instructions to close file in Excel
   - ✅ No crash or hang

### Test 5B: Invalid VBA Code

1. Inject invalid code:
   ```
   In test.xlsm, inject module "InvalidTest" with:

   Sub Test(
       ' Missing closing paren
   End Sub
   ```

2. **Expected output:**
   - ❌ "Syntax error at line X"
   - ❌ Code NOT injected
   - ✅ File unchanged (no corruption)

### Test 5C: Verification Failure

This is internal - verification catches when Save() succeeds but module doesn't persist.
Should never happen in normal operation, but if it does:

- ❌ "Injection verification failed: [reason]"
- ✅ File restored from backup
- ✅ Clear error message

**Expected Result:** ✅ All errors have clear, actionable messages

---

## Test 6: Compatibility (Optional)

**Goal:** Verify works with different Office file types

### Excel (.xlsm)
Already tested above ✅

### Word (.docm)

1. Create test.docm
2. Inject VBA module:
   ```
   In test.docm, inject module "WordTest":

   Sub TestWord()
       MsgBox "Word VBA OK"
   End Sub
   ```

3. **Verify:**
   - ✅ Injection succeeds
   - ✅ "Verified: Yes"
   - ✅ Module exists in Word VBA

---

## Test 7: Performance & Cleanup

**Goal:** Verify no memory leaks or zombie processes

### Steps:

1. Open Task Manager (Ctrl+Shift+Esc)
2. Note number of Excel.exe processes before testing
3. Perform 5 injections and 5 macro executions
4. Close all sessions via Claude:
   ```
   List all open Office files, then close all of them
   ```
5. Wait 30 seconds
6. Check Task Manager again

**Expected Result:**
- ✅ No more Excel.exe processes than before
- ✅ No zombie/orphan processes
- ✅ Memory usage returns to normal

---

## 🎯 Success Criteria

**Pass all tests = v0.4.0 is production-ready**

### Phase 1 (Infrastructure):
- ✅ Injections use session_manager (no separate COM instances)
- ✅ Multiple injections work without "file locked" errors
- ✅ No zombie processes after operations

### Phase 2 (Validation):
- ✅ "Verified: Yes" appears in output
- ✅ Modules actually exist after injection
- ✅ Invalid code is rejected with clear errors

### Phase 3 (Security):
- ✅ Injected macros execute without manual intervention
- ✅ AutomationSecurity lowered and restored (check logs)
- ✅ enable_macros parameter works

### Overall:
- ✅ Problem #1 RESOLVED (no false positives)
- ✅ Problem #2 RESOLVED (macros execute)
- ✅ Problem #3 RESOLVED (no corruption)

---

## 📝 Reporting Issues

If any test fails:

1. **Capture:**
   - Exact Claude prompts used
   - Complete error messages
   - Screenshots if relevant

2. **Check logs:**
   - Claude Desktop → Help → View Logs
   - Search for "vba-mcp-pro"
   - Copy relevant error stack traces

3. **Document:**
   - Which test failed
   - Expected vs actual behavior
   - Steps to reproduce

4. **Report in:** `KNOWN_ISSUES.md` or GitHub Issues

---

## ✅ Sign-Off

**Tester:** _______________
**Date:** _______________
**Version Tested:** v0.4.0
**Status:** ⬜ PASS / ⬜ FAIL

**Notes:**
_________________________________
_________________________________
_________________________________
