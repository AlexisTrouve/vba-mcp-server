#!/usr/bin/env python3
"""
Simple demonstration of VBA Validation Features

This script directly tests the validation functions without full imports.
"""

import sys
import os

# Add packages to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'packages/pro/src'))

# Import just the module file directly
import importlib.util
spec = importlib.util.spec_from_file_location(
    "inject",
    "packages/pro/src/vba_mcp_pro/tools/inject.py"
)
inject_module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(inject_module)

# Get the functions
_detect_non_ascii = inject_module._detect_non_ascii
_suggest_ascii_replacement = inject_module._suggest_ascii_replacement

def print_header(text):
    """Print formatted header."""
    print(f"\n{'=' * 70}")
    print(f"  {text}")
    print('=' * 70)

def main():
    """Run demonstration tests."""
    print_header("VBA VALIDATION FEATURE DEMONSTRATION")

    # Test 1: Valid ASCII code
    print_header("TEST 1: Valid Code (ASCII only)")
    valid_code = '''Sub HelloWorld()
    MsgBox "Hello World"
End Sub'''
    print(f"Code:\n{valid_code}")
    has_non_ascii, error = _detect_non_ascii(valid_code)
    print(f"\nResult: {'❌ FAILED (unexpected)' if has_non_ascii else '✅ PASSED'}")

    # Test 2: Invalid code with checkmark
    print_header("TEST 2: Invalid Code (contains ✓)")
    invalid_code = '''Sub Test()
    MsgBox "✓ Success"
End Sub'''
    print(f"Code:\n{invalid_code}")
    has_non_ascii, error = _detect_non_ascii(invalid_code)
    print(f"\nResult: {'✅ DETECTED' if has_non_ascii else '❌ MISSED'}")
    if error:
        print(f"\n{error}")

    # Test 3: Replacement suggestions
    print_header("TEST 3: ASCII Replacement Suggestions")
    code_with_unicode = 'MsgBox "✓ Task → Next • Item"'
    print(f"Original code: {code_with_unicode}")

    suggested, changes = _suggest_ascii_replacement(code_with_unicode)
    print(f"\nSuggested code: {suggested}")
    print(f"\n{changes}")

    # Test 4: Math symbols
    print_header("TEST 4: Math Symbol Replacements")
    math_code = 'If x ≤ 10 And y ≥ 20 And z ≠ 0 Then'
    print(f"Original: {math_code}")

    suggested, changes = _suggest_ascii_replacement(math_code)
    print(f"Suggested: {suggested}")
    print(f"\n{changes}")

    # Test 5: Complete workflow
    print_header("TEST 5: Complete Validation Workflow")
    print("\nScenario: Developer tries to inject code with Unicode characters")

    bad_code = '''Sub Report()
    Debug.Print "✓ Passed"
    Debug.Print "✗ Failed"
End Sub'''

    print(f"\nOriginal code:\n{bad_code}")

    # Step 1: Detection
    print("\n[STEP 1] Pre-validation check...")
    has_non_ascii, error = _detect_non_ascii(bad_code)

    if has_non_ascii:
        print("❌ VALIDATION FAILED - Code rejected before injection")
        print("\nError message:")
        print(error)

        # Step 2: Suggestions
        print("\n[STEP 2] Providing suggestions...")
        suggested, changes = _suggest_ascii_replacement(bad_code)
        print(changes)
        print(f"\nCorrected code:\n{suggested}")

        # Step 3: Verify
        print("\n[STEP 3] Validating corrected code...")
        has_non_ascii_2, _ = _detect_non_ascii(suggested)
        print(f"Result: {'✅ VALID' if not has_non_ascii_2 else '❌ STILL INVALID'}")

    # Summary
    print_header("SUMMARY")
    print("""
✅ IMPLEMENTATION COMPLETE

Three helper functions added to inject.py:

1. _detect_non_ascii(code: str) -> Tuple[bool, str]
   - Detects any character with ord() > 127
   - Returns line numbers of first occurrence
   - Provides helpful error messages

2. _suggest_ascii_replacement(code: str) -> Tuple[str, str]
   - Suggests ASCII replacements for common Unicode chars
   - Examples: ✓ → [OK], ✗ → [ERROR], → → ->
   - Returns both suggested code and change description

3. _compile_vba_module(vb_module) -> Tuple[bool, Optional[str]]
   - Validates VBA code after injection
   - Attempts to force VBA parsing by accessing properties
   - Returns validation status and error message

Modified inject_vba_tool():
   - PRE-VALIDATION: Checks for non-ASCII before injection
   - POST-VALIDATION: Validates code after injection
   - ROLLBACK: Restores old code or deletes module if validation fails
   - IMPROVED ERRORS: Clear messages with suggestions

Modified _inject_vba_windows():
   - Stores old code before modification
   - Validates after injection
   - Rolls back on validation failure
   - Returns validation status

IMPACT:
   - Prevents VBA syntax errors
   - Prevents non-ASCII character issues
   - Automatic rollback prevents file corruption
   - Clear, actionable error messages
   - CRITICAL FIX (P0) - prevents data loss
    """)

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n❌ Error: {e}")
        import traceback
        traceback.print_exc()
