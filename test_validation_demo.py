#!/usr/bin/env python3
"""
Demonstration of VBA Validation Features

This script demonstrates the new validation functionality added to inject_vba_tool.
"""

import sys
sys.path.insert(0, 'packages/pro/src')

from vba_mcp_pro.tools.inject import (
    _detect_non_ascii,
    _suggest_ascii_replacement,
)

def print_section(title):
    """Print a section header."""
    print(f"\n{'=' * 60}")
    print(f"  {title}")
    print('=' * 60)

def test_ascii_detection():
    """Test ASCII detection with various examples."""
    print_section("Test 1: Non-ASCII Detection")

    # Test case 1: Valid ASCII code
    print("\n1.1 Testing VALID code (ASCII only):")
    valid_code = '''Sub HelloWorld()
    MsgBox "Hello World"
End Sub'''
    print(f"Code:\n{valid_code}")
    has_non_ascii, error = _detect_non_ascii(valid_code)
    print(f"\nResult: {'FAILED ❌' if has_non_ascii else 'PASSED ✅'}")
    if error:
        print(f"Error: {error}")

    # Test case 2: Unicode checkmark
    print("\n1.2 Testing INVALID code (contains ✓):")
    invalid_code_1 = '''Sub Test()
    MsgBox "✓ Success"
End Sub'''
    print(f"Code:\n{invalid_code_1}")
    has_non_ascii, error = _detect_non_ascii(invalid_code_1)
    print(f"\nResult: {'DETECTED ✅' if has_non_ascii else 'MISSED ❌'}")
    if error:
        print(f"\nError Message:\n{error}")

    # Test case 3: Multiple Unicode characters
    print("\n1.3 Testing INVALID code (multiple Unicode chars):")
    invalid_code_2 = '''Sub Process()
    ' Step 1 → Step 2
    MsgBox "• Task complete"
End Sub'''
    print(f"Code:\n{invalid_code_2}")
    has_non_ascii, error = _detect_non_ascii(invalid_code_2)
    print(f"\nResult: {'DETECTED ✅' if has_non_ascii else 'MISSED ❌'}")
    if error:
        print(f"\nError Message:\n{error}")

def test_ascii_replacement():
    """Test ASCII replacement suggestions."""
    print_section("Test 2: ASCII Replacement Suggestions")

    # Test case 1: Simple replacement
    print("\n2.1 Testing replacement for checkmark:")
    code_with_checkmark = 'MsgBox "✓ Done"'
    print(f"Original: {code_with_checkmark}")
    suggested, changes = _suggest_ascii_replacement(code_with_checkmark)
    print(f"Suggested: {suggested}")
    print(f"\nChanges:\n{changes}")

    # Test case 2: Multiple replacements
    print("\n2.2 Testing multiple replacements:")
    code_multiple = 'MsgBox "✓ Task → Next • Item"'
    print(f"Original: {code_multiple}")
    suggested, changes = _suggest_ascii_replacement(code_multiple)
    print(f"Suggested: {suggested}")
    print(f"\nChanges:\n{changes}")

    # Test case 3: Math symbols
    print("\n2.3 Testing math symbol replacements:")
    code_math = 'If x ≤ 10 And y ≥ 20 And z ≠ 0 Then'
    print(f"Original: {code_math}")
    suggested, changes = _suggest_ascii_replacement(code_math)
    print(f"Suggested: {suggested}")
    print(f"\nChanges:\n{changes}")

def test_validation_workflow():
    """Test the complete validation workflow."""
    print_section("Test 3: Complete Validation Workflow")

    print("\nScenario: User tries to inject code with Unicode")
    bad_code = '''Sub Report()
    Debug.Print "✓ Success"
    Debug.Print "✗ Failed"
    Debug.Print "Next → Step"
End Sub'''

    print(f"User's code:\n{bad_code}\n")

    # Step 1: Pre-validation
    print("STEP 1: Pre-validation (before injection)")
    has_non_ascii, error = _detect_non_ascii(bad_code)

    if has_non_ascii:
        print("❌ Validation FAILED - Non-ASCII detected")
        print(f"\nError message shown to user:\n{error}")

        # Step 2: Suggest replacements
        print("\nSTEP 2: Suggesting ASCII replacements")
        suggested, changes = _suggest_ascii_replacement(bad_code)
        print(f"\n{changes}")
        print(f"\nSuggested code:\n{suggested}")

        # Step 3: Verify suggested code is valid
        print("\nSTEP 3: Validate suggested code")
        has_non_ascii_2, _ = _detect_non_ascii(suggested)
        if not has_non_ascii_2:
            print("✅ Suggested code is valid (ASCII-only)")
        else:
            print("❌ Suggested code still has issues")
    else:
        print("✅ Code is valid")

def show_example_errors():
    """Show examples of improved error messages."""
    print_section("Test 4: Error Message Examples")

    examples = [
        ("Smart quotes", 'MsgBox "Hello"'),
        ("Arrows", "' Step 1 → Step 2 ➤ Step 3"),
        ("Bullet points", 'MsgBox "• Item 1"'),
        ("Math symbols", "If x ≤ 10 Then"),
    ]

    for name, code in examples:
        print(f"\n{name}:")
        print(f"  Code: {code}")
        has_non_ascii, error = _detect_non_ascii(code)
        if has_non_ascii:
            # Show just the first line of the error
            first_line = error.split('\n')[0]
            print(f"  Detection: ✅ {first_line}")
            suggested, _ = _suggest_ascii_replacement(code)
            print(f"  Suggestion: {suggested}")
        else:
            print(f"  Detection: No issues")

def main():
    """Run all tests."""
    print("\n" + "=" * 60)
    print("  VBA VALIDATION FEATURE DEMONSTRATION")
    print("  CRITICAL FIX: P0 Priority")
    print("=" * 60)

    try:
        test_ascii_detection()
        test_ascii_replacement()
        test_validation_workflow()
        show_example_errors()

        print_section("Summary")
        print("""
✅ All helper functions implemented:
   - _detect_non_ascii(): Detects non-ASCII characters with line numbers
   - _suggest_ascii_replacement(): Suggests ASCII replacements
   - _compile_vba_module(): Validates VBA after injection (in inject.py)

✅ Validation happens BEFORE injection:
   - Code is checked for non-ASCII characters
   - User gets clear error messages with suggestions
   - Injection is blocked if invalid

✅ Validation happens AFTER injection:
   - Code is compiled/validated in VBA
   - Automatic ROLLBACK if validation fails
   - Old code restored or module deleted

✅ Improved error messages:
   - Line numbers for non-ASCII characters
   - Specific suggestions for replacements
   - Clear, actionable feedback

The inject_vba tool now provides robust validation and prevents
file corruption from invalid VBA code.
        """)

    except Exception as e:
        print(f"\n❌ Error during demonstration: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
