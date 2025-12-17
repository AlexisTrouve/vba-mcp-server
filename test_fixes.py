"""
Test script for VBA validation fixes.
"""
import sys
sys.path.insert(0, r'C:\Users\alexi\Documents\projects\vba-mcp-monorepo\packages\pro\src')

from vba_mcp_pro.tools.inject import _check_vba_syntax, _normalize_vba_code

def test_syntax_validation():
    """Test the new syntax validation."""
    print("=" * 60)
    print("Testing VBA Syntax Validation")
    print("=" * 60)

    # Test 1: Valid code
    print("\n1. Testing VALID code (should pass):")
    valid_code = """
Sub Test()
    If True Then
        MsgBox "Hello"
    End If
End Sub
"""
    success, error = _check_vba_syntax(valid_code)
    print(f"   Result: {'✓ PASS' if success else '✗ FAIL'}")
    if error:
        print(f"   Error: {error}")

    # Test 2: Missing End If (SHOULD FAIL)
    print("\n2. Testing INVALID code - Missing End If (should fail):")
    invalid_code1 = """
Sub Test()
    If True Then
        MsgBox "Hello"
End Sub
"""
    success, error = _check_vba_syntax(invalid_code1)
    print(f"   Result: {'✓ DETECTED' if not success else '✗ MISSED'}")
    if error:
        print(f"   Error: {error}")

    # Test 3: Missing Next (SHOULD FAIL)
    print("\n3. Testing INVALID code - Missing Next (should fail):")
    invalid_code2 = """
Sub Test()
    For i = 1 To 10
        MsgBox i
End Sub
"""
    success, error = _check_vba_syntax(invalid_code2)
    print(f"   Result: {'✓ DETECTED' if not success else '✗ MISSED'}")
    if error:
        print(f"   Error: {error}")

    # Test 4: Single-line If (SHOULD PASS)
    print("\n4. Testing VALID code - Single-line If (should pass):")
    valid_code2 = """
Sub Test()
    If True Then MsgBox "Hello"
End Sub
"""
    success, error = _check_vba_syntax(valid_code2)
    print(f"   Result: {'✓ PASS' if success else '✗ FAIL'}")
    if error:
        print(f"   Error: {error}")

    # Test 5: Missing End Function (SHOULD FAIL)
    print("\n5. Testing INVALID code - Missing End Function (should fail):")
    invalid_code3 = """
Function Add(x, y)
    Add = x + y
"""
    success, error = _check_vba_syntax(invalid_code3)
    print(f"   Result: {'✓ DETECTED' if not success else '✗ MISSED'}")
    if error:
        print(f"   Error: {error}")

    # Test 6: Complex valid code
    print("\n6. Testing VALID complex code (should pass):")
    complex_valid = """
Option Explicit

Public Function Calculate(x As Long, y As Long) As Long
    Dim result As Long
    result = 0

    For i = 1 To x
        If i Mod 2 = 0 Then
            result = result + i
        Else
            result = result - i
        End If
    Next i

    Calculate = result + y
End Function

Sub Main()
    MsgBox Calculate(10, 5)
End Sub
"""
    success, error = _check_vba_syntax(complex_valid)
    print(f"   Result: {'✓ PASS' if success else '✗ FAIL'}")
    if error:
        print(f"   Error: {error}")


def test_code_normalization():
    """Test code normalization."""
    print("\n" + "=" * 60)
    print("Testing VBA Code Normalization")
    print("=" * 60)

    code1 = "  \nSub Test()\n    MsgBox \"Hi\"\nEnd Sub\n  \n"
    code2 = "Sub Test()\n    MsgBox \"Hi\"\nEnd Sub"

    normalized1 = _normalize_vba_code(code1)
    normalized2 = _normalize_vba_code(code2)

    print(f"\nCode 1: {repr(code1)}")
    print(f"Normalized: {repr(normalized1)}")
    print(f"\nCode 2: {repr(code2)}")
    print(f"Normalized: {repr(normalized2)}")
    print(f"\nAre they equal after normalization? {normalized1 == normalized2}")


if __name__ == "__main__":
    test_syntax_validation()
    test_code_normalization()
    print("\n" + "=" * 60)
    print("All tests completed!")
    print("=" * 60)
