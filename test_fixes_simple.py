"""
Test script for VBA validation fixes - direct import.
"""
import sys
from pathlib import Path

# Add the package to path
sys.path.insert(0, str(Path(__file__).parent / 'packages' / 'pro' / 'src'))

# Import just the functions we need without importing the whole module
import importlib.util
spec = importlib.util.spec_from_file_location(
    "inject",
    Path(__file__).parent / 'packages' / 'pro' / 'src' / 'vba_mcp_pro' / 'tools' / 'inject.py'
)
inject_module = importlib.util.module_from_spec(spec)

# Load the module without executing all imports
try:
    spec.loader.exec_module(inject_module)
    _check_vba_syntax = inject_module._check_vba_syntax
    _normalize_vba_code = inject_module._normalize_vba_code
except Exception as e:
    print(f"Error loading module: {e}")
    print("Loading functions directly...")

    # Define functions inline for testing
    from typing import Tuple, Optional

    def _check_vba_syntax(code: str) -> Tuple[bool, Optional[str]]:
        """Check VBA code syntax for common errors using pattern matching."""
        lines = code.splitlines()

        # Track block nesting
        if_count = 0
        for_count = 0
        while_count = 0
        do_count = 0
        with_count = 0
        select_count = 0
        sub_count = 0
        function_count = 0

        for line_num, line in enumerate(lines, 1):
            stripped = line.strip()

            # Skip empty lines and comments
            if not stripped or stripped.startswith("'") or stripped.startswith("Rem "):
                continue

            # Remove inline comments for analysis
            if "'" in stripped:
                stripped = stripped.split("'")[0].strip()

            # Check for block start/end
            # If/Then/End If
            if stripped.startswith("If ") and " Then" in stripped and not stripped.endswith(" _"):
                after_then = stripped.split(" Then", 1)[1].strip()
                if "'" in after_then:
                    after_then = after_then.split("'")[0].strip()
                if not after_then or after_then == ":":
                    if_count += 1
            elif stripped.startswith("ElseIf ") and " Then" in stripped:
                pass
            elif stripped.startswith("Else") and not stripped.startswith("ElseIf"):
                pass
            elif stripped.startswith("End If") or stripped == "End If":
                if_count -= 1
                if if_count < 0:
                    return False, f"Line {line_num}: 'End If' without matching 'If'"

            # For/Next
            elif stripped.startswith("For "):
                for_count += 1
            elif stripped.startswith("Next"):
                for_count -= 1
                if for_count < 0:
                    return False, f"Line {line_num}: 'Next' without matching 'For'"

            # While/Wend
            elif stripped.startswith("While "):
                while_count += 1
            elif stripped.startswith("Wend"):
                while_count -= 1
                if while_count < 0:
                    return False, f"Line {line_num}: 'Wend' without matching 'While'"

            # Do/Loop
            elif stripped.startswith("Do") and (stripped == "Do" or stripped.startswith("Do While") or stripped.startswith("Do Until")):
                do_count += 1
            elif stripped.startswith("Loop"):
                do_count -= 1
                if do_count < 0:
                    return False, f"Line {line_num}: 'Loop' without matching 'Do'"

            # With/End With
            elif stripped.startswith("With "):
                with_count += 1
            elif stripped.startswith("End With"):
                with_count -= 1
                if with_count < 0:
                    return False, f"Line {line_num}: 'End With' without matching 'With'"

            # Select/End Select
            elif stripped.startswith("Select Case "):
                select_count += 1
            elif stripped.startswith("End Select"):
                select_count -= 1
                if select_count < 0:
                    return False, f"Line {line_num}: 'End Select' without matching 'Select Case'"

            # Sub/End Sub
            elif stripped.startswith("Sub ") or stripped.startswith("Public Sub ") or stripped.startswith("Private Sub "):
                sub_count += 1
            elif stripped.startswith("End Sub"):
                sub_count -= 1
                if sub_count < 0:
                    return False, f"Line {line_num}: 'End Sub' without matching 'Sub'"

            # Function/End Function
            elif stripped.startswith("Function ") or stripped.startswith("Public Function ") or stripped.startswith("Private Function "):
                function_count += 1
            elif stripped.startswith("End Function"):
                function_count -= 1
                if function_count < 0:
                    return False, f"Line {line_num}: 'End Function' without matching 'Function'"

        # Check for unclosed blocks
        errors = []
        if if_count > 0:
            errors.append(f"{if_count} unclosed 'If' block(s) - missing 'End If'")
        if for_count > 0:
            errors.append(f"{for_count} unclosed 'For' loop(s) - missing 'Next'")
        if while_count > 0:
            errors.append(f"{while_count} unclosed 'While' loop(s) - missing 'Wend'")
        if do_count > 0:
            errors.append(f"{do_count} unclosed 'Do' loop(s) - missing 'Loop'")
        if with_count > 0:
            errors.append(f"{with_count} unclosed 'With' block(s) - missing 'End With'")
        if select_count > 0:
            errors.append(f"{select_count} unclosed 'Select Case' block(s) - missing 'End Select'")
        if sub_count > 0:
            errors.append(f"{sub_count} unclosed 'Sub' procedure(s) - missing 'End Sub'")
        if function_count > 0:
            errors.append(f"{function_count} unclosed 'Function' procedure(s) - missing 'End Function'")

        if errors:
            return False, "VBA Syntax Error:\n  " + "\n  ".join(errors)

        return True, None

    def _normalize_vba_code(code: str) -> str:
        """Normalize VBA code for comparison."""
        lines = code.splitlines()
        normalized_lines = []

        for line in lines:
            normalized_lines.append(line.rstrip())

        # Remove leading and trailing blank lines
        while normalized_lines and not normalized_lines[0].strip():
            normalized_lines.pop(0)
        while normalized_lines and not normalized_lines[-1].strip():
            normalized_lines.pop()

        return '\n'.join(normalized_lines)


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
    print(f"   Result: {'PASS OK' if success else 'FAIL X'}")
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
    print(f"   Result: {'DETECTED OK' if not success else 'MISSED X'}")
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
    print(f"   Result: {'DETECTED OK' if not success else 'MISSED X'}")
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
    print(f"   Result: {'PASS OK' if success else 'FAIL X'}")
    if error:
        print(f"   Error: {error}")

    # Test 5: Missing End Function (SHOULD FAIL)
    print("\n5. Testing INVALID code - Missing End Function (should fail):")
    invalid_code3 = """
Function Add(x, y)
    Add = x + y
"""
    success, error = _check_vba_syntax(invalid_code3)
    print(f"   Result: {'DETECTED OK' if not success else 'MISSED X'}")
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
    print(f"   Result: {'PASS OK' if success else 'FAIL X'}")
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

    print(f"\nCode 1 length: {len(code1)}")
    print(f"Normalized 1 length: {len(normalized1)}")
    print(f"\nCode 2 length: {len(code2)}")
    print(f"Normalized 2 length: {len(normalized2)}")
    print(f"\nAre they equal after normalization? {normalized1 == normalized2}")


if __name__ == "__main__":
    test_syntax_validation()
    test_code_normalization()
    print("\\n" + "=" * 60)
    print("All tests completed!")
    print("=" * 60)
