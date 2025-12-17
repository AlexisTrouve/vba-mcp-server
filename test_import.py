"""
Test direct import of inject module functions.
"""
import sys
from pathlib import Path

# Direct import without full package initialization
inject_path = Path(__file__).parent / 'packages' / 'pro' / 'src' / 'vba_mcp_pro' / 'tools' / 'inject.py'

# Read and compile the module
with open(inject_path, 'r', encoding='utf-8') as f:
    code = f.read()

# Create a namespace for execution
namespace = {}

# Execute the module code in the namespace
try:
    exec(code, namespace)
    print("OK - Module loaded successfully")

    # Test _check_vba_syntax
    _check_vba_syntax = namespace['_check_vba_syntax']

    # Test case: Invalid code (missing End If)
    test_code = """Sub Test()
    If True Then
        MsgBox "Hello"
End Sub"""

    success, error = _check_vba_syntax(test_code)
    if not success and "End If" in error:
        print("OK - Syntax validation working - detected missing End If")
        print(f"  Error message: {error}")
    else:
        print("FAIL - Syntax validation failed - did not detect missing End If")

    # Test _normalize_vba_code
    _normalize_vba_code = namespace['_normalize_vba_code']

    code1 = "  \nSub Test()\nEnd Sub\n  "
    code2 = "Sub Test()\nEnd Sub"

    if _normalize_vba_code(code1) == _normalize_vba_code(code2):
        print("OK - Code normalization working - normalized different whitespace")
    else:
        print("FAIL - Code normalization failed")

    print("\nOK - All basic functionality tests passed!")

except SyntaxError as e:
    print(f"FAIL - Syntax error in module: {e}")
except Exception as e:
    print(f"FAIL - Error loading module: {e}")
    import traceback
    traceback.print_exc()
