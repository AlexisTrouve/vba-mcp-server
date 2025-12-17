"""Tests for VBAParser."""
import pytest

from vba_mcp_core.lib.vba_parser import VBAParser


class TestVBAParser:
    """Test suite for VBAParser class."""

    @pytest.fixture
    def parser(self):
        """Create a VBAParser instance."""
        return VBAParser()

    @pytest.fixture
    def simple_code(self):
        """Simple VBA code sample."""
        return """
Public Function HelloWorld() As String
    HelloWorld = "Hello!"
End Function

Private Sub LogMessage(msg As String)
    Debug.Print msg
End Sub
"""

    @pytest.fixture
    def complex_code(self):
        """Complex VBA code with multiple structures."""
        return """
Public Function Calculate(x As Integer) As Integer
    Dim result As Integer

    If x > 10 Then
        result = x * 2
    ElseIf x > 5 Then
        result = x + 10
    Else
        result = x
    End If

    For i = 1 To 5
        result = result + i
    Next i

    Calculate = result
End Function
"""

    def test_parse_module_basic(self, parser):
        """Test basic module parsing."""
        module = {
            "name": "TestModule",
            "type": "standard",
            "code": "Public Sub Test()\nEnd Sub",
            "line_count": 2
        }

        result = parser.parse_module(module)

        assert "procedures" in result
        assert "dependencies" in result
        assert result["name"] == "TestModule"

    def test_extract_procedures_sub(self, parser, simple_code):
        """Test extraction of Sub procedures."""
        procedures = parser._extract_procedures(simple_code)

        # Find the LogMessage sub
        log_sub = next((p for p in procedures if p["name"] == "LogMessage"), None)
        assert log_sub is not None
        assert log_sub["type"] == "Sub"
        assert log_sub["visibility"] == "Private"

    def test_extract_procedures_function(self, parser, simple_code):
        """Test extraction of Function procedures."""
        procedures = parser._extract_procedures(simple_code)

        # Find the HelloWorld function
        hello_func = next((p for p in procedures if p["name"] == "HelloWorld"), None)
        assert hello_func is not None
        assert hello_func["type"] == "Function"
        assert hello_func["visibility"] == "Public"

    def test_extract_calls(self, parser):
        """Test extraction of function calls."""
        code = """
    result = Calculate(10)
    DoSomething(x, y)
    Helper()
"""
        calls = parser._extract_calls(code)

        assert "Calculate" in calls
        assert "DoSomething" in calls
        assert "Helper" in calls

    def test_is_vba_keyword(self, parser):
        """Test VBA keyword detection."""
        assert parser._is_vba_keyword("If") is True
        assert parser._is_vba_keyword("Then") is True
        assert parser._is_vba_keyword("End") is True
        assert parser._is_vba_keyword("MyFunction") is False
        assert parser._is_vba_keyword("Calculate") is False

    def test_calculate_complexity_simple(self, parser):
        """Test complexity calculation for simple code."""
        code = """
Public Sub Simple()
    x = 1
End Sub
"""
        lines = code.splitlines()
        complexity = parser._calculate_complexity(code, 2, 4)

        # Base complexity is 1, no decision points
        assert complexity == 1

    def test_calculate_complexity_with_conditionals(self, parser, complex_code):
        """Test complexity calculation with conditionals."""
        procedures = parser._extract_procedures(complex_code)
        calc_func = procedures[0]

        complexity = parser._calculate_complexity(
            complex_code,
            calc_func["line_start"],
            calc_func["line_end"]
        )

        # Base (1) + If (1) + ElseIf (1) + For (1) = 4
        assert complexity >= 4

    def test_find_end_statement(self, parser):
        """Test finding end of procedure."""
        code = """
Public Sub Test()
    x = 1
    y = 2
End Sub

Public Sub Another()
"""
        lines = code.splitlines()

        # Find end of first sub (line 5)
        end_line = parser._find_end_statement(lines, 2, "Sub")
        assert end_line == 5

    def test_parse_module_with_real_code(self, parser):
        """Test parsing a realistic VBA module."""
        module = {
            "name": "TestModule",
            "type": "standard",
            "code": """
Option Explicit

Public Function Add(a As Integer, b As Integer) As Integer
    Add = a + b
End Function

Private Sub Initialize()
    Dim x As Integer
    x = Add(5, 10)
    MsgBox x
End Sub
""",
            "line_count": 11
        }

        result = parser.parse_module(module)

        assert len(result["procedures"]) == 2

        # Check Add function
        add_func = next((p for p in result["procedures"] if p["name"] == "Add"), None)
        assert add_func is not None
        assert add_func["type"] == "Function"

        # Check Initialize sub
        init_sub = next((p for p in result["procedures"] if p["name"] == "Initialize"), None)
        assert init_sub is not None
        assert init_sub["type"] == "Sub"
        assert "Add" in init_sub["calls"]
        # Note: MsgBox without parentheses is not detected as a call by the parser
