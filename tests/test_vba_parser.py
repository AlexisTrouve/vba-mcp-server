"""
Tests for VBA Parser
"""

import pytest
import sys
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from lib.vba_parser import VBAParser


class TestVBAParser:
    """Test suite for VBA Parser"""

    @pytest.fixture
    def parser(self):
        """Create a VBAParser instance"""
        return VBAParser()

    @pytest.fixture
    def simple_sub_code(self):
        """Simple VBA Sub procedure"""
        return """
Sub HelloWorld()
    MsgBox "Hello!"
End Sub
"""

    @pytest.fixture
    def simple_function_code(self):
        """Simple VBA Function"""
        return """
Function AddNumbers(a As Double, b As Double) As Double
    AddNumbers = a + b
End Function
"""

    @pytest.fixture
    def complex_code(self):
        """Complex VBA code with multiple procedures"""
        return """
Option Explicit

Public gCounter As Long

Sub TestProcedure()
    Dim i As Long
    For i = 1 To 10
        If i Mod 2 = 0 Then
            gCounter = gCounter + i
        End If
    Next i
End Sub

Function Calculate(x As Double) As Double
    If x > 0 Then
        Calculate = x * 2
    Else
        Calculate = 0
    End If
End Function

Private Sub PrivateHelper()
    TestProcedure()
End Sub
"""

    def test_parse_simple_sub(self, parser, simple_sub_code):
        """Test parsing a simple Sub procedure"""
        module = {
            "name": "Module1",
            "type": "standard",
            "code": simple_sub_code,
            "line_count": 4
        }

        result = parser.parse_module(module)

        assert "procedures" in result
        assert len(result["procedures"]) == 1

        proc = result["procedures"][0]
        assert proc["name"] == "HelloWorld"
        assert proc["type"] == "Sub"
        assert proc["visibility"] == "Public"

    def test_parse_simple_function(self, parser, simple_function_code):
        """Test parsing a simple Function"""
        module = {
            "name": "Module1",
            "type": "standard",
            "code": simple_function_code,
            "line_count": 4
        }

        result = parser.parse_module(module)

        assert "procedures" in result
        assert len(result["procedures"]) == 1

        proc = result["procedures"][0]
        assert proc["name"] == "AddNumbers"
        assert proc["type"] == "Function"

    def test_parse_complex_code(self, parser, complex_code):
        """Test parsing complex code with multiple procedures"""
        module = {
            "name": "Module1",
            "type": "standard",
            "code": complex_code,
            "line_count": len(complex_code.splitlines())
        }

        result = parser.parse_module(module)

        assert "procedures" in result
        assert len(result["procedures"]) == 3

        # Check procedure names
        proc_names = {p["name"] for p in result["procedures"]}
        assert "TestProcedure" in proc_names
        assert "Calculate" in proc_names
        assert "PrivateHelper" in proc_names

        # Check visibility
        private_procs = [p for p in result["procedures"] if p["visibility"] == "Private"]
        assert len(private_procs) == 1
        assert private_procs[0]["name"] == "PrivateHelper"

    def test_extract_calls(self, parser, complex_code):
        """Test extraction of function calls"""
        module = {
            "name": "Module1",
            "type": "standard",
            "code": complex_code,
            "line_count": len(complex_code.splitlines())
        }

        result = parser.parse_module(module)

        # Find PrivateHelper which calls TestProcedure
        private_helper = [p for p in result["procedures"] if p["name"] == "PrivateHelper"][0]

        # Should detect the call to TestProcedure
        assert "TestProcedure" in private_helper["calls"]

    def test_calculate_complexity(self, parser, complex_code):
        """Test complexity calculation"""
        module = {
            "name": "Module1",
            "type": "standard",
            "code": complex_code,
            "line_count": len(complex_code.splitlines())
        }

        result = parser.parse_module(module)

        # TestProcedure has: For, If, Mod (which could trigger decision)
        test_proc = [p for p in result["procedures"] if p["name"] == "TestProcedure"][0]
        assert test_proc["complexity"] > 1  # Should have some complexity

        # Calculate has: If, Else
        calc_func = [p for p in result["procedures"] if p["name"] == "Calculate"][0]
        assert calc_func["complexity"] > 1

    def test_empty_module(self, parser):
        """Test parsing an empty module"""
        module = {
            "name": "EmptyModule",
            "type": "standard",
            "code": "",
            "line_count": 0
        }

        result = parser.parse_module(module)

        assert "procedures" in result
        assert len(result["procedures"]) == 0
        assert "dependencies" in result
        assert len(result["dependencies"]) == 0

    def test_property_procedures(self, parser):
        """Test parsing Property Get/Set/Let"""
        code = """
Private m_value As String

Public Property Get Value() As String
    Value = m_value
End Property

Public Property Let Value(v As String)
    m_value = v
End Property
"""

        module = {
            "name": "ClassModule1",
            "type": "class",
            "code": code,
            "line_count": len(code.splitlines())
        }

        result = parser.parse_module(module)

        assert len(result["procedures"]) == 2

        prop_names = {p["name"] for p in result["procedures"]}
        assert "Value" in prop_names

        # Check property types
        prop_types = {p["type"] for p in result["procedures"]}
        assert "Property Get" in prop_types
        assert "Property Let" in prop_types

    def test_keyword_filtering(self, parser):
        """Test that VBA keywords are not detected as calls"""
        code = """
Sub TestKeywords()
    If True Then
        For i = 1 To 10
            Do While x > 0
                x = x - 1
            Loop
        Next i
    End If
End Sub
"""

        module = {
            "name": "Module1",
            "type": "standard",
            "code": code,
            "line_count": len(code.splitlines())
        }

        result = parser.parse_module(module)
        proc = result["procedures"][0]

        # Keywords should not appear in calls
        keywords = {'If', 'Then', 'For', 'To', 'Do', 'While', 'Loop', 'Next', 'End'}
        for keyword in keywords:
            assert keyword not in proc["calls"]
