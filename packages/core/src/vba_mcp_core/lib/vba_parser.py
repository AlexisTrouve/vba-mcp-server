"""
VBA Code Parser

Parses VBA source code to extract procedures, variables, dependencies, etc.
"""

import re
from typing import Dict, List, Optional


class VBAParser:
    """
    Parser for VBA source code.

    Extracts procedures, functions, dependencies, and calculates complexity.
    """

    # VBA procedure patterns
    SUB_PATTERN = re.compile(
        r'^\s*(Public|Private|Friend)?\s*(Static)?\s*Sub\s+(\w+)\s*\(',
        re.MULTILINE | re.IGNORECASE
    )

    FUNCTION_PATTERN = re.compile(
        r'^\s*(Public|Private|Friend)?\s*(Static)?\s*Function\s+(\w+)\s*\(',
        re.MULTILINE | re.IGNORECASE
    )

    PROPERTY_PATTERN = re.compile(
        r'^\s*(Public|Private|Friend)?\s*Property\s+(Get|Set|Let)\s+(\w+)\s*\(',
        re.MULTILINE | re.IGNORECASE
    )

    # Call pattern (simplified)
    CALL_PATTERN = re.compile(r'\b(\w+)\s*\(', re.MULTILINE)

    # Variable declaration
    DIM_PATTERN = re.compile(
        r'^\s*Dim\s+(\w+)\s+As\s+(\w+)',
        re.MULTILINE | re.IGNORECASE
    )

    def parse_module(self, module: Dict) -> Dict:
        """
        Parse a VBA module.

        Args:
            module: Module dictionary with 'name', 'code', etc.

        Returns:
            Enhanced module dictionary with parsed information
        """
        code = module.get("code", "")

        # Extract procedures
        procedures = self._extract_procedures(code)

        # Extract dependencies (called modules)
        dependencies = self._extract_dependencies(code, procedures)

        # Calculate complexity
        for proc in procedures:
            proc["complexity"] = self._calculate_complexity(
                code, proc["line_start"], proc["line_end"]
            )

        return {
            **module,
            "procedures": procedures,
            "dependencies": dependencies
        }

    def _extract_procedures(self, code: str) -> List[Dict]:
        """
        Extract all procedures (Subs, Functions, Properties) from code.

        Args:
            code: VBA source code

        Returns:
            List of procedure dictionaries
        """
        procedures = []
        lines = code.splitlines()

        # Find Subs
        for match in self.SUB_PATTERN.finditer(code):
            visibility = match.group(1) or "Public"
            name = match.group(3)
            line_num = code[:match.start()].count('\n') + 1

            # Find End Sub
            end_line = self._find_end_statement(lines, line_num, "Sub")

            # Extract calls
            proc_code = '\n'.join(lines[line_num-1:end_line])
            calls = self._extract_calls(proc_code)

            procedures.append({
                "name": name,
                "type": "Sub",
                "visibility": visibility,
                "line_start": line_num,
                "line_end": end_line,
                "calls": calls,
                "parameters": []  # TODO: parse parameters
            })

        # Find Functions
        for match in self.FUNCTION_PATTERN.finditer(code):
            visibility = match.group(1) or "Public"
            name = match.group(3)
            line_num = code[:match.start()].count('\n') + 1

            end_line = self._find_end_statement(lines, line_num, "Function")
            proc_code = '\n'.join(lines[line_num-1:end_line])
            calls = self._extract_calls(proc_code)

            procedures.append({
                "name": name,
                "type": "Function",
                "visibility": visibility,
                "line_start": line_num,
                "line_end": end_line,
                "calls": calls,
                "parameters": []
            })

        # Find Properties
        for match in self.PROPERTY_PATTERN.finditer(code):
            visibility = match.group(1) or "Public"
            prop_type = match.group(2)
            name = match.group(3)
            line_num = code[:match.start()].count('\n') + 1

            end_line = self._find_end_statement(lines, line_num, "Property")

            procedures.append({
                "name": name,
                "type": f"Property {prop_type}",
                "visibility": visibility,
                "line_start": line_num,
                "line_end": end_line,
                "calls": [],
                "parameters": []
            })

        return procedures

    def _find_end_statement(self, lines: List[str], start_line: int, statement_type: str) -> int:
        """
        Find the 'End Sub/Function/Property' for a procedure.

        Args:
            lines: Code lines
            start_line: Starting line number (1-indexed)
            statement_type: "Sub", "Function", or "Property"

        Returns:
            End line number (1-indexed)
        """
        end_pattern = re.compile(rf'^\s*End\s+{statement_type}\b', re.IGNORECASE)

        for i in range(start_line, len(lines) + 1):
            if i > len(lines):
                return len(lines)

            if end_pattern.match(lines[i - 1]):
                return i

        return len(lines)

    def _extract_calls(self, code: str) -> List[str]:
        """
        Extract function/sub calls from code.

        Args:
            code: VBA code snippet

        Returns:
            List of called function names
        """
        calls = set()

        # Find all potential calls
        for match in self.CALL_PATTERN.finditer(code):
            func_name = match.group(1)

            # Filter out VBA keywords and common built-ins
            if not self._is_vba_keyword(func_name):
                calls.add(func_name)

        return sorted(list(calls))

    def _is_vba_keyword(self, word: str) -> bool:
        """
        Check if word is a VBA keyword.

        Args:
            word: Word to check

        Returns:
            True if keyword
        """
        keywords = {
            'If', 'Then', 'Else', 'ElseIf', 'End', 'For', 'Next', 'Do', 'Loop',
            'While', 'Wend', 'Select', 'Case', 'With', 'Exit', 'Sub', 'Function',
            'Property', 'Public', 'Private', 'Dim', 'ReDim', 'Const', 'Type',
            'Enum', 'Class', 'New', 'Set', 'Let', 'Get', 'Call', 'Return'
        }
        return word.lower() in {k.lower() for k in keywords}

    def _extract_dependencies(self, code: str, procedures: List[Dict]) -> List[str]:
        """
        Extract module dependencies (very basic).

        Args:
            code: VBA code
            procedures: List of procedures

        Returns:
            List of potentially referenced modules
        """
        # TODO: Implement proper dependency analysis
        # This would require understanding module.procedure calls
        return []

    def _calculate_complexity(self, code: str, start_line: int, end_line: int) -> int:
        """
        Calculate cyclomatic complexity.

        Args:
            code: Full code
            start_line: Procedure start
            end_line: Procedure end

        Returns:
            Complexity score
        """
        lines = code.splitlines()
        proc_code = '\n'.join(lines[start_line-1:end_line])

        complexity = 1  # Base complexity

        # Count decision points
        decision_keywords = ['If', 'ElseIf', 'For', 'While', 'Do', 'Case', 'And', 'Or']

        for keyword in decision_keywords:
            pattern = re.compile(rf'\b{keyword}\b', re.IGNORECASE)
            complexity += len(pattern.findall(proc_code))

        return complexity
