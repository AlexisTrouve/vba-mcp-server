"""
Office File Handler

Handles opening Office files and extracting VBA projects.
Supports .xlsm, .xlsb, .accdb, .docm formats.
"""

import zipfile
from pathlib import Path
from typing import Dict, List, Optional

try:
    from oletools.olevba import VBA_Parser
    OLETOOLS_AVAILABLE = True
except ImportError:
    OLETOOLS_AVAILABLE = False


class OfficeHandler:
    """
    Handler for Microsoft Office files.

    Detects file format and extracts VBA code using appropriate method.
    """

    SUPPORTED_FORMATS = {
        '.xlsm': 'Excel Macro-Enabled Workbook',
        '.xlsb': 'Excel Binary Workbook',
        '.accdb': 'Access Database',
        '.docm': 'Word Macro-Enabled Document',
        '.pptm': 'PowerPoint Macro-Enabled Presentation'
    }

    def extract_vba_project(self, file_path: Path) -> Dict:
        """
        Extract VBA project from Office file.

        Args:
            file_path: Path to Office file

        Returns:
            Dictionary with modules and metadata

        Raises:
            ValueError: If file format not supported
            FileNotFoundError: If file doesn't exist
        """
        if not file_path.exists():
            raise FileNotFoundError(f"File not found: {file_path}")

        file_format = file_path.suffix.lower()

        if file_format not in self.SUPPORTED_FORMATS:
            raise ValueError(
                f"Unsupported format: {file_format}. "
                f"Supported: {', '.join(self.SUPPORTED_FORMATS.keys())}"
            )

        # Use oletools if available
        if OLETOOLS_AVAILABLE:
            return self._extract_with_oletools(file_path)
        else:
            # Fallback to manual OOXML extraction
            return self._extract_ooxml(file_path)

    def _extract_with_oletools(self, file_path: Path) -> Dict:
        """
        Extract VBA using oletools library.

        Args:
            file_path: Path to Office file

        Returns:
            VBA project dictionary
        """
        try:
            vba_parser = VBA_Parser(str(file_path))

            # Check if VBA macros exist
            if not vba_parser.detect_vba_macros():
                vba_parser.close()
                return {"modules": []}

            # Extract all modules
            modules = []

            for (filename, stream_path, vba_filename, vba_code) in vba_parser.extract_macros():
                if vba_code:
                    # Parse module name from stream path
                    module_name = self._parse_module_name(vba_filename or stream_path)

                    # Determine module type
                    module_type = self._determine_module_type(module_name, stream_path)

                    modules.append({
                        "name": module_name,
                        "type": module_type,
                        "code": vba_code.decode('utf-8', errors='ignore') if isinstance(vba_code, bytes) else vba_code,
                        "line_count": len(vba_code.splitlines()) if isinstance(vba_code, str) else len(vba_code.decode('utf-8', errors='ignore').splitlines())
                    })

            vba_parser.close()

            return {"modules": modules}

        except Exception as e:
            raise ValueError(f"Failed to extract VBA with oletools: {str(e)}")

    def _extract_ooxml(self, file_path: Path) -> Dict:
        """
        Extract VBA from OOXML files manually (fallback).

        Args:
            file_path: Path to OOXML file (.xlsm, .docm)

        Returns:
            VBA project dictionary
        """
        try:
            # OOXML files are ZIP archives
            with zipfile.ZipFile(file_path, 'r') as zip_file:
                # Look for vbaProject.bin
                vba_bin_paths = [
                    'xl/vbaProject.bin',        # Excel
                    'word/vbaProject.bin',      # Word
                    'ppt/vbaProject.bin'        # PowerPoint
                ]

                vba_bin = None
                for path in vba_bin_paths:
                    try:
                        vba_bin = zip_file.read(path)
                        break
                    except KeyError:
                        continue

                if not vba_bin:
                    return {"modules": []}

                # vbaProject.bin is an OLE2 file, needs oletools to parse
                if not OLETOOLS_AVAILABLE:
                    raise ValueError(
                        "oletools library required for VBA extraction. "
                        "Install with: pip install oletools"
                    )

                # Parse VBA from binary
                vba_parser = VBA_Parser('vbaProject.bin', data=vba_bin)

                modules = []
                for (filename, stream_path, vba_filename, vba_code) in vba_parser.extract_macros():
                    if vba_code:
                        module_name = self._parse_module_name(vba_filename or stream_path)
                        module_type = self._determine_module_type(module_name, stream_path)

                        modules.append({
                            "name": module_name,
                            "type": module_type,
                            "code": vba_code.decode('utf-8', errors='ignore') if isinstance(vba_code, bytes) else vba_code,
                            "line_count": len(vba_code.splitlines()) if isinstance(vba_code, str) else len(vba_code.decode('utf-8', errors='ignore').splitlines())
                        })

                vba_parser.close()
                return {"modules": modules}

        except zipfile.BadZipFile:
            raise ValueError("File is not a valid OOXML (ZIP) file")
        except Exception as e:
            raise ValueError(f"Failed to extract OOXML VBA: {str(e)}")

    def _parse_module_name(self, path: str) -> str:
        """
        Parse module name from VBA path.

        Args:
            path: VBA stream path or filename

        Returns:
            Module name
        """
        if '/' in path:
            return path.split('/')[-1]
        return path

    def _determine_module_type(self, module_name: str, stream_path: str) -> str:
        """
        Determine VBA module type.

        Args:
            module_name: Module name
            stream_path: OLE stream path

        Returns:
            Module type (standard, class, worksheet, workbook, form)
        """
        module_lower = module_name.lower()

        if module_lower == 'thisworkbook':
            return 'workbook'
        elif module_lower.startswith('sheet'):
            return 'worksheet'
        elif module_lower.startswith('userform'):
            return 'form'
        elif 'class' in stream_path.lower():
            return 'class'
        else:
            return 'standard'
