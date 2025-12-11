# VBA MCP Core

Core library for VBA extraction and analysis from Microsoft Office files.

This package provides the shared functionality used by both the lite and pro versions of VBA MCP Server.

## Features

- Extract VBA code from Office files (.xlsm, .docm, .accdb, .pptm)
- Parse VBA procedures (Sub, Function, Property)
- Calculate code complexity metrics
- Analyze module dependencies

## Installation

```bash
pip install vba-mcp-core
```

## Usage

```python
from vba_mcp_core import OfficeHandler, VBAParser
from pathlib import Path

# Extract VBA from Excel file
handler = OfficeHandler()
vba_project = handler.extract_vba_project(Path("workbook.xlsm"))

# Parse modules
parser = VBAParser()
for module in vba_project["modules"]:
    parsed = parser.parse_module(module)
    print(f"Module: {parsed['name']}")
    for proc in parsed["procedures"]:
        print(f"  - {proc['name']} (complexity: {proc['complexity']})")
```

## License

MIT License - See LICENSE file.
