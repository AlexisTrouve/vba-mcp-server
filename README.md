# VBA MCP Server

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)

Model Context Protocol (MCP) server for VBA extraction and analysis in Microsoft Office files.

> **MCP** enables Claude to interact with Office files through specialized tools. Extract, analyze, and understand VBA code from Excel, Word, and Access files.

## Features

### Lite (Free - MIT)
- **Extract VBA code** from Office files (.xlsm, .xlsb, .docm, .accdb)
- **List all VBA modules** and procedures
- **Analyze code structure** and complexity metrics

### Pro (Commercial)
- All Lite features, plus:
- **Inject VBA code** back into Office files
- **Run macros** with parameters
- **Read/write data** from Excel and Access
- **Access Forms** - Create, export, import forms programmatically
- **VBA compilation check** - Detect errors before running
- **Backup/restore** with automatic rollback

See [vba-mcp-server-pro](https://git.etheryale.com/StillHammer/vba-mcp-server-pro) for the commercial version.

## Quick Start

### Installation

```bash
pip install vba-mcp-server
```

### Claude Desktop Configuration

Add to your Claude Desktop config (`%APPDATA%\Claude\claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "vba": {
      "command": "python",
      "args": ["-m", "vba_mcp_server"]
    }
  }
}
```

Restart Claude Desktop.

## MCP Tools

| Tool | Description |
|------|-------------|
| `extract_vba` | Extract VBA code from a specific module |
| `list_modules` | List all VBA modules in a file |
| `analyze_structure` | Analyze code structure and complexity |

## Usage Examples

### Extract VBA Code

```
Extract the VBA code from the "Module1" module in C:\path\to\file.xlsm
```

### List All Modules

```
List all VBA modules in C:\path\to\workbook.xlsm
```

### Analyze Code Structure

```
Analyze the VBA code structure in C:\path\to\file.xlsm
```

## Supported File Types

| Extension | Application | Support |
|-----------|-------------|---------|
| `.xlsm` | Excel (macro-enabled) | Full |
| `.xlsb` | Excel (binary) | Full |
| `.xls` | Excel (legacy) | Full |
| `.docm` | Word (macro-enabled) | Full |
| `.doc` | Word (legacy) | Full |
| `.accdb` | Access | Partial* |

*For full Access support including VBA extraction via COM, see the Pro version.

## Development

### Setup

```bash
# Clone the repository
git clone https://github.com/AlexisTrouve/vba-mcp-server.git
cd vba-mcp-server

# Create virtual environment
python -m venv venv
source venv/bin/activate  # or venv\Scripts\activate on Windows

# Install in editable mode
pip install -e packages/core
pip install -e packages/lite
```

### Project Structure

```
vba-mcp-server/
├── packages/
│   ├── core/       # vba-mcp-core - Shared extraction library
│   └── lite/       # vba-mcp-server - MCP server (this package)
├── docs/           # Documentation
├── examples/       # Example files
└── tests/          # Tests
```

### Running Tests

```bash
pytest packages/lite/tests/
```

## Requirements

- Python 3.8+
- oletools (for VBA extraction)

## Known Limitations

1. **oletools** doesn't fully support `.accdb` files - some Access features require COM (Pro version)
2. Read-only operations - for VBA injection, see the Pro version

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

MIT License - see [LICENSE](LICENSE) for details.

## Author

Alexis Trouve - alexistrouve.pro@gmail.com

## See Also

- [MCP Protocol](https://modelcontextprotocol.io/) - Model Context Protocol specification
- [oletools](https://github.com/decalage2/oletools) - Python tools for OLE/Office files
