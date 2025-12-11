# VBA MCP Server

**Model Context Protocol (MCP) server for extracting and analyzing VBA code from Microsoft Office files.**

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)

This MCP server enables Claude Code (and other MCP clients) to interact with VBA macros in Excel, Access, and Word files.

## Features

### Lite Version (Open Source)
- ✅ **Extract VBA code** from Office files (.xlsm, .xlsb, .accdb, .docm)
- ✅ **List all modules** in a workbook/document
- ✅ **Analyze VBA structure** (procedures, functions, dependencies)
- ✅ **Read-only access** to macro content

### Pro Version (Private)
> Contact for enterprise licensing

- 🔒 **Modify and reinject** VBA code into Office files
- 🔒 **Automated refactoring** with AI assistance
- 🔒 **Run macros** from command line
- 🔒 **Testing framework** integration
- 🔒 **Backup and version control** for VBA projects

## Quick Start

### Prerequisites

- Python 3.8+
- pip

### Installation

```bash
# Clone the repository
git clone https://github.com/AlexisTrouve/vba-mcp-server.git
cd vba-mcp-server

# Install dependencies
pip install -r requirements.txt
```

### Configure with Claude Code

Add to your `.mcp.json` or project settings:

```json
{
  "mcpServers": {
    "vba": {
      "command": "python",
      "args": ["path/to/vba-mcp-server/src/server.py"]
    }
  }
}
```

Or using Claude Code CLI:

```bash
claude mcp add --transport stdio vba \
  --scope project \
  -- python path/to/vba-mcp-server/src/server.py
```

## Usage

Once configured, you can use the following commands in Claude Code:

### Extract VBA from a file

```
Extract the VBA code from myfile.xlsm
```

Claude Code will call the `extract_vba` tool and show you the macro content.

### List all modules

```
What VBA modules are in budget.xlsm?
```

### Analyze code structure

```
Analyze the VBA structure in report.xlsm
```

## How It Works

```
Excel/Access file (.xlsm, .accdb)
    ↓
MCP Server (Python)
    ├── Extracts VBA using oletools/openpyxl
    ├── Parses module structure
    └── Returns code to Claude Code
         ↓
    Claude Code displays and helps you analyze/refactor
```

## Architecture

```
vba-mcp-server/
├── src/
│   ├── server.py              # MCP server entry point
│   ├── tools/
│   │   ├── extract.py         # VBA extraction tool
│   │   ├── list_modules.py    # List modules tool
│   │   └── analyze.py         # Code analysis tool
│   └── lib/
│       ├── office_handler.py  # Handle Office files
│       └── vba_parser.py      # Parse VBA structure
├── docs/
│   ├── ARCHITECTURE.md        # Technical architecture
│   ├── API.md                 # MCP tools API reference
│   └── EXAMPLES.md            # Usage examples
├── examples/
│   └── sample.xlsm            # Sample Excel file with VBA
├── tests/
│   └── test_extract.py        # Unit tests
├── requirements.txt           # Python dependencies
├── .gitignore
├── LICENSE
└── README.md
```

## Supported File Formats

| Format | Description | Status |
|--------|-------------|--------|
| `.xlsm` | Excel Macro-Enabled Workbook | ✅ Supported |
| `.xlsb` | Excel Binary Workbook | ✅ Supported |
| `.accdb` | Access Database | ✅ Supported |
| `.docm` | Word Macro-Enabled Document | ✅ Supported |
| `.xls` | Legacy Excel | 🚧 Planned |
| `.mdb` | Legacy Access | 🚧 Planned |

## Technical Details

### Dependencies

- **oletools** - Extract VBA from OLE2 files
- **openpyxl** - Parse Excel OOXML files
- **mcp** - Model Context Protocol SDK

### Security

⚠️ **Important**: VBA macros can contain malicious code. This tool:
- Never executes macros (lite version)
- Only reads and extracts source code
- Recommends scanning with antivirus before opening Office files

## Development

### Running Tests

```bash
pytest tests/
```

### Contributing

This is the **lite/open-source version**. Contributions welcome for:
- Bug fixes
- Documentation improvements
- New extraction features
- Support for additional file formats

Please ensure:
- Code follows PEP 8
- Tests pass
- Documentation is updated

## Roadmap

### Version 1.0 (Lite - Open Source)
- [x] Extract VBA from .xlsm files
- [x] List modules and procedures
- [x] Basic structure analysis
- [ ] Support .accdb files
- [ ] Support .docm files
- [ ] Cross-platform compatibility

### Version 2.0 (Pro - Private)
- [ ] VBA modification and reinjection
- [ ] Automated refactoring
- [ ] Macro execution
- [ ] Testing framework
- [ ] Version control integration

## License

MIT License - see [LICENSE](LICENSE) file for details.

## Pro Version

For enterprise features (modification, reinjection, automated refactoring), contact:
- Email: alexistrouve.pro@gmail.com
- GitHub: [@AlexisTrouve](https://github.com/AlexisTrouve)

## Related Projects

- [Claude Code](https://claude.com/claude-code) - AI-powered coding assistant
- [Model Context Protocol](https://modelcontextprotocol.io) - Open protocol for AI tool integration
- [oletools](https://github.com/decalage2/oletools) - Tools for Office file analysis

---

**Note**: This tool is for legitimate VBA development and maintenance. Always ensure you have permission to access and modify Office files.
