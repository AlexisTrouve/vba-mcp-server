# VBA MCP Server

MCP (Model Context Protocol) server for extracting and analyzing VBA code from Microsoft Office files.

Works with Claude Code to help you understand, analyze, and refactor VBA macros.

## Features

- **Extract VBA** - Pull VBA source code from .xlsm, .docm, .accdb, .pptm files
- **List Modules** - Quick overview of all VBA modules without extracting code
- **Analyze Structure** - Code complexity metrics, procedure analysis, dependencies

## Installation

```bash
pip install vba-mcp-server
```

## Configuration

Add to your Claude Code MCP settings:

```json
{
  "mcpServers": {
    "vba": {
      "command": "vba-mcp-server"
    }
  }
}
```

Or run directly:

```json
{
  "mcpServers": {
    "vba": {
      "command": "python",
      "args": ["-m", "vba_mcp_lite.server"]
    }
  }
}
```

## Usage with Claude Code

Once configured, ask Claude:

- "Extract VBA from budget.xlsm"
- "List all modules in report.xlsm"
- "Analyze the structure of my Excel macros"

## Supported Formats

| Format | Description | Status |
|--------|-------------|--------|
| `.xlsm` | Excel Macro-Enabled | Supported |
| `.xlsb` | Excel Binary | Supported |
| `.docm` | Word Macro-Enabled | Supported |
| `.accdb` | Access Database | Supported |
| `.pptm` | PowerPoint Macro-Enabled | Supported |

## Pro Version

Need more features? Check out [VBA MCP Server Pro](https://github.com/AlexisTrouve/vba-mcp-server-pro):

- Modify and inject VBA code back into Office files
- AI-powered refactoring suggestions
- Testing framework for VBA
- And more...

## License

MIT License - See LICENSE file.

## Author

Alexis Trouve - [GitHub](https://github.com/AlexisTrouve)
