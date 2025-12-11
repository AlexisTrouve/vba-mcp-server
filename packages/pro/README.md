# VBA MCP Server Pro

Professional MCP server for VBA extraction, analysis, and modification.

## Features

### Lite Features (included)
- Extract VBA code from Office files
- List VBA modules
- Analyze code structure and complexity

### Pro Features
- **Inject VBA** - Modify and inject VBA code back into Office files
- **Refactor** - AI-powered refactoring suggestions
- **Backup Management** - Create, restore, and manage backups

## Installation

```bash
pip install vba-mcp-server-pro
```

On Windows, for full VBA injection support:
```bash
pip install vba-mcp-server-pro[windows]
```

## Configuration

```json
{
  "mcpServers": {
    "vba-pro": {
      "command": "vba-mcp-server-pro"
    }
  }
}
```

## Usage

### Refactoring
"Suggest refactoring for my Excel macros"

### Backup Management
"Create a backup of budget.xlsm before I modify it"
"Restore budget.xlsm from yesterday's backup"

### Code Injection
"Update the CalculateTotal function in budget.xlsm with this new code..."

## License

Commercial License - Contact alexistrouve.pro@gmail.com for licensing.

## Support

Email: alexistrouve.pro@gmail.com
