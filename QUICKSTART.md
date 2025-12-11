# Quick Start Guide

Get up and running with VBA MCP Server in 5 minutes.

## Step 1: Install Dependencies

```bash
cd vba-mcp-server
pip install -r requirements.txt
```

## Step 2: Test the Server

Run a quick test to ensure everything works:

```bash
python src/server.py --test
```

## Step 3: Configure Claude Code

### Option A: Global Configuration

Edit your global Claude Code config (`~/.config/claude-code/mcp.json`):

```json
{
  "mcpServers": {
    "vba": {
      "command": "python",
      "args": ["/absolute/path/to/vba-mcp-server/src/server.py"]
    }
  }
}
```

### Option B: Project-Specific

Create `.mcp.json` in your project directory:

```json
{
  "mcpServers": {
    "vba": {
      "command": "python",
      "args": ["C:/path/to/vba-mcp-server/src/server.py"]
    }
  }
}
```

### Option C: Using Claude Code CLI

```bash
claude mcp add --transport stdio vba \
  --scope project \
  -- python /path/to/vba-mcp-server/src/server.py
```

## Step 4: Restart Claude Code

After configuration, restart Claude Code to load the MCP server.

## Step 5: Try It Out!

In Claude Code, try these commands:

### Extract VBA from a file

```
Extract the VBA code from C:/Users/john/Documents/budget.xlsm
```

### List modules

```
What VBA modules are in my Excel file?
```

### Analyze structure

```
Analyze the VBA structure in report.xlsm
```

## Troubleshooting

### Server not responding

1. Check server is configured correctly in `.mcp.json`
2. Verify Python path is correct
3. Check dependencies are installed: `pip list | grep -E "mcp|oletools|openpyxl"`

### "oletools not found" error

```bash
pip install oletools
```

### Permission errors

Ensure you have read permission for the Office files you're trying to access.

### File format not supported

Currently supported formats:
- .xlsm (Excel Macro-Enabled)
- .xlsb (Excel Binary)
- .accdb (Access Database)
- .docm (Word Macro-Enabled)

## Next Steps

- Read [API.md](docs/API.md) for detailed tool documentation
- See [EXAMPLES.md](docs/EXAMPLES.md) for usage examples
- Check [ARCHITECTURE.md](docs/ARCHITECTURE.md) for technical details

## Getting Help

- GitHub Issues: https://github.com/AlexisTrouve/vba-mcp-server/issues
- Email: alexistrouve.pro@gmail.com
