# VBA MCP Server - Monorepo

Model Context Protocol (MCP) servers for VBA extraction, analysis, and modification in Microsoft Office files.

> **MCP** enables Claude to interact with Office files through specialized tools. Perfect for automating VBA analysis, refactoring, and code injection.

## 🚀 Quick Start

**→ READ THIS FIRST:** [START_HERE.md](START_HERE.md)

Complete setup guide with 3 simple steps (5 minutes).

## Features

### Lite (Free - MIT)
- Extract VBA code from Office files (.xlsm, .xlsb, .docm, .accdb)
- List all VBA modules and procedures
- Analyze code structure and complexity metrics

### Pro (Commercial) - v0.3.0
- All Lite features
- **Inject VBA code** back into Office files (Windows + COM)
- **VBA code validation** - Validate VBA syntax before injection
- **List macros** - Discover all public macros in Office files
- **AI-powered refactoring** suggestions
- **Backup management** (create, restore, list backups)
- **Interactive Office automation**
  - Open Excel/Word/Access visibly on screen
  - Run VBA macros with parameters (improved error handling)
  - Read/write Excel data as JSON
  - Persistent sessions (files stay open between operations)
  - Auto-cleanup after 1-hour timeout
- **Excel Tables (NEW v0.3.0)** - 6 tools for structured table operations
  - List all Excel Tables (ListObjects)
  - Insert/delete rows in worksheets or tables
  - Insert/delete columns by letter, number, or header name
  - Create Excel Tables from ranges with formatting
  - Read/write table data with column selection
- **Enhanced reliability** - Automatic backups, validation, and rollback on errors

**Total: 21 MCP tools** (15 Pro + 6 Lite)

## Structure

```
vba-mcp-monorepo/
├── packages/
│   ├── core/       # vba-mcp-core - Shared library (MIT)
│   ├── lite/       # vba-mcp-server - Open source server (MIT)
│   └── pro/        # vba-mcp-server-pro - Commercial version
├── docs/           # Documentation
├── examples/       # Example files
└── tests/          # Integration tests
```

## Packages

| Package | Description | License |
|---------|-------------|---------|
| `vba-mcp-core` | Core extraction/parsing library | MIT |
| `vba-mcp-server` | Lite MCP server (read-only) | MIT |
| `vba-mcp-server-pro` | Pro server (modification features) | Commercial |

## Quick Start

### For Users

```bash
# Install Lite (free)
pip install vba-mcp-server

# Install Pro (commercial)
pip install vba-mcp-server-pro[windows]  # Windows only
```

### HTTP/SSE Transport (Cross-Platform)

**NEW**: Run the server on Windows and connect from any OS (WSL, macOS, Linux)!

```bash
# On Windows - Start HTTP server
vba-mcp-server-pro-http --host 0.0.0.0 --port 8000

# Configure client (any OS)
{
  "mcpServers": {
    "vba-pro": {
      "url": "http://YOUR_WINDOWS_IP:8000/sse"
    }
  }
}
```

**Perfect for:**
- 🐧 Connecting from WSL to Windows
- 🍎 Connecting from macOS to Windows server
- 🌐 Team setups with shared Windows server
- 📦 Docker containers accessing Windows host

See **[packages/pro/HTTP_SETUP.md](packages/pro/HTTP_SETUP.md)** for complete setup guide.

### For Developers

See **[DEVELOPMENT.md](DEVELOPMENT.md)** for complete setup instructions.

Quick install:
```bash
# Create virtual environment
python -m venv venv
source venv/bin/activate

# Install in editable mode
pip install -e packages/core
pip install -e packages/lite
pip install -e packages/pro[windows]  # Windows only
```

## Publishing Strategy

### Public Repo (GitHub)
Only `core/` and `lite/` are published publicly.

```bash
# Use public .gitignore (excludes packages/pro/)
cp .gitignore.public .gitignore
git add .
git commit -m "Release"
git push origin main
```

### Private Repo
Full monorepo including `pro/`.

```bash
# Use default .gitignore (includes everything)
git push private main
```

## Testing

```bash
# Run all tests
pytest

# Run specific package tests
pytest packages/core/tests/
pytest packages/lite/tests/
pytest packages/pro/tests/
```

## Author

Alexis Trouve - alexistrouve.pro@gmail.com
