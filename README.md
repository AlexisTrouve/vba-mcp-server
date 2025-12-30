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

### Pro (Commercial) - v0.6.0 ✅ Production Ready
- All Lite features
- **Inject VBA code** back into Office files (Windows + COM) - **100% operational** ✅
- **VBA code validation** - Validate VBA syntax before injection with smart detection
- **List macros** - Discover all public macros in Office files
- **AI-powered refactoring** suggestions
- **Backup management** (create, restore, list backups) with automatic rollback
- **Interactive Office automation**
  - Open Excel/Word/Access visibly on screen
  - Run VBA macros with parameters (improved error handling)
  - Read/write Excel data as JSON
  - Persistent sessions (files stay open between operations)
  - Auto-cleanup after 1-hour timeout
- **Excel Tables** - 6 tools for structured table operations
  - List all Excel Tables (ListObjects)
  - Insert/delete rows in worksheets or tables
  - Insert/delete columns by letter, number, or header name
  - Create Excel Tables from ranges with formatting
  - Read/write table data with column selection
- **Microsoft Access Support (NEW v0.6.0)** - Full Access database integration
  - Read/write data from Access tables with SQL support
  - Filter data with WHERE clauses and ORDER BY
  - List all tables with schema (fields, types, record counts)
  - List and execute saved queries (QueryDefs)
  - Run custom SQL queries directly
  - VBA injection and validation for Access
- **Enhanced reliability** - Automatic backups, validation, and rollback on errors

**Total: 24 MCP tools** (18 Pro + 6 Lite)

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

## MCP Tools Reference

### Lite Tools (6)
| Tool | Description |
|------|-------------|
| `extract_vba` | Extract VBA code from modules |
| `list_modules` | List all VBA modules |
| `analyze_structure` | Analyze code structure |

### Pro Tools (18)
| Tool | Description |
|------|-------------|
| `inject_vba` | Inject VBA code into modules |
| `validate_vba` | Validate VBA syntax |
| `run_macro` | Execute VBA macros |
| `open_in_office` | Open file in Office application |
| `get_worksheet_data` | Read data from Excel/Access |
| `set_worksheet_data` | Write data to Excel/Access |
| `list_excel_tables` | List Excel Tables |
| `create_excel_table` | Create new Excel Table |
| `insert_rows` | Insert rows in worksheet/table |
| `delete_rows` | Delete rows |
| `insert_columns` | Insert columns |
| `delete_columns` | Delete columns |
| `create_backup` | Create file backup |
| `restore_backup` | Restore from backup |
| `list_backups` | List available backups |
| `list_access_tables` | List Access tables with schema |
| `list_access_queries` | List saved queries |
| `run_access_query` | Execute SQL/saved queries |

## Testing

```bash
# Run all tests
pytest

# Run specific tests
pytest packages/pro/tests/
python test_access_complete.py
```

## Known Limitations

1. **Windows only** - VBA injection requires pywin32 + COM
2. **Trust VBA** - Enable "Trust access to VBA project object model" in Office
3. **oletools** - Doesn't support .accdb for module listing (use COM workaround)

## Roadmap

See **[TODO.md](TODO.md)** for planned features.

## Author

Alexis Trouve - alexistrouve.pro@gmail.com
