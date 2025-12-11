# VBA MCP Server - Monorepo

Model Context Protocol (MCP) servers for VBA extraction, analysis, and modification in Microsoft Office files.

> **MCP** enables Claude to interact with Office files through specialized tools. Perfect for automating VBA analysis, refactoring, and code injection.

## Features

### Lite (Free - MIT)
- Extract VBA code from Office files (.xlsm, .xlsb, .docm, .accdb)
- List all VBA modules and procedures
- Analyze code structure and complexity metrics

### Pro (Commercial)
- All Lite features
- **Inject VBA code** back into Office files (Windows + COM)
- **AI-powered refactoring** suggestions
- **Backup management** (create, restore, list backups)

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
