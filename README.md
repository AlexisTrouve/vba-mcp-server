# VBA MCP Server - Monorepo

Monorepo containing all VBA MCP Server packages.

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

## Development Setup

```bash
# Create virtual environment
python -m venv venv
source venv/bin/activate  # or venv\Scripts\activate on Windows

# Install all packages in editable mode
pip install -e packages/core
pip install -e packages/lite
pip install -e packages/pro  # Optional, for pro development

# Install dev dependencies
pip install pytest pytest-asyncio
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
