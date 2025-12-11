# Architecture

## Overview

The VBA MCP Server is a Python-based MCP (Model Context Protocol) server that enables AI assistants like Claude Code to interact with VBA macros in Microsoft Office files.

## System Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                        Claude Code                          │
│                    (MCP Client)                             │
└──────────────────────┬──────────────────────────────────────┘
                       │ MCP Protocol (stdio/HTTP)
                       │
┌──────────────────────▼──────────────────────────────────────┐
│                   VBA MCP Server                            │
│  ┌──────────────────────────────────────────────────────┐   │
│  │              MCP Server Core                         │   │
│  │  - Tool registration                                 │   │
│  │  - Request handling                                  │   │
│  │  - Response formatting                               │   │
│  └───────────┬──────────────────────────────────────────┘   │
│              │                                              │
│  ┌───────────▼──────────────────────────────────────────┐   │
│  │              Tool Layer                              │   │
│  │  ┌──────────┐  ┌──────────┐  ┌──────────┐           │   │
│  │  │ Extract  │  │   List   │  │ Analyze  │           │   │
│  │  │   VBA    │  │ Modules  │  │   Code   │           │   │
│  │  └────┬─────┘  └────┬─────┘  └────┬─────┘           │   │
│  └───────┼─────────────┼─────────────┼──────────────────┘   │
│          │             │             │                      │
│  ┌───────▼─────────────▼─────────────▼──────────────────┐   │
│  │              Business Logic Layer                    │   │
│  │  ┌──────────────────┐  ┌──────────────────┐          │   │
│  │  │ Office Handler   │  │   VBA Parser     │          │   │
│  │  │ - File opening   │  │ - AST parsing    │          │   │
│  │  │ - Format detect  │  │ - Code analysis  │          │   │
│  │  │ - Binary parsing │  │ - Dependencies   │          │   │
│  │  └────────┬─────────┘  └────────┬─────────┘          │   │
│  └───────────┼─────────────────────┼────────────────────┘   │
│              │                     │                        │
│  ┌───────────▼─────────────────────▼────────────────────┐   │
│  │           External Libraries                         │   │
│  │  - oletools (OLE2 extraction)                        │   │
│  │  - openpyxl (OOXML parsing)                          │   │
│  │  - zipfile (Archive handling)                        │   │
│  └──────────────────────────────────────────────────────┘   │
└──────────────────────┬──────────────────────────────────────┘
                       │
┌──────────────────────▼──────────────────────────────────────┐
│               Office Files (.xlsm, .accdb, .docm)           │
└─────────────────────────────────────────────────────────────┘
```

## Component Details

### 1. MCP Server Core

**Responsibilities:**
- Initialize MCP server with stdio/HTTP transport
- Register available tools
- Handle incoming requests from Claude Code
- Format and return responses

**Key Files:**
- `src/server.py` - Main entry point

**MCP Protocol:**
```json
{
  "jsonrpc": "2.0",
  "method": "tools/call",
  "params": {
    "name": "extract_vba",
    "arguments": {
      "file_path": "/path/to/file.xlsm",
      "module_name": "Module1"
    }
  }
}
```

### 2. Tool Layer

Each tool is a callable function exposed via MCP protocol.

#### Tool: `extract_vba`

**Purpose:** Extract VBA source code from Office files

**Input:**
```json
{
  "file_path": "string",
  "module_name": "string (optional)"
}
```

**Output:**
```json
{
  "modules": [
    {
      "name": "Module1",
      "type": "standard",
      "code": "Sub HelloWorld()...",
      "procedures": ["HelloWorld", "Calculate"]
    }
  ],
  "file_info": {
    "format": "xlsm",
    "size": 45678
  }
}
```

#### Tool: `list_modules`

**Purpose:** List all VBA modules without extracting code

**Input:**
```json
{
  "file_path": "string"
}
```

**Output:**
```json
{
  "modules": [
    {"name": "Module1", "type": "standard", "line_count": 150},
    {"name": "Sheet1", "type": "worksheet", "line_count": 45},
    {"name": "ThisWorkbook", "type": "workbook", "line_count": 30}
  ],
  "total_count": 3
}
```

#### Tool: `analyze_structure`

**Purpose:** Analyze VBA code structure and dependencies

**Input:**
```json
{
  "file_path": "string"
}
```

**Output:**
```json
{
  "procedures": [
    {
      "name": "HelloWorld",
      "type": "Sub",
      "module": "Module1",
      "calls": ["MsgBox"],
      "variables": ["result"]
    }
  ],
  "dependencies": {
    "Module1": ["Module2"],
    "Module2": []
  }
}
```

### 3. Business Logic Layer

#### Office Handler (`lib/office_handler.py`)

**Responsibilities:**
- Detect Office file format
- Open and parse binary structures
- Extract VBA project containers

**Supported Formats:**

| Format | Structure | Library |
|--------|-----------|---------|
| .xlsm/.docm | ZIP (OOXML) | openpyxl + zipfile |
| .xlsb | Binary OOXML | openpyxl |
| .accdb | JET/ACE database | pyodbc (Windows) |
| .xls/.mdb | OLE2 compound | oletools |

**Key Methods:**
```python
class OfficeHandler:
    def open_file(file_path: str) -> OfficeFile
    def detect_format(file_path: str) -> FileFormat
    def extract_vba_project(office_file: OfficeFile) -> VBAProject
```

#### VBA Parser (`lib/vba_parser.py`)

**Responsibilities:**
- Parse VBA source code
- Build abstract syntax tree (AST)
- Extract procedures, functions, variables
- Analyze dependencies

**Key Methods:**
```python
class VBAParser:
    def parse(vba_code: str) -> VBAModule
    def extract_procedures(module: VBAModule) -> List[Procedure]
    def find_dependencies(module: VBAModule) -> List[str]
```

### 4. External Libraries

#### oletools
- **Purpose:** Extract VBA from OLE2 compound files
- **Use case:** Legacy .xls, .mdb files
- **Key module:** `oletools.olevba`

#### openpyxl
- **Purpose:** Parse modern Excel files
- **Use case:** .xlsx, .xlsm (OOXML)
- **Key module:** `openpyxl.reader.excel`

#### zipfile
- **Purpose:** Unpack OOXML archives
- **Use case:** Extract vbaProject.bin from .xlsm
- **Built-in:** Python standard library

## Data Flow

### Extraction Flow

```
1. Claude Code request
   ↓
2. MCP Server receives "extract_vba" call
   ↓
3. OfficeHandler.open_file(path)
   ├─ Detect format (.xlsm)
   ├─ Unzip archive
   └─ Locate xl/vbaProject.bin
   ↓
4. oletools.olevba.VBA_Parser(vbaProject.bin)
   ├─ Parse OLE2 streams
   └─ Extract VBA source
   ↓
5. VBAParser.parse(source_code)
   ├─ Tokenize
   ├─ Build AST
   └─ Extract metadata
   ↓
6. Format JSON response
   ↓
7. Return to Claude Code
```

### Error Handling

```python
try:
    office_file = handler.open_file(path)
    vba_project = handler.extract_vba_project(office_file)
except FileNotFoundError:
    return {"error": "File not found", "code": "FILE_NOT_FOUND"}
except PasswordProtected:
    return {"error": "File is password protected", "code": "PASSWORD_REQUIRED"}
except NoVBAFound:
    return {"error": "No VBA macros found", "code": "NO_VBA"}
except Exception as e:
    return {"error": str(e), "code": "INTERNAL_ERROR"}
```

## Security Considerations

### Sandboxing

- **Never execute** VBA code in lite version
- Read-only file access
- No write operations to Office files

### Input Validation

```python
def validate_file_path(path: str):
    # Check file exists
    if not os.path.exists(path):
        raise FileNotFoundError()

    # Check file extension
    allowed = ['.xlsm', '.xlsb', '.accdb', '.docm']
    if not any(path.endswith(ext) for ext in allowed):
        raise ValueError("Unsupported file format")

    # Check file size (prevent DOS)
    if os.path.getsize(path) > 100_000_000:  # 100MB
        raise ValueError("File too large")
```

### Malware Protection

⚠️ VBA can contain malicious macros:
- Recommend antivirus scan before processing
- Log all file accesses
- Never auto-execute macros

## Performance

### Optimization Strategies

1. **Lazy Loading:** Only extract requested modules
2. **Caching:** Cache parsed results for repeated access
3. **Streaming:** Process large files in chunks

### Benchmarks

| File Size | Modules | Extraction Time |
|-----------|---------|-----------------|
| 100 KB | 5 | ~0.2s |
| 1 MB | 20 | ~1.5s |
| 10 MB | 100 | ~8s |

## Testing Strategy

### Unit Tests

```python
# tests/test_extract.py
def test_extract_xlsm():
    result = extract_vba("examples/sample.xlsm")
    assert len(result["modules"]) > 0
    assert result["modules"][0]["name"] == "Module1"
```

### Integration Tests

```python
# tests/test_mcp_server.py
def test_mcp_tool_call():
    response = mcp_client.call_tool("extract_vba", {
        "file_path": "test.xlsm"
    })
    assert response["status"] == "success"
```

## Deployment

### Local Development

```bash
python src/server.py --transport stdio
```

### Production (HTTP)

```bash
python src/server.py --transport http --port 8080
```

### Docker (Future)

```dockerfile
FROM python:3.11-slim
COPY . /app
RUN pip install -r requirements.txt
CMD ["python", "src/server.py"]
```

## Future Enhancements

### Lite Version
- [ ] Cross-platform support (Windows/Mac/Linux)
- [ ] Support for .xls, .mdb (legacy formats)
- [ ] VBA syntax highlighting in responses
- [ ] Code complexity metrics

### Pro Version
- [ ] VBA modification and reinjection
- [ ] Automated refactoring engine
- [ ] Macro execution with sandboxing
- [ ] Version control integration
- [ ] Testing framework (VBA unit tests)

## References

- [Model Context Protocol Spec](https://modelcontextprotocol.io)
- [oletools Documentation](https://github.com/decalage2/oletools)
- [MS-OVBA Specification](https://docs.microsoft.com/en-us/openspecs/office_file_formats/)
