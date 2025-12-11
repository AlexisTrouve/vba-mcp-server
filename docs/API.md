# API Reference

This document describes all MCP tools exposed by the VBA MCP Server.

## Available Tools

### 1. extract_vba

Extract VBA source code from Office files.

**Tool Name:** `extract_vba`

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `file_path` | string | Yes | Absolute path to Office file |
| `module_name` | string | No | Specific module to extract (if omitted, extracts all) |

**Example Request:**

```json
{
  "name": "extract_vba",
  "arguments": {
    "file_path": "/Users/john/Documents/budget.xlsm",
    "module_name": "Module1"
  }
}
```

**Example Response:**

```json
{
  "status": "success",
  "modules": [
    {
      "name": "Module1",
      "type": "standard",
      "code": "Sub HelloWorld()\n    MsgBox \"Hello!\"\nEnd Sub",
      "procedures": [
        {
          "name": "HelloWorld",
          "type": "Sub",
          "line_start": 1,
          "line_end": 3
        }
      ],
      "line_count": 3
    }
  ],
  "file_info": {
    "path": "/Users/john/Documents/budget.xlsm",
    "format": "xlsm",
    "size_bytes": 45678,
    "modified": "2025-01-15T10:30:00Z"
  }
}
```

**Error Responses:**

```json
{
  "status": "error",
  "error": "File not found",
  "code": "FILE_NOT_FOUND"
}
```

**Error Codes:**

| Code | Description |
|------|-------------|
| `FILE_NOT_FOUND` | File does not exist |
| `UNSUPPORTED_FORMAT` | File format not supported |
| `NO_VBA` | No VBA macros found in file |
| `PASSWORD_PROTECTED` | File is password protected |
| `MODULE_NOT_FOUND` | Requested module doesn't exist |
| `INTERNAL_ERROR` | Unexpected error during extraction |

---

### 2. list_modules

List all VBA modules in an Office file without extracting code.

**Tool Name:** `list_modules`

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `file_path` | string | Yes | Absolute path to Office file |

**Example Request:**

```json
{
  "name": "list_modules",
  "arguments": {
    "file_path": "/Users/john/Documents/budget.xlsm"
  }
}
```

**Example Response:**

```json
{
  "status": "success",
  "modules": [
    {
      "name": "Module1",
      "type": "standard",
      "line_count": 150,
      "procedures_count": 5
    },
    {
      "name": "Sheet1",
      "type": "worksheet",
      "line_count": 45,
      "procedures_count": 2
    },
    {
      "name": "ThisWorkbook",
      "type": "workbook",
      "line_count": 30,
      "procedures_count": 3
    }
  ],
  "total_modules": 3,
  "total_lines": 225,
  "file_info": {
    "path": "/Users/john/Documents/budget.xlsm",
    "format": "xlsm"
  }
}
```

**Module Types:**

| Type | Description |
|------|-------------|
| `standard` | Standard module (.bas) |
| `class` | Class module (.cls) |
| `worksheet` | Worksheet code-behind |
| `workbook` | Workbook code-behind |
| `form` | UserForm module |

---

### 3. analyze_structure

Analyze VBA code structure, dependencies, and complexity.

**Tool Name:** `analyze_structure`

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `file_path` | string | Yes | Absolute path to Office file |
| `module_name` | string | No | Analyze specific module only |

**Example Request:**

```json
{
  "name": "analyze_structure",
  "arguments": {
    "file_path": "/Users/john/Documents/budget.xlsm"
  }
}
```

**Example Response:**

```json
{
  "status": "success",
  "analysis": {
    "procedures": [
      {
        "name": "CalculateTotal",
        "type": "Function",
        "module": "Module1",
        "visibility": "Public",
        "parameters": ["amount", "tax"],
        "return_type": "Double",
        "calls": ["MsgBox", "Format"],
        "called_by": ["GenerateReport"],
        "variables": ["result", "finalAmount"],
        "line_count": 15,
        "complexity": 3
      },
      {
        "name": "GenerateReport",
        "type": "Sub",
        "module": "Module1",
        "visibility": "Public",
        "parameters": [],
        "calls": ["CalculateTotal", "Worksheets"],
        "called_by": [],
        "variables": ["ws", "total"],
        "line_count": 25,
        "complexity": 5
      }
    ],
    "dependencies": {
      "Module1": [],
      "Sheet1": ["Module1"]
    },
    "external_references": [
      {
        "name": "Scripting.FileSystemObject",
        "type": "COM",
        "guid": "{420B2830-E718-11CF-893D-00A0C9054228}"
      }
    ],
    "metrics": {
      "total_procedures": 2,
      "total_lines": 40,
      "avg_complexity": 4,
      "max_complexity": 5
    }
  }
}
```

**Complexity Metrics:**

Cyclomatic complexity is calculated based on:
- Decision points (If, Select Case, loops)
- Boolean operators (And, Or)
- Error handlers (On Error)

| Complexity | Quality | Recommendation |
|------------|---------|----------------|
| 1-5 | Good | No action needed |
| 6-10 | Moderate | Consider refactoring |
| 11-20 | High | Refactor recommended |
| 21+ | Very High | Refactor required |

---

## Common Usage Patterns

### Pattern 1: Extract all modules

```
User: "Extract all VBA code from budget.xlsm"

Claude Code calls:
{
  "name": "extract_vba",
  "arguments": {
    "file_path": "C:/Users/john/Documents/budget.xlsm"
  }
}
```

### Pattern 2: List modules first, then extract specific one

```
User: "What modules are in my Excel file?"

Claude Code calls list_modules
Response shows: Module1, Sheet1, ThisWorkbook

User: "Show me Module1"

Claude Code calls:
{
  "name": "extract_vba",
  "arguments": {
    "file_path": "...",
    "module_name": "Module1"
  }
}
```

### Pattern 3: Analyze before refactoring

```
User: "Analyze the VBA structure in report.xlsm"

Claude Code calls analyze_structure
Response shows complexity = 15

User: "That's too complex, can you help refactor?"

Claude uses the analysis + extracted code to suggest refactoring
```

## Response Format

All tool responses follow this structure:

```json
{
  "status": "success" | "error",
  "data": { ... },           // Only if success
  "error": "string",          // Only if error
  "code": "ERROR_CODE"        // Only if error
}
```

## Rate Limiting

**Lite Version:**
- No rate limiting
- Local execution only

**Pro Version:**
- HTTP mode: 100 requests/minute
- Concurrent requests: 5

## File Size Limits

| Version | Max File Size |
|---------|---------------|
| Lite | 100 MB |
| Pro | 500 MB |

## Supported Office Versions

| Application | Versions | Notes |
|-------------|----------|-------|
| Excel | 2007+ | .xlsm, .xlsb |
| Access | 2007+ | .accdb |
| Word | 2007+ | .docm |
| PowerPoint | 2007+ | .pptm (planned) |

## Error Handling Best Practices

```python
# Example: Handle common errors gracefully

try:
    result = mcp.call_tool("extract_vba", {
        "file_path": file_path
    })

    if result["status"] == "error":
        if result["code"] == "NO_VBA":
            print("This file has no macros")
        elif result["code"] == "PASSWORD_PROTECTED":
            print("Please unlock the file first")
        else:
            print(f"Error: {result['error']}")
except Exception as e:
    print(f"Unexpected error: {e}")
```

## Performance Characteristics

| Operation | Small File (<1MB) | Medium (1-10MB) | Large (10-100MB) |
|-----------|-------------------|-----------------|------------------|
| `list_modules` | <100ms | <500ms | 1-3s |
| `extract_vba` | <200ms | 1-2s | 5-10s |
| `analyze_structure` | <500ms | 2-5s | 10-30s |

## Examples

See [EXAMPLES.md](EXAMPLES.md) for detailed usage examples.

## Changelog

### Version 1.0.0 (Lite)
- Initial release
- Basic extraction for .xlsm files
- Module listing
- Structure analysis

### Planned (Pro)
- VBA modification
- Code reinjection
- Macro execution
- Advanced refactoring

## Support

For issues or questions:
- GitHub Issues: https://github.com/AlexisTrouve/vba-mcp-server/issues
- Email: alexistrouve.pro@gmail.com
