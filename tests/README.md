# VBA MCP Tests

This directory contains integration tests for the VBA MCP monorepo. Each package also has its own unit tests in `packages/{package}/tests/`.

## Running Tests

### Quick Start

Run all unit tests (excluding integration tests):

```bash
python run_tests.py
```

### Run Tests by Package

```bash
# Core package
pytest packages/core/tests -v -m "not integration"

# Lite package
pytest packages/lite/tests -v -m "not integration"

# Pro package
pytest packages/pro/tests -v -m "not integration"
```

### Run All Tests (Including Integration Tests)

```bash
# Requires sample.xlsm file in examples/
pytest packages/core/tests -v
pytest packages/lite/tests -v
pytest packages/pro/tests -v
```

### Run Windows-Only Tests

```bash
# Requires Windows + Microsoft Office installed
pytest packages/pro/tests -v -m "windows_only"
```

## Test Markers

Tests are organized with the following markers:

- `unit` - Unit tests (default)
- `integration` - Integration tests that require sample files
- `windows_only` - Tests requiring Windows + Office COM
- `slow` - Slow-running tests

## Test Structure

```
vba-mcp-monorepo/
├── packages/
│   ├── core/tests/       # Core library tests
│   ├── lite/tests/       # Lite MCP server tests
│   └── pro/tests/        # Pro server tests
├── tests/                # Integration tests (this directory)
├── examples/             # Sample files for testing
│   └── sample.xlsm      # Created with examples/create_sample.py
└── run_tests.py          # Test runner script
```

## Coverage

Current test coverage:

- **Core Package**: 16 unit tests
  - OfficeHandler: File handling, VBA extraction
  - VBAParser: Code parsing, procedure extraction
  - Tools: extract_vba, list_modules, analyze_structure

- **Lite Package**: 4 unit tests
  - MCP Server: Tool registration, call handling
  - Error handling

- **Pro Package**: 14 unit tests
  - Backup management: Create, list, restore, delete
  - VBA injection: Platform checks, file validation

## Adding New Tests

1. Create test file in appropriate package:
   ```
   packages/{package}/tests/test_feature.py
   ```

2. Use fixtures from conftest.py:
   ```python
   def test_something(sample_xlsm):
       # sample_xlsm is a pytest fixture
       assert sample_xlsm.exists()
   ```

3. Mark tests appropriately:
   ```python
   @pytest.mark.integration
   @pytest.mark.asyncio
   async def test_integration():
       pass
   ```

## Troubleshooting

### Sample File Missing

If tests fail with "Sample file not found":

```bash
cd examples
python create_sample.py
```

This requires Windows + Excel installed.

### Import Errors

Install development dependencies:

```bash
pip install pytest pytest-asyncio oletools mcp
```

Or install packages in editable mode:

```bash
pip install -e packages/core
pip install -e packages/lite
pip install -e packages/pro[windows]
```
