# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Planned
- Support for .xlsb (Excel Binary Workbook)
- Support for .accdb (Access Database)
- Support for .docm (Word Macro-Enabled Document)
- CI/CD pipeline with GitHub Actions
- Code coverage reporting
- Performance optimizations for large files

## [1.0.0] - 2025-12-11

### Added
- Initial release of VBA MCP Server
- MCP server implementation with stdio transport
- Three MCP tools:
  - `extract_vba` - Extract VBA source code from Office files
  - `list_modules` - List all VBA modules without extracting code
  - `analyze_structure` - Analyze VBA code structure and complexity
- Support for .xlsm (Excel Macro-Enabled Workbook) format
- VBA parser with procedure detection (Subs, Functions, Properties)
- Cyclomatic complexity calculation
- Office file handler using oletools library
- Complete test suite with 29 unit tests
- Comprehensive documentation:
  - README.md with quick start guide
  - API.md with tool reference
  - ARCHITECTURE.md with technical details
  - EXAMPLES.md with usage examples
  - ROADMAP.md with future plans
- Example VBA code for testing
- Test automation script (create_test_excel.py)
- Local testing script (test_local.py)
- MIT License for open source version
- Python 3.8+ support

### Technical Details
- Python libraries: mcp>=0.9.0, oletools>=0.60, openpyxl>=3.1.0
- Development tools: pytest, black, flake8, mypy
- Code quality: PEP 8 compliant, type hints, docstrings
- Test coverage: 100% of core functionality

### Documentation
- Complete README with installation instructions
- API reference for all MCP tools
- Architecture documentation
- Code examples and use cases
- Project roadmap
- Contributing guidelines

## [0.1.0] - 2025-12-10 (Internal)

### Added
- Initial project structure
- Basic VBA extraction prototype
- Proof of concept with oletools

---

## Version History Summary

- **v1.0.0** (2025-12-11) - First public release (Lite version)
- **v0.1.0** (2025-12-10) - Internal prototype

---

## Notes

### Version Numbering

We use Semantic Versioning:
- **MAJOR** version for incompatible API changes
- **MINOR** version for new functionality (backwards-compatible)
- **PATCH** version for bug fixes (backwards-compatible)

### Lite vs Pro

This changelog tracks the **Lite (open source)** version only.

The **Pro version** (commercial, closed source) includes additional features:
- VBA code modification and reinjection
- Automated refactoring with AI
- Macro execution (sandboxed)
- Version control integration
- Advanced testing framework

For Pro version changelog, contact: alexistrouve.pro@gmail.com

---

[Unreleased]: https://github.com/AlexisTrouve/vba-mcp-server/compare/v1.0.0...HEAD
[1.0.0]: https://github.com/AlexisTrouve/vba-mcp-server/releases/tag/v1.0.0
[0.1.0]: https://github.com/AlexisTrouve/vba-mcp-server/releases/tag/v0.1.0
