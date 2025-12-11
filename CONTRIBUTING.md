# Contributing to VBA MCP Server

First off, thank you for considering contributing to VBA MCP Server! 🎉

This document provides guidelines for contributing to this project.

## Table of Contents

- [Code of Conduct](#code-of-conduct)
- [How Can I Contribute?](#how-can-i-contribute)
- [Development Setup](#development-setup)
- [Pull Request Process](#pull-request-process)
- [Coding Standards](#coding-standards)
- [Testing](#testing)
- [Documentation](#documentation)

---

## Code of Conduct

This project adheres to a code of conduct. By participating, you are expected to uphold this code:

- Be respectful and inclusive
- Accept constructive criticism gracefully
- Focus on what's best for the community
- Show empathy towards other community members

## How Can I Contribute?

### Reporting Bugs

Before creating bug reports, please check existing issues to avoid duplicates.

When creating a bug report, include:

- **Clear title** describing the issue
- **Steps to reproduce** the problem
- **Expected behavior** vs **actual behavior**
- **Environment details** (OS, Python version, Office version)
- **Code samples** or test files if applicable
- **Error messages** or stack traces

### Suggesting Enhancements

Enhancement suggestions are welcome! Please provide:

- **Clear use case** - Why is this enhancement useful?
- **Detailed description** - What should the feature do?
- **Examples** - How would users interact with it?
- **Alternatives** - Have you considered other solutions?

### Pull Requests

We actively welcome pull requests:

1. **Fork** the repository
2. **Create a branch** from `main`: `git checkout -b feature/my-feature`
3. **Make your changes** with clear commits
4. **Add tests** for new functionality
5. **Update documentation** as needed
6. **Run tests** to ensure nothing breaks
7. **Submit pull request** with clear description

---

## Development Setup

### Prerequisites

- Python 3.8 or higher
- pip and virtualenv
- Microsoft Office (for testing with real files)
- Git

### Installation

```bash
# Clone your fork
git clone https://github.com/YOUR_USERNAME/vba-mcp-server.git
cd vba-mcp-server

# Create virtual environment
python -m venv venv

# Activate virtual environment
# On Windows:
venv\Scripts\activate
# On macOS/Linux:
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt

# Install development dependencies
pip install -e .
```

### Running Tests

```bash
# Run all tests
pytest tests/ -v

# Run specific test file
pytest tests/test_vba_parser.py -v

# Run with coverage
pytest tests/ --cov=src --cov-report=html

# Run linters
black src/ tests/
flake8 src/ tests/
mypy src/
```

---

## Pull Request Process

1. **Update documentation** - Ensure README, API docs, and docstrings are updated
2. **Add tests** - New features must have corresponding tests
3. **Pass all tests** - All 29+ tests must pass
4. **Follow style guide** - Code must be PEP 8 compliant
5. **Update CHANGELOG.md** - Add your changes to Unreleased section
6. **One feature per PR** - Keep pull requests focused
7. **Clear commit messages** - Use descriptive commit messages

### Commit Message Format

```
<type>: <short description>

<optional detailed description>

<optional footer>
```

Types:
- `feat:` - New feature
- `fix:` - Bug fix
- `docs:` - Documentation changes
- `test:` - Adding or updating tests
- `refactor:` - Code refactoring
- `style:` - Code style changes (formatting)
- `chore:` - Maintenance tasks

Example:
```
feat: add support for .accdb files

Implement Access database VBA extraction using oletools.
Includes tests and documentation updates.

Closes #42
```

---

## Coding Standards

### Python Style Guide

- Follow **PEP 8** style guide
- Use **type hints** for function parameters and return values
- Maximum line length: **88 characters** (Black default)
- Use **docstrings** for all public functions and classes

### Code Quality Tools

We use these tools (run before committing):

```bash
# Format code
black src/ tests/

# Check style
flake8 src/ tests/

# Type checking
mypy src/

# All at once
black src/ tests/ && flake8 src/ tests/ && mypy src/
```

### Docstring Format

Use Google-style docstrings:

```python
def extract_vba(file_path: str, module_name: str = None) -> dict:
    """
    Extract VBA code from an Office file.

    Args:
        file_path: Absolute path to the Office file
        module_name: Optional specific module to extract

    Returns:
        Dictionary containing extracted VBA code and metadata

    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If file format is unsupported
    """
    pass
```

---

## Testing

### Writing Tests

- **Test file naming**: `test_<module_name>.py`
- **Test function naming**: `test_<feature>_<scenario>`
- **Use fixtures** for reusable test data
- **Mock external dependencies** (file system, Office APIs)
- **Test edge cases** (empty files, invalid formats, etc.)

### Test Structure

```python
import pytest

class TestFeature:
    """Test suite for specific feature"""

    @pytest.fixture
    def sample_data(self):
        """Fixture providing test data"""
        return {"key": "value"}

    def test_normal_case(self, sample_data):
        """Test normal operation"""
        result = my_function(sample_data)
        assert result == expected

    def test_error_case(self):
        """Test error handling"""
        with pytest.raises(ValueError):
            my_function(invalid_input)
```

### Test Coverage

- Aim for **>80% code coverage**
- All new features must have tests
- Bug fixes should include regression tests

---

## Documentation

### What to Document

- **Public APIs** - All user-facing functions
- **Complex logic** - Non-obvious code sections
- **Configuration** - Setup and configuration options
- **Examples** - Usage examples for new features

### Documentation Files

Update these as needed:

- `README.md` - Project overview and quick start
- `docs/API.md` - API reference
- `docs/ARCHITECTURE.md` - Technical architecture
- `docs/EXAMPLES.md` - Usage examples
- `CHANGELOG.md` - Version history

---

## Project Structure

```
vba-mcp-server/
├── src/
│   ├── server.py              # MCP server entry point
│   ├── tools/                 # MCP tools
│   │   ├── extract.py
│   │   ├── list_modules.py
│   │   └── analyze.py
│   └── lib/                   # Core libraries
│       ├── office_handler.py
│       └── vba_parser.py
├── tests/                     # Test suite
├── docs/                      # Documentation
├── examples/                  # Example files
└── requirements.txt           # Dependencies
```

---

## Scope: Lite vs Pro

**Important**: This repository is for the **Lite (open source)** version only.

### Lite Version (this repo)
✅ VBA extraction (read-only)
✅ Code analysis
✅ Structure parsing
✅ Complexity metrics

### Pro Version (private repo)
❌ VBA modification
❌ Code reinjection
❌ Automated refactoring
❌ Macro execution

**Do NOT submit PRs for Pro features** - they will be rejected.

---

## Getting Help

- **Questions?** Open a GitHub issue with `question` label
- **Discussion?** Use GitHub Discussions
- **Email?** alexistrouve.pro@gmail.com

---

## Recognition

Contributors will be recognized in:
- README.md Contributors section
- Release notes
- CHANGELOG.md

Thank you for contributing! 🙏
