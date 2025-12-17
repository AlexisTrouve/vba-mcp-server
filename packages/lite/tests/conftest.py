"""Pytest configuration for lite package tests."""
import sys
from pathlib import Path

import pytest

# Add src to path
lite_src = Path(__file__).parent.parent / "src"
core_src = Path(__file__).parent.parent.parent / "core" / "src"
sys.path.insert(0, str(lite_src))
sys.path.insert(0, str(core_src))

# Path to examples directory
EXAMPLES_DIR = Path(__file__).parent.parent.parent.parent / "examples"


@pytest.fixture
def sample_xlsm():
    """Path to sample Excel file."""
    path = EXAMPLES_DIR / "sample.xlsm"
    if not path.exists():
        pytest.skip(f"Sample file not found: {path}")
    return path
