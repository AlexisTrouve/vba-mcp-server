"""Pytest configuration and shared fixtures."""
import os
import sys
from pathlib import Path

import pytest

# Add packages to path
repo_root = Path(__file__).parent.parent
sys.path.insert(0, str(repo_root / "packages" / "core" / "src"))
sys.path.insert(0, str(repo_root / "packages" / "lite" / "src"))
sys.path.insert(0, str(repo_root / "packages" / "pro" / "src"))


@pytest.fixture
def examples_dir():
    """Path to examples directory."""
    return repo_root / "examples"


@pytest.fixture
def sample_xlsm(examples_dir):
    """Path to sample Excel file with macros."""
    path = examples_dir / "sample.xlsm"
    if not path.exists():
        pytest.skip(f"Sample file not found: {path}")
    return path


@pytest.fixture
def temp_backup_dir(tmp_path):
    """Temporary directory for backups."""
    backup_dir = tmp_path / "backups"
    backup_dir.mkdir()
    return backup_dir
