"""
VBA MCP Core - Shared functionality for VBA extraction and analysis.

This package contains the core logic used by both lite and pro versions.
"""

__version__ = "0.1.0"

from .lib.office_handler import OfficeHandler
from .lib.vba_parser import VBAParser

__all__ = ["OfficeHandler", "VBAParser", "__version__"]
