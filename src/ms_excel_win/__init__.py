"""
Core package for the MainSequence Excel integration built on xlOil.
"""

__version__ = "0.1.0"

# Import the Excel/xlOil wrappers so ribbon and functions are registered on load.
from . import ms_excel_wrappers as _ms_excel_wrappers  # noqa: F401
