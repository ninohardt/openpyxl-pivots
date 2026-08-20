"""Experimental PivotTable creation for openpyxl."""

from .builder import PivotBuildError, add_pivot_table

__all__ = ["PivotBuildError", "add_pivot_table", "install"]
__version__ = "0.2.1"


def install():
    """Install ``Worksheet.add_pivot_table`` as a convenience method."""
    from openpyxl.worksheet.worksheet import Worksheet

    if not hasattr(Worksheet, "add_pivot_table"):
        Worksheet.add_pivot_table = add_pivot_table
