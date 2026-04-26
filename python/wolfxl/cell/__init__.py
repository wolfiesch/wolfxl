"""openpyxl-compatible ``wolfxl.cell`` namespace.

Currently re-exports the rich-text shims from
:mod:`wolfxl.cell.rich_text`.  The package shape mirrors openpyxl's
``openpyxl.cell`` so existing code that imports
``from openpyxl.cell.rich_text import CellRichText`` can be redirected
to ``wolfxl.cell.rich_text`` with a one-line change.
"""

from wolfxl.cell import rich_text

__all__ = ["rich_text"]
