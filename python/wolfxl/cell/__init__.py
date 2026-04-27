"""openpyxl-compatible ``wolfxl.cell`` namespace.

Re-exports the rich-text shims from :mod:`wolfxl.cell.rich_text` and
the array / data-table formula shims from :mod:`wolfxl.cell.cell`
(RFC-057 — Pod 1C).  The package shape mirrors openpyxl's
``openpyxl.cell`` so existing code that imports
``from openpyxl.cell.rich_text import CellRichText`` can be redirected
to ``wolfxl.cell.rich_text`` with a one-line change.
"""

from wolfxl.cell import cell, rich_text
from wolfxl.cell.cell import ArrayFormula, DataTableFormula

__all__ = ["ArrayFormula", "DataTableFormula", "cell", "rich_text"]
