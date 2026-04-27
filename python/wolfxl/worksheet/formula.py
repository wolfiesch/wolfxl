"""``openpyxl.worksheet.formula`` — re-export for ArrayFormula / DataTableFormula.

Wolfxl's canonical home for these is :mod:`wolfxl.cell.cell` (RFC-057).
This module surfaces them at the openpyxl-shaped path.

Pod 2 (RFC-060 §2.1).
"""

from __future__ import annotations

from wolfxl.cell.cell import ArrayFormula, DataTableFormula

__all__ = ["ArrayFormula", "DataTableFormula"]
