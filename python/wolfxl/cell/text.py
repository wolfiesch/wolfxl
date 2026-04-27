"""``openpyxl.cell.text``-shaped re-exports.

openpyxl exposes the rich-text classes at *both* ``openpyxl.cell.rich_text``
and ``openpyxl.cell.text``.  Wolfxl mirrors the former at
:mod:`wolfxl.cell.rich_text`; this module is the openpyxl-shaped alias so
``from openpyxl.cell.text import CellRichText`` swaps to wolfxl mechanically.

Pod 2 (RFC-060) — re-export shim only.
"""

from __future__ import annotations

from wolfxl.cell.rich_text import CellRichText, InlineFont, TextBlock

__all__ = ["CellRichText", "InlineFont", "TextBlock"]
