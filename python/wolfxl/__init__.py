"""WolfXL — fast, openpyxl-compatible Excel I/O backed by Rust.

Usage::

    from wolfxl import load_workbook, Workbook, Font, PatternFill

    # Read
    wb = load_workbook("data.xlsx")
    ws = wb["Sheet1"]
    print(ws["A1"].value, ws["A1"].font.bold)

    # Write
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "Hello"
    ws["A1"].font = Font(bold=True)
    wb.save("out.xlsx")
"""

from __future__ import annotations

import os

from wolfxl._cell import Cell
from wolfxl._rust import __version__, classify_format
from wolfxl._styles import Alignment, Border, Color, Font, PatternFill, Side
from wolfxl._workbook import CopyOptions, Workbook
from wolfxl._worksheet import Worksheet

__all__ = [
    "__version__",
    "Alignment",
    "Border",
    "Cell",
    "Color",
    "CopyOptions",
    "Font",
    "PatternFill",
    "Side",
    "Workbook",
    "Worksheet",
    "classify_format",
    "load_workbook",
]


def load_workbook(
    filename: str | os.PathLike[str],
    read_only: bool = False,
    data_only: bool = False,
    keep_links: bool = True,
    modify: bool = False,
    permissive: bool = False,
) -> Workbook:
    """Open an .xlsx file for reading or modification.

    Parameters
    ----------
    modify : bool
        If True, enable read-modify-write mode.  Values and formats can be
        changed and saved back to disk via ``wb.save(path)``.  Uses the WolfXL
        engine (surgical ZIP patching) instead of a full DOM rewrite.
    permissive : bool
        If True, fall back to the workbook rels graph when
        ``xl/workbook.xml``'s ``<sheets>`` block is empty or
        self-closing. Each worksheet relationship target is registered
        under a synthesized title (``Sheet1``, ``Sheet2``, ...). This
        unblocks workflows that need to operate on technically-malformed
        (but Excel-tolerant) workbooks — e.g. a self-closing
        ``<sheets/>`` whose rels still reference
        ``xl/worksheets/sheet1.xml``. Default is ``False`` so well-formed
        inputs round-trip unchanged. Added in Sprint Θ Pod-A; tracked in
        tests/parity/KNOWN_GAPS.md (RFC-035 cross-RFC composition bug
        #4).

    Extra keyword arguments (``read_only``, ``data_only``, ``keep_links``) are
    accepted for openpyxl compatibility. ``data_only=True`` returns cached
    formula results when they exist; ``read_only`` and ``keep_links`` remain
    no-op compatibility shims.
    """
    if modify:
        return Workbook._from_patcher(  # noqa: SLF001
            str(filename), data_only=data_only, permissive=permissive
        )
    return Workbook._from_reader(  # noqa: SLF001
        str(filename), data_only=data_only, permissive=permissive
    )
