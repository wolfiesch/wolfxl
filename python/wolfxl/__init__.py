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

import os

from wolfxl._rust import __version__
from wolfxl._styles import Alignment, Border, Color, Font, PatternFill, Side
from wolfxl._workbook import Workbook

__all__ = [
    "__version__",
    "Alignment",
    "Border",
    "Color",
    "Font",
    "PatternFill",
    "Side",
    "Workbook",
    "load_workbook",
]


def load_workbook(
    filename: str | os.PathLike[str],
    read_only: bool = False,
    data_only: bool = False,
    keep_links: bool = True,
    modify: bool = False,
) -> Workbook:
    """Open an .xlsx file for reading or modification.

    Parameters
    ----------
    modify : bool
        If True, enable read-modify-write mode.  Values and formats can be
        changed and saved back to disk via ``wb.save(path)``.  Uses the WolfXL
        engine (surgical ZIP patching) instead of a full DOM rewrite.

    Extra keyword arguments (``read_only``, ``data_only``, ``keep_links``) are
    accepted for openpyxl compatibility but currently ignored.
    """
    if modify:
        return Workbook._from_patcher(str(filename))  # noqa: SLF001
    return Workbook._from_reader(str(filename))  # noqa: SLF001
