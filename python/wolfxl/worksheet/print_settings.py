"""Print-settings classes (RFC-055 §2.4 — Pod 2 re-export targets).

Provides ``PrintArea``, ``PrintTitles``, ``ColRange``, ``RowRange``
mirroring openpyxl's ``openpyxl.worksheet.print_settings`` surface.
"""

from __future__ import annotations

import re
from dataclasses import dataclass


_ROW_RANGE_RE = re.compile(r"^\$?(\d+):\$?(\d+)$")
_COL_RANGE_RE = re.compile(r"^\$?([A-Za-z]+):\$?([A-Za-z]+)$")


@dataclass
class RowRange:
    """A row range like ``"1:2"`` (rows 1-2 inclusive, 1-based)."""

    min_row: int
    max_row: int

    @classmethod
    def from_string(cls, s: str) -> "RowRange":
        m = _ROW_RANGE_RE.match(s.strip())
        if not m:
            raise ValueError(f"invalid row range: {s!r}")
        a, b = int(m.group(1)), int(m.group(2))
        if a < 1 or b < 1:
            raise ValueError(f"row indices must be >=1: {s!r}")
        if a > b:
            a, b = b, a
        return cls(min_row=a, max_row=b)

    def __str__(self) -> str:
        return f"{self.min_row}:{self.max_row}"


@dataclass
class ColRange:
    """A column range like ``"A:B"`` (columns A-B inclusive)."""

    min_col: str
    max_col: str

    @classmethod
    def from_string(cls, s: str) -> "ColRange":
        m = _COL_RANGE_RE.match(s.strip())
        if not m:
            raise ValueError(f"invalid column range: {s!r}")
        a, b = m.group(1).upper(), m.group(2).upper()
        return cls(min_col=a, max_col=b)

    def __str__(self) -> str:
        return f"{self.min_col}:{self.max_col}"


@dataclass
class PrintTitles:
    """Container for repeat-rows + repeat-cols on a sheet.

    The OOXML representation is a workbook-level `<definedName
    name="_xlnm.Print_Titles" localSheetId="N">` node — the formula
    string concatenates the rows and cols ranges separated by a
    comma. This class is the typed Python view onto that string.
    """

    rows: RowRange | None = None
    cols: ColRange | None = None

    def is_empty(self) -> bool:
        return self.rows is None and self.cols is None

    def to_definedname_value(self, sheet_name: str) -> str | None:
        """Compose the `_xlnm.Print_Titles` formula string for ``sheet_name``."""
        if self.is_empty():
            return None
        # Excel needs the sheet name quoted if it contains spaces or
        # punctuation — we use the same conservative rule as openpyxl:
        # quote if any non-alphanumeric/non-underscore character is
        # present.
        if any(not (c.isalnum() or c == "_") for c in sheet_name):
            quoted = "'" + sheet_name.replace("'", "''") + "'"
        else:
            quoted = sheet_name
        parts: list[str] = []
        if self.rows is not None:
            parts.append(f"{quoted}!${self.rows.min_row}:${self.rows.max_row}")
        if self.cols is not None:
            parts.append(f"{quoted}!${self.cols.min_col}:${self.cols.max_col}")
        return ",".join(parts)


@dataclass
class PrintArea:
    """A print area definition (range string, e.g. ``"A1:D10"``)."""

    sqref: str | None = None

    def is_empty(self) -> bool:
        return not self.sqref


__all__ = [
    "RowRange",
    "ColRange",
    "PrintTitles",
    "PrintArea",
]
