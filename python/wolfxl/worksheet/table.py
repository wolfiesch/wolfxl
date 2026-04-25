"""openpyxl.worksheet.table compatibility.

T1 makes ``Table``, ``TableStyleInfo``, and ``TableColumn`` real
dataclasses. Read access (``ws.tables["SalesTable"].ref``) works for any
file opened in read/modify mode. Write access (``ws.add_table(t)``) works
in write mode (T1 PR5).

Field naming follows openpyxl's camelCase convention (``displayName``,
``headerRowCount``, ``tableStyleInfo``) even though that's un-Pythonic —
the whole point of this shim is drop-in compatibility with openpyxl.
"""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class TableColumn:
    """A single column within a table.

    openpyxl's ``TableColumn`` also carries calculated-column formulas,
    totals-row formulas, and style IDs. wolfxl preserves those on round-
    trip but does not expose them on the Python side yet; construction
    accepts just ``id`` + ``name``, which is what covers 99% of
    user-built tables.
    """

    id: int
    name: str


@dataclass
class TableStyleInfo:
    """Table style reference (``name``) plus banded-row/column flags.

    Excel ships named styles like ``"TableStyleLight9"``; this object
    records which style and which banding options are active.
    """

    name: str | None = None
    showFirstColumn: bool = False  # noqa: N815 - openpyxl public API
    showLastColumn: bool = False  # noqa: N815
    showRowStripes: bool = False  # noqa: N815
    showColumnStripes: bool = False  # noqa: N815


@dataclass
class Table:
    """An Excel table (ListObject) — a named, styled range.

    ``name`` is the internal identifier; ``displayName`` is what users
    see in the Name Box. openpyxl allows them to differ but they usually
    match. ``ref`` is the A1 range string (e.g. ``"A1:D10"``).

    ``headerRowCount`` is 1 when the first row is a header, 0 otherwise.
    ``totalsRowCount`` is the number of totals rows at the bottom.

    When constructed from a Rust-side dict, boolean fields like
    ``header_row=True`` map to ``headerRowCount=1``.
    """

    name: str
    displayName: str = ""  # noqa: N815 - openpyxl public API
    ref: str = ""
    headerRowCount: int = 1  # noqa: N815
    totalsRowCount: int = 0  # noqa: N815
    tableStyleInfo: TableStyleInfo | None = None  # noqa: N815
    tableColumns: list[TableColumn] = field(default_factory=list)  # noqa: N815

    def __post_init__(self) -> None:
        # Default displayName to name if not supplied (openpyxl behavior).
        if not self.displayName:
            self.displayName = self.name


__all__ = ["Table", "TableColumn", "TableStyleInfo"]
