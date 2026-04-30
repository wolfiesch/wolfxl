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

from collections.abc import Iterator
from dataclasses import dataclass, field
from typing import Any


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
    comment: str | None = None
    tableType: str | None = None  # noqa: N815
    headerRowCount: int = 1  # noqa: N815
    totalsRowCount: int = 0  # noqa: N815
    totalsRowShown: bool | None = None  # noqa: N815
    tableStyleInfo: TableStyleInfo | None = None  # noqa: N815
    tableColumns: list[TableColumn] = field(default_factory=list)  # noqa: N815

    def __post_init__(self) -> None:
        # Default displayName to name if not supplied (openpyxl behavior).
        if not self.displayName:
            self.displayName = self.name


# ---------------------------------------------------------------------------
# Pod 2 (RFC-060 §2.1) — extended re-exports.
# Sprint Π Pod-β (RFC-063) replaced the four construction stubs below
# (TableList / TablePartList / Related / XMLColumnProps) with real
# value types.  No save() pipeline changes — these are openpyxl-shaped
# wrappers over state that already plumbs through the patcher / native
# writer via the existing ``Table`` class and ``ws.tables`` dict.
# ---------------------------------------------------------------------------

from wolfxl.worksheet.filters import AutoFilter, SortState  # noqa: E402


@dataclass
class Related:
    """rels-pointer dataclass mirroring ``r:id="rId1"``.

    Used by openpyxl to point a ``<tablePart>`` element at the table's
    relationship entry.  Wolfxl tracks the rId allocation internally,
    but the dataclass is exposed here so user code that hand-builds a
    :class:`TablePartList` continues to work.
    """

    id: str = ""


@dataclass
class XMLColumnProps:
    """XML-column metadata for table-bound columns (CT_XmlColumnPr).

    Wolfxl preserves these properties on round-trip via the patcher.
    Construction is exposed here so user code that explicitly attaches
    a ``XMLColumnProps`` to a :class:`TableColumn` (openpyxl's
    ``column.xmlColumnPr``) ports mechanically.
    """

    mapId: int = 0  # noqa: N815 - openpyxl public API
    xpath: str = ""
    denormalized: bool = False
    xmlDataType: str = "string"  # noqa: N815


class TableList:
    """View over ``ws.tables`` (an existing ``dict[str, Table]``).

    Provides the openpyxl-shape ``__iter__`` / ``__len__`` /
    ``__contains__`` / :meth:`add` / :meth:`remove` API.  When bound
    to a worksheet, mutations are mirrored back onto ``ws.tables`` and
    (for additions) onto the worksheet's pending-tables queue so the
    save pipeline picks them up automatically.
    """

    __slots__ = ("worksheet", "_extra")

    def __init__(self, worksheet: Any = None) -> None:
        self.worksheet = worksheet
        self._extra: dict[str, Any] = {}

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _ws_tables(self) -> dict[str, Any] | None:
        ws = self.worksheet
        if ws is None:
            return None
        # ``ws.tables`` lazily materialises the cache; calling the
        # property is fine because the underlying dict is the source of
        # truth, and the cache is populated by the property itself.
        try:
            return ws.tables
        except Exception:
            return None

    # ------------------------------------------------------------------
    # openpyxl-shape API
    # ------------------------------------------------------------------

    def add(self, table: Any) -> None:
        """Register *table* on the underlying worksheet (or local view)."""
        ws = self.worksheet
        if ws is not None and hasattr(ws, "add_table"):
            ws.add_table(table)
            return
        name = getattr(table, "name", None) or getattr(table, "displayName", None)
        if not name:
            raise ValueError("TableList.add: table must expose a non-empty `name`")
        self._extra[name] = table

    def remove(self, table_name: str) -> None:
        """Remove a table by *name*.  Silently no-ops if the name is unknown."""
        backing = self._ws_tables()
        if backing is not None:
            backing.pop(table_name, None)
        else:
            self._extra.pop(table_name, None)

    def __iter__(self) -> Iterator[Any]:
        backing = self._ws_tables()
        if backing is not None:
            return iter(backing.values())
        return iter(self._extra.values())

    def __len__(self) -> int:
        backing = self._ws_tables()
        if backing is not None:
            return len(backing)
        return len(self._extra)

    def __contains__(self, name: Any) -> bool:
        backing = self._ws_tables()
        if backing is not None:
            return name in backing
        return name in self._extra

    def __getitem__(self, name: str) -> Any:
        backing = self._ws_tables()
        if backing is not None:
            return backing[name]
        return self._extra[name]

    def items(self) -> list[tuple[str, Any]]:
        backing = self._ws_tables()
        if backing is not None:
            return list(backing.items())
        return list(self._extra.items())

    def __repr__(self) -> str:  # pragma: no cover - trivial
        return f"TableList(count={len(self)})"


@dataclass
class TablePartList:
    """`<tableParts>` serialization helper (CT_TableParts §18.3.1.91).

    A simple holder for the count + list of :class:`Related` pointers
    Excel writes into ``<tableParts>`` underneath each ``<worksheet>``.
    Wolfxl regenerates this block from ``ws.tables`` at save time, so
    this dataclass is informational — but exposed for openpyxl source
    compatibility.
    """

    count: int = 0
    tablePart: list[Related] = field(default_factory=list)  # noqa: N815

    def __post_init__(self) -> None:
        if self.tablePart is None:  # pragma: no cover - defensive
            self.tablePart = []
        # Keep ``count`` and the list length in sync when the user only
        # supplied one of them.
        if self.count == 0 and self.tablePart:
            self.count = len(self.tablePart)

    def append(self, part: Related) -> None:
        """Add a :class:`Related` entry and bump :attr:`count`."""
        if not isinstance(part, Related):
            raise TypeError(
                f"TablePartList.append expects Related, got {type(part).__name__}"
            )
        self.tablePart.append(part)
        self.count = len(self.tablePart)

    def __iter__(self) -> Iterator[Related]:
        return iter(self.tablePart)

    def __len__(self) -> int:
        return len(self.tablePart)


__all__ = [
    "AutoFilter",
    "Related",
    "SortState",
    "Table",
    "TableColumn",
    "TableList",
    "TablePartList",
    "TableStyleInfo",
    "XMLColumnProps",
]
