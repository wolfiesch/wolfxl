"""Worksheet proxy — provides ``ws['A1']`` access and tracks dirty cells."""

from __future__ import annotations

import datetime as _dt
from collections.abc import Iterable, Iterator
from typing import TYPE_CHECKING, Any

from wolfxl._cell import Cell
from wolfxl._utils import a1_to_rowcol, column_index, rowcol_to_a1
from wolfxl.utils.cell import column_index_from_string, range_boundaries


def _coerce_col_idx(idx: int | str, op: str) -> int:
    """Accept either a 1-based int or an Excel column letter for col ops.

    Used by RFC-031 ``insert_cols`` / ``delete_cols``.
    """
    if isinstance(idx, str):
        try:
            i = column_index_from_string(idx)
        except Exception as exc:
            raise ValueError(
                f"{op}: idx {idx!r} is not a valid column letter"
            ) from exc
    elif isinstance(idx, int) and not isinstance(idx, bool):
        i = idx
    else:
        raise ValueError(f"{op}: idx must be int or str, got {idx!r}")
    if i < 1:
        raise ValueError(f"{op}: idx must be >= 1, got {idx!r}")
    return i


def _canonical_data_type(value: Any) -> str:
    """Map a Python value to the same canonical label the Rust reader emits.

    Rust's `read_sheet_records` returns `data_type` strings from a closed set
    (`string` / `number` / `boolean` / `datetime` / `error` / `formula` /
    `blank`). Overlay/Python-side records must use the same vocabulary so
    consumers that filter by these tokens see one schema across pure-read
    mode, modify mode, and pure-write mode.

    A string value beginning with ``=`` is classified as ``"formula"`` to
    match openpyxl's convention (and Rust's formula_map_cache path) — without
    this, pending formula edits in modify mode would silently downgrade to
    plain strings and any consumer counting/filtering formula records would
    miss them.
    """
    if value is None:
        return "blank"
    # bool is a subclass of int — check it first or "number" wins.
    if isinstance(value, bool):
        return "boolean"
    if isinstance(value, (int, float)):
        return "number"
    if isinstance(value, str):
        return "formula" if value.startswith("=") else "string"
    # All temporal types collapse to "datetime" to match the Rust reader,
    # whose `data_type_name()` emits a single "datetime" label for both
    # `Data::DateTime` and `Data::DateTimeIso`. Returning "date" for a
    # `datetime.date` would produce mixed schemas inside one
    # `cell_records()` result whenever an overlay edit touched a date
    # cell — consumers filtering on the documented tokens would silently
    # miss those records.
    if isinstance(value, (_dt.datetime, _dt.date, _dt.time)):
        return "datetime"
    return "string"

def _cellrichtext_to_runs_payload(crt: Any) -> list[tuple[str, dict[str, Any] | None]]:
    """Sprint Ι Pod-α: convert a ``CellRichText`` into the Rust-side
    payload (a list of ``(text, font_dict_or_None)`` tuples).

    This lives at module scope so both the write-mode and modify-mode
    flush paths can share it.  The Rust side reconstructs runs via
    ``py_runs_to_rust`` in ``src/wolfxl/mod.rs``.
    """
    out: list[tuple[str, dict[str, Any] | None]] = []
    for item in crt:
        if isinstance(item, str):
            out.append((item, None))
            continue
        # TextBlock — pull font props.
        font = item.font
        d: dict[str, Any] = {}
        if font.b is not None:
            d["b"] = bool(font.b)
        if font.i is not None:
            d["i"] = bool(font.i)
        if font.strike is not None:
            d["strike"] = bool(font.strike)
        if font.u is not None:
            d["u"] = font.u
        if font.sz is not None:
            d["sz"] = float(font.sz)
        if font.color is not None:
            d["color"] = font.color
        if font.rFont is not None:
            d["rFont"] = font.rFont
        if font.family is not None:
            d["family"] = int(font.family)
        if font.charset is not None:
            d["charset"] = int(font.charset)
        if font.vertAlign is not None:
            d["vertAlign"] = font.vertAlign
        if font.scheme is not None:
            d["scheme"] = font.scheme
        out.append((item.text, d if d else None))
    return out


if TYPE_CHECKING:
    from wolfxl._workbook import Workbook


class _RowDimensionProxy:
    """Dict-like proxy: ``ws.row_dimensions[1].height = 30``."""

    __slots__ = ("_ws",)

    def __init__(self, ws: Worksheet) -> None:
        self._ws = ws

    def __getitem__(self, row: int) -> _RowDimension:
        return _RowDimension(self._ws, row)

    def get(self, row: int, default: Any = None) -> _RowDimension | Any:
        if not isinstance(row, int):
            return default
        dimension = _RowDimension(self._ws, row)
        if dimension.height is not None or dimension.hidden or dimension.outline_level:
            return dimension
        return default


class _RowDimension:
    """Single row dimension with a readable/writable ``height`` property."""

    __slots__ = ("_ws", "_row")

    def __init__(self, ws: Worksheet, row: int) -> None:
        self._ws = ws
        self._row = row

    @property
    def height(self) -> float | None:
        wb = self._ws._workbook  # noqa: SLF001
        if wb._rust_reader is not None:  # noqa: SLF001
            return wb._rust_reader.read_row_height(self._ws._title, self._row)  # noqa: SLF001
        return self._ws._row_heights.get(self._row)  # noqa: SLF001

    @height.setter
    def height(self, value: float | None) -> None:
        self._ws._row_heights[self._row] = value  # noqa: SLF001

    @property
    def hidden(self) -> bool:
        wb = self._ws._workbook  # noqa: SLF001
        if wb._rust_reader is not None:  # noqa: SLF001
            return self._row in self._ws.sheet_visibility()["hidden_rows"]
        return False

    @property
    def outlineLevel(self) -> int:  # noqa: N802 - openpyxl public API
        return self.outline_level

    @property
    def outline_level(self) -> int:
        wb = self._ws._workbook  # noqa: SLF001
        if wb._rust_reader is not None:  # noqa: SLF001
            return int(self._ws.sheet_visibility()["row_outline_levels"].get(self._row, 0))
        return 0


class _ColumnDimensionProxy:
    """Dict-like proxy: ``ws.column_dimensions['A'].width = 15``."""

    __slots__ = ("_ws",)

    def __init__(self, ws: Worksheet) -> None:
        self._ws = ws

    def __getitem__(self, col_letter: str) -> _ColumnDimension:
        return _ColumnDimension(self._ws, col_letter.upper())

    def get(self, col_letter: str, default: Any = None) -> _ColumnDimension | Any:
        if not isinstance(col_letter, str):
            return default
        dimension = _ColumnDimension(self._ws, col_letter.upper())
        if dimension.width is not None or dimension.hidden or dimension.outline_level:
            return dimension
        return default


class _ColumnDimension:
    """Single column dimension with a readable/writable ``width`` property."""

    __slots__ = ("_ws", "_col_letter")

    def __init__(self, ws: Worksheet, col_letter: str) -> None:
        self._ws = ws
        self._col_letter = col_letter

    @property
    def width(self) -> float | None:
        wb = self._ws._workbook  # noqa: SLF001
        if wb._rust_reader is not None:  # noqa: SLF001
            return wb._rust_reader.read_column_width(self._ws._title, self._col_letter)  # noqa: SLF001
        return self._ws._col_widths.get(self._col_letter)  # noqa: SLF001

    @width.setter
    def width(self, value: float | None) -> None:
        self._ws._col_widths[self._col_letter] = value  # noqa: SLF001

    @property
    def hidden(self) -> bool:
        wb = self._ws._workbook  # noqa: SLF001
        if wb._rust_reader is not None:  # noqa: SLF001
            return column_index(self._col_letter) in self._ws.sheet_visibility()["hidden_columns"]
        return False

    @property
    def outlineLevel(self) -> int:  # noqa: N802 - openpyxl public API
        return self.outline_level

    @property
    def outline_level(self) -> int:
        wb = self._ws._workbook  # noqa: SLF001
        if wb._rust_reader is not None:  # noqa: SLF001
            col = column_index(self._col_letter)
            return int(self._ws.sheet_visibility()["column_outline_levels"].get(col, 0))
        return 0


class _AutoFilter:
    """Proxy for ``ws.auto_filter`` (RFC-056 Sprint Ο Pod 1B).

    Exposes the openpyxl-shaped surface:

    * ``ref`` — A1 range string (read/write).
    * ``filter_columns`` — list of :class:`FilterColumn`.
    * ``sort_state`` — optional :class:`SortState`.
    * ``add_filter_column(col_id, filter, **kw)`` — fluent builder.
    * ``add_sort_condition(ref, descending=…, sort_by=…, **kw)``.
    * ``to_rust_dict()`` — RFC-056 §10 dict shape for the patcher /
      native-writer PyO3 boundary.
    """

    __slots__ = ("_ref", "filter_columns", "sort_state")

    def __init__(self) -> None:
        from wolfxl.worksheet.filters import FilterColumn  # noqa: F401 — side-effectful import

        self._ref: str | None = None
        self.filter_columns: list[Any] = []
        self.sort_state: Any = None

    @property
    def ref(self) -> str | None:
        return self._ref

    @ref.setter
    def ref(self, value: str | None) -> None:
        self._ref = value

    def add_filter_column(
        self,
        col_id: int,
        filter: Any = None,
        *,
        hidden_button: bool = False,
        show_button: bool = True,
        date_group_items: Any = None,
    ) -> Any:
        """Append a ``FilterColumn``. Returns the new entry."""
        from wolfxl.worksheet.filters import FilterColumn

        fc = FilterColumn(
            col_id=col_id,
            hidden_button=hidden_button,
            show_button=show_button,
            filter=filter,
            date_group_items=list(date_group_items) if date_group_items else [],
        )
        self.filter_columns.append(fc)
        return fc

    def add_sort_condition(
        self,
        ref: str,
        descending: bool = False,
        sort_by: str = "value",
        *,
        custom_list: str | None = None,
        dxf_id: int | None = None,
        icon_set: str | None = None,
        icon_id: int | None = None,
    ) -> Any:
        """Append a ``SortCondition``. Auto-creates ``sort_state`` if absent."""
        from wolfxl.worksheet.filters import SortCondition, SortState

        if self.sort_state is None:
            self.sort_state = SortState(ref=ref)
        sc = SortCondition(
            ref=ref,
            descending=descending,
            sort_by=sort_by,
            custom_list=custom_list,
            dxf_id=dxf_id,
            icon_set=icon_set,
            icon_id=icon_id,
        )
        self.sort_state.sort_conditions.append(sc)
        return sc

    def to_rust_dict(self) -> dict[str, Any]:
        """Serialise to the RFC-056 §10 dict shape."""
        # Use the AutoFilter dataclass's own marshaller so the §10
        # contract has a single body of code.
        from wolfxl.worksheet.filters import AutoFilter as _AF

        af = _AF(
            ref=self._ref,
            filter_columns=list(self.filter_columns),
            sort_state=self.sort_state,
        )
        return af.to_rust_dict()


class _MergedCellsProxy:
    """openpyxl-shape proxy for ``Worksheet.merged_cells``.

    openpyxl's ``MultiCellRange`` exposes ``.ranges`` as an iterable of
    ``CellRange`` objects. SynthGL only iterates ``.ranges`` and stringifies
    each entry, so we expose a list of range strings — adequate for parity
    on the read path. Write-mode mutations still go through
    ``Worksheet.merge_cells`` / ``unmerge_cells``.
    """

    __slots__ = ("_ws",)

    def __init__(self, ws: Worksheet) -> None:
        self._ws = ws

    @property
    def ranges(self) -> list[str]:
        ws = self._ws
        # Write mode: trust the in-memory set (kept in sync by
        # ``merge_cells`` / ``unmerge_cells``).
        wb = ws._workbook  # noqa: SLF001
        if wb._rust_reader is None:  # noqa: SLF001
            return list(ws._merged_ranges)  # noqa: SLF001
        # Read mode: pull from the Rust calamine backend (already cached
        # there after first call). Falls back to the in-memory set if the
        # reader rejects the call (e.g. sheet was added in modify mode).
        try:
            return wb._rust_reader.read_merged_ranges(ws._title)  # noqa: SLF001
        except Exception:
            return list(ws._merged_ranges)  # noqa: SLF001

    def __iter__(self):  # type: ignore[no-untyped-def]
        return iter(self.ranges)

    def __len__(self) -> int:
        return len(self.ranges)


class Worksheet:
    """Proxy for a single worksheet in a Workbook."""

    __slots__ = (
        "_workbook", "_title", "_cells", "_dirty", "_dimensions",
        "_max_col_idx", "_next_append_row",
        "_append_buffer", "_append_buffer_start", "_bulk_writes",
        "_freeze_panes", "_auto_filter",
        "_row_heights", "_col_widths",
        "_merged_ranges", "_print_area", "_sheet_visibility_cache",
        # T1 read caches — populated lazily on first access.
        "_comments_cache", "_hyperlinks_cache",
        "_tables_cache", "_data_validations_cache",
        "_conditional_formatting_cache",
        # T1 write-mode pending queues — flushed in _flush() on save().
        "_pending_comments", "_pending_hyperlinks",
        "_pending_tables", "_pending_data_validations",
        "_pending_conditional_formats",
        # Sprint Ι Pod-α — pending rich-text values keyed by
        # (row, col).  Both write-mode (NativeWorkbook flush) and
        # modify-mode (XlsxPatcher flush) consume this map.
        "_pending_rich_text",
        # Sprint Ο Pod 1C (RFC-057) — pending array / data-table
        # formulas keyed by ``(row, col)``.  Each entry is a
        # ``(kind, payload)`` tuple where ``kind`` is one of
        # ``"array"``, ``"data_table"``, ``"spill_child"``.  Drained
        # at save time alongside the regular cell flush.
        "_pending_array_formulas",
        # Sprint Λ Pod-β (RFC-045) — pending image queue.
        "_pending_images",
        # Sprint Μ Pod-β (RFC-046) — pending chart queue.
        "_pending_charts",
        # Sprint Ν Pod-γ (RFC-048) — pending pivot table queue.
        "_pending_pivot_tables",
        # Sprint Ο Pod 1A (RFC-055) — print/view/protection lazy slots.
        "_page_setup", "_page_margins", "_header_footer",
        "_sheet_view", "_protection",
        "_print_title_rows", "_print_title_cols",
        # RFC-061 Sub-feature 3.1 — pending slicer presentations.
        "_pending_slicers",
        # Sprint Π Pod Π-α (RFC-062) — page breaks + sheetFormatPr.
        "_row_breaks", "_col_breaks", "_sheet_format",
    )

    def __init__(self, workbook: Workbook, title: str) -> None:
        self._workbook = workbook
        self._title = title
        self._cells: dict[tuple[int, int], Cell] = {}
        self._dirty: set[tuple[int, int]] = set()
        self._dimensions: tuple[int, int] | None = None
        self._max_col_idx: int = 0
        self._next_append_row: int = 1
        # Fast-path append buffer: raw value lists, no Cell objects.
        self._append_buffer: list[list[Any]] = []
        self._append_buffer_start: int = 1
        # Bulk write buffer: list of (grid, start_row, start_col) tuples.
        self._bulk_writes: list[tuple[list[list[Any]], int, int]] = []
        # openpyxl compat properties
        self._freeze_panes: str | None = None
        self._auto_filter = _AutoFilter()
        self._row_heights: dict[int, float | None] = {}
        self._col_widths: dict[str, float | None] = {}
        self._merged_ranges: set[str] = set()
        self._print_area: str | None = None
        self._sheet_visibility_cache: dict[str, Any] | None = None
        # T1 read caches (None = not loaded yet; dict/list = loaded, possibly empty).
        self._comments_cache: dict[str, Any] | None = None
        self._hyperlinks_cache: dict[str, Any] | None = None
        self._tables_cache: dict[str, Any] | None = None
        self._data_validations_cache: Any | None = None
        self._conditional_formatting_cache: Any | None = None
        # T1 write-mode pending queues (flushed in _flush() on save()).
        self._pending_comments: dict[str, Any] = {}
        self._pending_hyperlinks: dict[str, Any] = {}
        # Sprint Ι Pod-α — keyed by (row, col) so the flush layer can
        # look it up by sparse coordinate without coordinate-string round-trip.
        self._pending_rich_text: dict[tuple[int, int], Any] = {}
        # Sprint Ο Pod 1C (RFC-057) — pending array / data-table
        # formulas keyed by ``(row, col)``.  Each entry is a
        # ``(kind, payload)`` tuple.  Master cells get
        # ``("array", {"ref": ..., "text": ...})`` or
        # ``("data_table", {...})``; cells inside the spill range
        # become ``("spill_child", {})`` placeholders.
        self._pending_array_formulas: dict[tuple[int, int], tuple[str, dict[str, Any]]] = {}
        self._pending_tables: list[Any] = []
        self._pending_data_validations: list[Any] = []
        self._pending_conditional_formats: list[tuple[str, Any]] = []
        # Sprint Λ Pod-β (RFC-045) — pending images attached to this
        # sheet via ``add_image``. Drained at save time into the Rust
        # writer (write mode) or the patcher (modify mode).
        self._pending_images: list[Any] = []
        # Sprint Μ Pod-β (RFC-046) — pending charts queued via ``add_chart``.
        # Drained at save time into ``_rust_writer.add_chart_native`` (write
        # mode) or the patcher (modify mode, queued by Pod-γ's plumbing).
        self._pending_charts: list[Any] = []
        # Sprint Ν Pod-γ (RFC-048) — pending pivot table queue. Drained
        # at save time into the patcher via
        # ``_workbook._flush_pending_pivots_to_patcher`` (modify mode
        # only — write-mode pivot tables are not yet supported and
        # should fail loud at ``add_pivot_table`` call site).
        self._pending_pivot_tables: list[Any] = []
        # Sprint Ο Pod 1A (RFC-055) — print/view/protection. All lazy:
        # we instantiate the openpyxl-shaped wrappers only on first
        # attribute access so a workbook that never touches these
        # surfaces pays zero overhead.
        self._page_setup: Any = None
        self._page_margins: Any = None
        self._header_footer: Any = None
        self._sheet_view: Any = None
        self._protection: Any = None
        self._print_title_rows: str | None = None
        self._print_title_cols: str | None = None
        # RFC-061 Sub-feature 3.1 — pending slicer presentations.
        self._pending_slicers: list[Any] = []
        # Sprint Π Pod Π-α (RFC-062) — page breaks + sheet format
        # defaults. All lazy: instantiate the wrappers only on first
        # attribute access so a workbook that never touches these
        # surfaces pays zero overhead.
        self._row_breaks: Any = None
        self._col_breaks: Any = None
        self._sheet_format: Any = None

    @property
    def title(self) -> str:
        return self._title

    @title.setter
    def title(self, value: str) -> None:
        """Rename this worksheet (write mode only)."""
        wb = self._workbook
        old = self._title
        if old == value:
            return
        if value in wb._sheets:  # noqa: SLF001
            raise ValueError(f"Sheet '{value}' already exists")
        # Update workbook bookkeeping.
        idx = wb._sheet_names.index(old)  # noqa: SLF001
        wb._sheet_names[idx] = value  # noqa: SLF001
        wb._sheets[value] = wb._sheets.pop(old)  # noqa: SLF001
        self._title = value
        # Sync the Rust writer so ensure_sheet_exists() sees the new name.
        if wb._rust_writer is not None:  # noqa: SLF001
            wb._rust_writer.rename_sheet(old, value)  # noqa: SLF001

    # ------------------------------------------------------------------
    # openpyxl compat properties
    # ------------------------------------------------------------------

    @property
    def freeze_panes(self) -> str | None:
        """Get/set the freeze panes cell reference (e.g. ``'B2'``).

        In read mode, reads from the Rust backend.  In write mode,
        the value is stored and flushed to Rust on ``save()``.
        """
        wb = self._workbook
        if wb._rust_reader is not None and self._freeze_panes is None:  # noqa: SLF001
            info = wb._rust_reader.read_freeze_panes(self._title)  # noqa: SLF001
            if info and info.get("mode"):
                return info.get("top_left_cell")
            return None
        return self._freeze_panes

    @freeze_panes.setter
    def freeze_panes(self, value: str | None) -> None:
        self._freeze_panes = value
        # Mirror the mutation onto sheet_view.pane so callers reading
        # ``ws.sheet_view.pane`` after a ``ws.freeze_panes = "B2"``
        # observe a consistent snapshot. RFC-055 §2.5.
        if self._sheet_view is not None:
            from wolfxl.worksheet.views import Pane
            if value is None:
                self._sheet_view.pane = None
            else:
                from wolfxl._utils import a1_to_rowcol
                try:
                    r, c = a1_to_rowcol(value)
                except Exception:
                    return
                self._sheet_view.pane = Pane(
                    xSplit=float(c - 1),
                    ySplit=float(r - 1),
                    topLeftCell=value,
                    activePane="bottomRight",
                    state="frozen",
                )

    # ------------------------------------------------------------------
    # Sprint Ο Pod 1A (RFC-055) — print / view / protection accessors
    # ------------------------------------------------------------------

    @property
    def page_setup(self) -> Any:
        """Lazy ``PageSetup`` accessor (RFC-055 §2.1)."""
        if self._page_setup is None:
            from wolfxl.worksheet.page_setup import PageSetup
            self._page_setup = PageSetup()
        return self._page_setup

    @page_setup.setter
    def page_setup(self, value: Any) -> None:
        self._page_setup = value

    @property
    def page_margins(self) -> Any:
        """Lazy ``PageMargins`` accessor (RFC-055 §2.2)."""
        if self._page_margins is None:
            from wolfxl.worksheet.page_setup import PageMargins
            self._page_margins = PageMargins()
        return self._page_margins

    @page_margins.setter
    def page_margins(self, value: Any) -> None:
        self._page_margins = value

    @property
    def HeaderFooter(self) -> Any:  # noqa: N802 - openpyxl alias
        return self.header_footer

    @property
    def header_footer(self) -> Any:
        """Lazy ``HeaderFooter`` accessor (RFC-055 §2.3)."""
        if self._header_footer is None:
            from wolfxl.worksheet.header_footer import HeaderFooter
            self._header_footer = HeaderFooter()
        return self._header_footer

    @header_footer.setter
    def header_footer(self, value: Any) -> None:
        self._header_footer = value

    @property
    def sheet_view(self) -> Any:
        """Lazy ``SheetView`` accessor (RFC-055 §2.5).

        ``ws.freeze_panes`` mutations are mirrored into ``sheet_view.pane``
        on the setter side; on the getter side, if a sheet view has a
        non-None pane we surface that as ``freeze_panes`` for parity.
        """
        if self._sheet_view is None:
            from wolfxl.worksheet.views import Pane, SheetView
            sv = SheetView()
            # If freeze_panes was set before the lazy view materialized,
            # carry the state across so the view is consistent.
            if self._freeze_panes is not None:
                from wolfxl._utils import a1_to_rowcol
                try:
                    r, c = a1_to_rowcol(self._freeze_panes)
                    sv.pane = Pane(
                        xSplit=float(c - 1),
                        ySplit=float(r - 1),
                        topLeftCell=self._freeze_panes,
                        activePane="bottomRight",
                        state="frozen",
                    )
                except Exception:
                    pass
            self._sheet_view = sv
        return self._sheet_view

    @sheet_view.setter
    def sheet_view(self, value: Any) -> None:
        self._sheet_view = value

    @property
    def protection(self) -> Any:
        """Lazy ``SheetProtection`` accessor (RFC-055 §2.6)."""
        if self._protection is None:
            from wolfxl.worksheet.protection import SheetProtection
            self._protection = SheetProtection()
        return self._protection

    @protection.setter
    def protection(self, value: Any) -> None:
        self._protection = value

    # ------------------------------------------------------------------
    # Sprint Π Pod Π-α (RFC-062) — page breaks + sheet format props
    # ------------------------------------------------------------------

    @property
    def row_breaks(self) -> Any:
        """Lazy ``PageBreakList`` of horizontal page breaks (RFC-062 §3)."""
        if self._row_breaks is None:
            from wolfxl.worksheet.pagebreak import PageBreakList
            self._row_breaks = PageBreakList()
        return self._row_breaks

    @row_breaks.setter
    def row_breaks(self, value: Any) -> None:
        self._row_breaks = value

    @property
    def col_breaks(self) -> Any:
        """Lazy ``PageBreakList`` of vertical page breaks (RFC-062 §3)."""
        if self._col_breaks is None:
            from wolfxl.worksheet.pagebreak import PageBreakList
            self._col_breaks = PageBreakList()
        return self._col_breaks

    @col_breaks.setter
    def col_breaks(self, value: Any) -> None:
        self._col_breaks = value

    @property
    def page_breaks(self) -> Any:
        """openpyxl alias — ``ws.page_breaks`` is a row-breaks alias."""
        return self.row_breaks

    @page_breaks.setter
    def page_breaks(self, value: Any) -> None:
        self._row_breaks = value

    @property
    def sheet_format(self) -> Any:
        """Lazy ``SheetFormatProperties`` accessor (RFC-062 §3)."""
        if self._sheet_format is None:
            from wolfxl.worksheet.dimensions import SheetFormatProperties
            self._sheet_format = SheetFormatProperties()
        return self._sheet_format

    @sheet_format.setter
    def sheet_format(self, value: Any) -> None:
        self._sheet_format = value

    @property
    def dimension_holder(self) -> Any:
        """Return a fresh ``DimensionHolder`` view bound to this worksheet."""
        from wolfxl.worksheet.dimensions import DimensionHolder
        return DimensionHolder(self)

    def to_rust_page_breaks_dict(self) -> dict[str, Any]:
        """Return the §10 dict shape for ``<rowBreaks>`` / ``<colBreaks>``.

        Each side is ``None`` when the corresponding ``PageBreakList``
        is un-touched OR carries zero breaks — the patcher / writer
        then knows to skip emitting the corresponding XML block.
        """
        d: dict[str, Any] = {}
        d["row_breaks"] = (
            self._row_breaks.to_rust_dict()
            if self._row_breaks is not None and len(self._row_breaks) > 0
            else None
        )
        d["col_breaks"] = (
            self._col_breaks.to_rust_dict()
            if self._col_breaks is not None and len(self._col_breaks) > 0
            else None
        )
        return d

    def to_rust_sheet_format_dict(self) -> dict[str, Any] | None:
        """Return the §10 dict for ``<sheetFormatPr>`` or ``None``.

        Returns ``None`` when the wrapper is un-touched OR at all-default
        values — the writer then keeps the legacy hardcoded
        ``<sheetFormatPr defaultRowHeight="15"/>`` emit path.
        """
        if self._sheet_format is None or self._sheet_format.is_default():
            return None
        return self._sheet_format.to_rust_dict()

    @property
    def print_title_rows(self) -> str | None:
        """Repeat-rows for printing (RFC-055 §2.4)."""
        return self._print_title_rows

    @print_title_rows.setter
    def print_title_rows(self, value: str | None) -> None:
        if value is not None:
            from wolfxl.worksheet.print_settings import RowRange
            # Validate and normalize.
            self._print_title_rows = str(RowRange.from_string(value))
        else:
            self._print_title_rows = None

    @property
    def print_title_cols(self) -> str | None:
        """Repeat-cols for printing (RFC-055 §2.4)."""
        return self._print_title_cols

    @print_title_cols.setter
    def print_title_cols(self, value: str | None) -> None:
        if value is not None:
            from wolfxl.worksheet.print_settings import ColRange
            self._print_title_cols = str(ColRange.from_string(value))
        else:
            self._print_title_cols = None

    def to_rust_setup_dict(self) -> dict[str, Any]:
        """Return the §10 dict contract for the Rust patcher / writer.

        Returns ``None`` for any sub-block whose Python wrapper is at
        its construction defaults — the Rust side then knows to skip
        emitting the corresponding XML.
        """
        d: dict[str, Any] = {}
        d["page_setup"] = (
            self._page_setup.to_rust_dict()
            if self._page_setup is not None and not self._page_setup.is_default()
            else None
        )
        d["page_margins"] = (
            self._page_margins.to_rust_dict()
            if self._page_margins is not None and not self._page_margins.is_default()
            else None
        )
        d["header_footer"] = (
            self._header_footer.to_rust_dict()
            if self._header_footer is not None and not self._header_footer.is_default()
            else None
        )
        d["sheet_view"] = (
            self._sheet_view.to_rust_dict()
            if self._sheet_view is not None and not self._sheet_view.is_default()
            else None
        )
        d["sheet_protection"] = (
            self._protection.to_rust_dict()
            if self._protection is not None and not self._protection.is_default()
            else None
        )
        if self._print_title_rows is not None or self._print_title_cols is not None:
            d["print_titles"] = {
                "rows": self._print_title_rows,
                "cols": self._print_title_cols,
            }
        else:
            d["print_titles"] = None
        return d

    @property
    def auto_filter(self) -> _AutoFilter:
        return self._auto_filter

    @property
    def row_dimensions(self) -> _RowDimensionProxy:
        return _RowDimensionProxy(self)

    @property
    def column_dimensions(self) -> _ColumnDimensionProxy:
        return _ColumnDimensionProxy(self)

    @property
    def print_area(self) -> str | None:
        """Get/set the print area range string (e.g. ``'A1:D10'``).

        Stored locally and flushed to the Rust writer on ``save()`` if the
        writer supports ``set_print_area()``.
        """
        return self._print_area

    @print_area.setter
    def print_area(self, value: str | None) -> None:
        self._print_area = value

    # ------------------------------------------------------------------
    # Cell access
    # ------------------------------------------------------------------

    def __getitem__(self, key: Any) -> Any:
        """openpyxl-compatible cell access.

        Supports:
        - ``ws["A1"]`` -> single Cell
        - ``ws["A1:B2"]`` -> tuple of tuples of Cell (2D range)
        - ``ws["A:B"]`` -> column range bounded by used range
        - ``ws["1:3"]`` -> row range
        - ``ws["A"]`` -> single column (tuple of Cell)
        - ``ws["1"]`` -> single row (str key; tuple of Cell)
        - ``ws[1]`` -> single row (int key; tuple of Cell)
        - ``ws[1:3]`` -> row slice (tuple of tuples of Cell)
        """
        # Integer row access: ws[1] -> row 1 cells
        if isinstance(key, int):
            return self._get_row_tuple(key, key)[0]

        # Integer slice: ws[1:3] -> rows 1..3 INCLUSIVE (openpyxl contract).
        if isinstance(key, slice):
            if key.start is None or key.stop is None:
                raise ValueError("Row slice bounds must be specified")
            return self._get_row_tuple(key.start, key.stop)

        if isinstance(key, str):
            return self._resolve_string_key(key)

        raise TypeError(f"Worksheet indices must be str, int, or slice, not {type(key).__name__}")

    def _resolve_string_key(self, key: str) -> Any:
        """Resolve a string key to Cell / tuple / tuple-of-tuples."""
        # Single A1 coord like "A1" — keep the fast path.
        try:
            row, col = a1_to_rowcol(key)
        except ValueError:
            pass
        else:
            return self._get_or_create_cell(row, col)

        # Pure digits "3" -> single row.
        if key.isdigit():
            n = int(key)
            return self._get_row_tuple(n, n)[0]

        # Pure letters "A" -> single column.
        try:
            col_idx = column_index_from_string(key)
        except ValueError:
            col_idx = None
        if col_idx is not None and not any(ch.isdigit() for ch in key):
            return tuple(row for row in self._get_col_tuple(col_idx, col_idx))[0]

        # Otherwise: range form ("A1:B2", "A:B", "1:3").
        min_col, min_row, max_col, max_row = range_boundaries(key)

        if min_row is None and max_row is None:
            # Whole-column range like "A:B" -> openpyxl returns column-major
            # (one tuple of cells per column). Bounded by the sheet's used rows.
            bounded_max_row = self._max_row()
            if min_col is None or max_col is None:
                raise ValueError(f"Invalid range: {key!r}")
            return self._get_col_tuple(min_col, max_col, 1, bounded_max_row)

        if min_col is None and max_col is None:
            # Whole-row range like "1:3" -> row-major
            bounded_max_col = self._max_col()
            if min_row is None or max_row is None:
                raise ValueError(f"Invalid range: {key!r}")
            return self._get_rect(min_row, 1, max_row, bounded_max_col)

        if min_row is None or max_row is None or min_col is None or max_col is None:
            raise ValueError(f"Invalid range: {key!r}")

        # Degenerate single-cell range like "A1:A1" -> still return single Cell
        # per openpyxl's contract for non-range strings — but a colon in the
        # key means the user asked for a range, so return a 2D tuple.
        return self._get_rect(min_row, min_col, max_row, max_col)

    def _get_rect(
        self, min_row: int, min_col: int, max_row: int, max_col: int,
    ) -> tuple[tuple[Cell, ...], ...]:
        """Return a 2D tuple of Cells for the inclusive rectangle."""
        return tuple(
            tuple(
                self._get_or_create_cell(r, c) for c in range(min_col, max_col + 1)
            )
            for r in range(min_row, max_row + 1)
        )

    def _get_row_tuple(
        self, min_row: int, max_row: int,
    ) -> tuple[tuple[Cell, ...], ...]:
        """Return a tuple of row-tuples for rows min_row..max_row inclusive."""
        max_c = self._max_col()
        return tuple(
            tuple(self._get_or_create_cell(r, c) for c in range(1, max_c + 1))
            for r in range(min_row, max_row + 1)
        )

    def _get_col_tuple(
        self,
        min_col: int,
        max_col: int,
        min_row: int | None = None,
        max_row: int | None = None,
    ) -> tuple[tuple[Cell, ...], ...]:
        """Return a tuple of column-tuples for cols min_col..max_col inclusive."""
        r_min = min_row if min_row is not None else 1
        r_max = max_row if max_row is not None else self._max_row()
        return tuple(
            tuple(self._get_or_create_cell(r, c) for r in range(r_min, r_max + 1))
            for c in range(min_col, max_col + 1)
        )

    def __setitem__(self, key: str, value: Any) -> None:
        """``ws['A1'] = 42`` — shorthand for setting a cell's value."""
        cell = self[key]
        cell.value = value

    def cell(self, row: int, column: int, value: Any = None) -> Cell:
        """Get or create a cell by 1-based (row, column). Matches openpyxl API."""
        c = self._get_or_create_cell(row, column)
        if value is not None:
            c.value = value
        return c

    def _get_or_create_cell(self, row: int, col: int) -> Cell:
        # Materialize the append buffer on first random cell access so that
        # Cell objects exist for previously-appended rows.
        if self._append_buffer:
            self._materialize_append_buffer()
        key = (row, col)
        if key not in self._cells:
            self._cells[key] = Cell(self, row, col)
        return self._cells[key]

    def _mark_dirty(self, row: int, col: int) -> None:
        self._dirty.add((row, col))

    # ------------------------------------------------------------------
    # Append (openpyxl-compatible row insertion)
    # ------------------------------------------------------------------

    def append(self, iterable: Iterable[Any]) -> None:
        """Append a row of values. Matches openpyxl's ``ws.append()`` API.

        Successive calls auto-increment the row index. Values are written to
        columns starting at 1 (A).

        Performance: rows are buffered as raw Python lists — no Cell objects
        are created. The buffer is flushed directly to ``write_sheet_values()``
        on save, bypassing per-cell FFI overhead entirely.  If cell-level
        access is needed later (e.g. ``ws.cell(1,1).font = ...``), the buffer
        is materialized into Cell objects on demand.
        """
        row = list(iterable)
        if not self._append_buffer:
            self._append_buffer_start = self._next_append_row
        self._append_buffer.append(row)
        ncols = len(row)
        if ncols > self._max_col_idx:
            self._max_col_idx = ncols
        self._next_append_row += 1

    def _materialize_append_buffer(self) -> None:
        """Convert the append buffer into Cell objects.

        Called lazily on the first ``cell()`` / ``__getitem__`` access after
        appending.  After this call ``_append_buffer`` is empty and all values
        live in the normal ``_cells`` / ``_dirty`` tracking.
        """
        start = self._append_buffer_start
        buf = self._append_buffer
        if not buf:
            return
        # Temporarily clear the buffer FIRST to avoid re-entrant
        # materialization when self.cell() calls _get_or_create_cell().
        self._append_buffer = []
        for i, row_vals in enumerate(buf):
            r = start + i
            for c, val in enumerate(row_vals, start=1):
                self.cell(row=r, column=c, value=val)
        # Buffer is already cleared above.

    def write_rows(
        self,
        rows: list[list[Any]],
        start_row: int = 1,
        start_col: int = 1,
    ) -> None:
        """Bulk-write a 2D grid of values starting at (start_row, start_col).

        Unlike ``append()``, this writes to an arbitrary position. Values are
        buffered and flushed via a single ``write_sheet_values()`` Rust call
        at save time, avoiding per-cell FFI overhead.

        ``rows`` is a list of lists. Each inner list is one row of values.
        """
        if not rows:
            return
        # Store a shallow copy so flush can safely mutate without affecting caller.
        copied = [list(row) for row in rows]
        self._bulk_writes.append((copied, start_row, start_col))

    def _materialize_bulk_writes(self) -> None:
        """Convert bulk write buffers into Cell objects.

        Called before the patcher flush path which has no batch API and
        needs all values as individual dirty cells.
        """
        writes = self._bulk_writes
        if not writes:
            return
        self._bulk_writes = []
        for grid, sr, sc in writes:
            for ri, row in enumerate(grid):
                for ci, val in enumerate(row):
                    if val is not None:
                        self.cell(row=sr + ri, column=sc + ci, value=val)

    @staticmethod
    def _extract_non_batchable(
        grid: list[list[Any]], start_row: int, start_col: int,
    ) -> list[tuple[int, int, Any]]:
        """Extract non-batchable values from grid, replacing them with None.

        Non-batchable: booleans, formulas (str starting with '='), and
        non-primitive types (dates, datetimes, etc.).  These require
        per-cell ``write_cell_value()`` calls with type-preserving payloads.
        """
        indiv: list[tuple[int, int, Any]] = []
        for ri, row in enumerate(grid):
            for ci, val in enumerate(row):
                if val is not None and (
                    isinstance(val, bool)
                    or (isinstance(val, str) and val.startswith("="))
                    or not isinstance(val, (int, float, str))
                ):
                    indiv.append((start_row + ri, start_col + ci, val))
                    row[ci] = None
        return indiv

    # ------------------------------------------------------------------
    # Iteration
    # ------------------------------------------------------------------

    def iter_rows(
        self,
        min_row: int | None = None,
        max_row: int | None = None,
        min_col: int | None = None,
        max_col: int | None = None,
        values_only: bool = False,
    ) -> Iterator[tuple[Any, ...]]:
        """Iterate over rows in a range. Matches openpyxl's iter_rows API.

        Sprint Ι Pod-β: when the workbook was opened with
        ``read_only=True`` (or when this sheet has more than
        :data:`wolfxl._streaming.AUTO_STREAM_ROW_THRESHOLD` rows) this
        method becomes a true SAX-streaming generator backed by the
        Rust ``StreamingSheetReader``. Cells yielded in that path are
        :class:`wolfxl._streaming.StreamingCell` instances —
        immutable read-only proxies whose ``font`` / ``fill`` /
        ``border`` / ``alignment`` / ``number_format`` properties
        delegate to the existing eager style table. ``values_only=True``
        yields plain value tuples and never instantiates a cell object.
        """
        wb = self._workbook  # noqa: SLF001
        # Sprint Ι Pod-β — streaming fast path (read_only=True OR auto-trigger).
        if wb._rust_reader is not None and getattr(wb, "_source_path", None):  # noqa: SLF001
            from wolfxl._streaming import should_auto_stream, stream_iter_rows

            stream_now = bool(getattr(wb, "_read_only", False)) or should_auto_stream(self)
            if stream_now:
                yield from stream_iter_rows(
                    self, min_row, max_row, min_col, max_col, values_only=values_only
                )
                return

        # Fast bulk path: read-mode + values_only -> single Rust FFI call.
        if values_only and self._workbook._rust_reader is not None:  # noqa: SLF001
            yield from self._iter_rows_bulk(min_row, max_row, min_col, max_col)
            return

        r_min = min_row or 1
        r_max = max_row or self._max_row()
        c_min = min_col or 1
        c_max = max_col or self._max_col()

        for r in range(r_min, r_max + 1):
            if values_only:
                yield tuple(
                    self._get_or_create_cell(r, c).value for c in range(c_min, c_max + 1)
                )
            else:
                yield tuple(
                    self._get_or_create_cell(r, c) for c in range(c_min, c_max + 1)
                )

    def iter_cols(
        self,
        min_col: int | None = None,
        max_col: int | None = None,
        min_row: int | None = None,
        max_row: int | None = None,
        values_only: bool = False,
    ) -> Iterator[tuple[Any, ...]]:
        """Iterate over columns in a range. Matches openpyxl's iter_cols API.

        Unlike ``iter_rows``, the first-position arguments are column bounds.
        Yields one tuple per column, each containing values (or Cells) for
        every row in range.
        """
        # Fast bulk path: read-mode + values_only -> single Rust FFI call.
        if values_only and self._workbook._rust_reader is not None:  # noqa: SLF001
            yield from self._iter_cols_bulk(min_col, max_col, min_row, max_row)
            return

        r_min = min_row or 1
        r_max = max_row or self._max_row()
        c_min = min_col or 1
        c_max = max_col or self._max_col()

        for c in range(c_min, c_max + 1):
            if values_only:
                yield tuple(
                    self._get_or_create_cell(r, c).value for r in range(r_min, r_max + 1)
                )
            else:
                yield tuple(
                    self._get_or_create_cell(r, c) for r in range(r_min, r_max + 1)
                )

    def _iter_cols_bulk(
        self,
        min_col: int | None,
        max_col: int | None,
        min_row: int | None,
        max_row: int | None,
    ) -> Iterator[tuple[Any, ...]]:
        """Bulk-read column values via a single Rust FFI call, then transpose.

        Mirrors ``_iter_rows_bulk`` but yields column-major tuples. One
        ``read_sheet_values_plain`` call reads the whole rectangle; transposition
        happens in Python. This avoids per-cell Rust calls in the values_only
        fast path and keeps parity with ``iter_rows`` performance characteristics.
        """
        from wolfxl._cell import _payload_to_python

        reader = self._workbook._rust_reader  # noqa: SLF001
        sheet = self._title
        data_only = getattr(self._workbook, "_data_only", False)

        r_min = min_row or 1
        r_max = max_row or self._max_row()
        c_min = min_col or 1
        c_max = max_col or self._max_col()
        range_str = f"{rowcol_to_a1(r_min, c_min)}:{rowcol_to_a1(r_max, c_max)}"

        use_plain = hasattr(reader, "read_sheet_values_plain")
        if use_plain:
            rows = reader.read_sheet_values_plain(sheet, range_str, data_only)
        else:
            rows = reader.read_sheet_values(sheet, range_str, data_only)

        if not rows:
            return

        expected_cols = c_max - c_min + 1
        expected_rows = r_max - r_min + 1

        # Normalize every row to expected_cols width so transposition is safe.
        normalized: list[list[Any]] = []
        for row in rows:
            if use_plain:
                vals = list(row)
            else:
                vals = [_payload_to_python(cell) for cell in row]
            n = len(vals)
            if n >= expected_cols:
                normalized.append(vals[:expected_cols])
            else:
                normalized.append(vals + [None] * (expected_cols - n))

        # Pad rows if Rust returned fewer rows than requested.
        while len(normalized) < expected_rows:
            normalized.append([None] * expected_cols)

        for c_offset in range(expected_cols):
            yield tuple(normalized[r_offset][c_offset] for r_offset in range(expected_rows))

    @property
    def rows(self) -> Iterator[tuple[Any, ...]]:
        """Iterator over rows (tuples of Cell) — openpyxl alias for ``iter_rows()``."""
        return self.iter_rows()

    @property
    def columns(self) -> Iterator[tuple[Any, ...]]:
        """Iterator over columns (tuples of Cell) — openpyxl alias for ``iter_cols()``."""
        return self.iter_cols()

    @property
    def values(self) -> Iterator[tuple[Any, ...]]:
        """Iterator over row values — openpyxl alias for ``iter_rows(values_only=True)``."""
        return self.iter_rows(values_only=True)

    def _iter_rows_bulk(
        self,
        min_row: int | None,
        max_row: int | None,
        min_col: int | None,
        max_col: int | None,
    ) -> Iterator[tuple[Any, ...]]:
        """Bulk-read values via a single Rust FFI call (values_only fast path).

        Uses ``read_sheet_values_plain()`` when available (returns native
        Python objects), falling back to ``read_sheet_values()`` + per-cell
        ``_payload_to_python()`` conversion otherwise.
        """
        from wolfxl._cell import _payload_to_python

        reader = self._workbook._rust_reader  # noqa: SLF001
        sheet = self._title
        data_only = getattr(self._workbook, "_data_only", False)

        # Build an A1:B2-style range string for Rust.
        r_min = min_row or 1
        r_max = max_row or self._max_row()
        c_min = min_col or 1
        c_max = max_col or self._max_col()
        range_str = f"{rowcol_to_a1(r_min, c_min)}:{rowcol_to_a1(r_max, c_max)}"

        # Prefer plain-value read (no dict overhead) if available.
        use_plain = hasattr(reader, "read_sheet_values_plain")
        if use_plain:
            rows = reader.read_sheet_values_plain(sheet, range_str, data_only)
        else:
            rows = reader.read_sheet_values(sheet, range_str, data_only)

        if not rows:
            return

        # The Rust range returns exactly the rows/cols we asked for,
        # so no Python-side slicing is needed.
        expected_cols = c_max - c_min + 1
        for row in rows:
            if use_plain:
                # Already native Python values; pad/trim to expected width.
                n = len(row)
                if n >= expected_cols:
                    yield tuple(row[:expected_cols])
                else:
                    yield tuple(row) + (None,) * (expected_cols - n)
            else:
                # Dict payloads need conversion.
                vals = [_payload_to_python(cell) for cell in row]
                n = len(vals)
                if n >= expected_cols:
                    yield tuple(vals[:expected_cols])
                else:
                    yield tuple(vals) + (None,) * (expected_cols - n)

    def iter_cell_records(
        self,
        min_row: int | None = None,
        max_row: int | None = None,
        min_col: int | None = None,
        max_col: int | None = None,
        *,
        data_only: bool | None = None,
        include_format: bool = True,
        include_empty: bool = False,
        include_formula_blanks: bool = True,
        include_coordinate: bool = True,
        include_style_id: bool = True,
        include_extended_format: bool = True,
        include_cached_formula_value: bool = False,
    ) -> Iterator[dict[str, Any]]:
        """Iterate populated cells as compact dictionaries.

        This is WolfXL's high-throughput read API for ingestion and dataframe
        workloads. In read mode it makes one Rust call for the requested range,
        returning native Python values plus optional formatting metadata such as
        ``number_format``, ``bold``, ``indent``, and border cues.

        Coordinates are openpyxl-style 1-based ``row`` / ``column`` integers.
        Empty cells are skipped by default; pass ``include_empty=True`` when a
        dense rectangular record stream is needed. Formula cells without a
        backing cached value are included by default; pass
        ``include_formula_blanks=False`` to skip those template-only formulas.
        Pass ``include_coordinate=False`` when row/column integers are enough
        and avoiding A1 string allocation matters. Pass
        ``include_style_id=False`` when semantic format fields are enough and
        callers do not need workbook-internal style identifiers. Pass
        ``include_extended_format=False`` to keep raw font flags and number
        formats while skipping expensive style-grid fields such as fill,
        alignment, and border cues. Pass
        ``include_cached_formula_value=True`` to include a ``cached_value`` key
        on formula records that have a saved cached result.
        """
        if self._workbook._rust_reader is None:  # noqa: SLF001
            yield from self._iter_cell_records_python(
                min_row=min_row,
                max_row=max_row,
                min_col=min_col,
                max_col=max_col,
                include_empty=include_empty,
                include_coordinate=include_coordinate,
            )
            return

        reader = self._workbook._rust_reader  # noqa: SLF001
        effective_data_only = self._workbook._data_only if data_only is None else data_only  # noqa: SLF001
        overlay = self._collect_pending_overlay()
        unbounded_sparse_read = (
            min_row is None
            and max_row is None
            and min_col is None
            and max_col is None
            and not include_empty
            and not overlay
        )
        if unbounded_sparse_read:
            r_min = c_min = 1
            r_max = c_max = None
            range_str = None
        else:
            r_min = min_row or 1
            r_max = max_row or self._max_row()
            c_min = min_col or 1
            c_max = max_col or self._max_col()
            range_str = f"{rowcol_to_a1(r_min, c_min)}:{rowcol_to_a1(r_max, c_max)}"
        records = reader.read_sheet_records(
            self._title,
            range_str,
            effective_data_only,
            include_format,
            include_empty,
            include_formula_blanks,
            include_coordinate,
            include_style_id,
            include_extended_format,
            include_cached_formula_value,
        )

        # Modify mode can have pending Python-side edits the Rust reader
        # can't see. Overlay them on top of the on-disk records so the
        # iterator reflects current worksheet state, not the last save.
        # The overlay is only built when something is dirty — pure read
        # mode pays no extra cost.
        if not overlay:
            yield from records
            return

        seen: set[tuple[int, int]] = set()
        for record in records:
            row, col = int(record["row"]), int(record["column"])
            key = (row, col)
            if key not in overlay:
                yield record
                continue
            seen.add(key)
            new_value = overlay[key]
            if new_value is None and not include_empty:
                continue
            patched = dict(record)
            patched["value"] = new_value
            patched["data_type"] = _canonical_data_type(new_value)
            # The on-disk record may carry a "formula" field from the
            # original cell. After an overlay edit, that field is stale:
            # a literal-overwrites-formula edit must drop it, and a
            # formula-overwrites-literal edit must replace it. Strip the
            # leading "=" to match the Rust reader's convention (formula
            # text is stored without the prefix; openpyxl writes it back).
            if isinstance(new_value, str) and new_value.startswith("="):
                patched["formula"] = new_value[1:]
            else:
                patched.pop("formula", None)
            patched.pop("cached_value", None)
            yield patched

        # Yield pending edits that were inside the requested range but the
        # Rust reader didn't return (e.g. empty-on-disk cell user just set).
        for (row, col), value in overlay.items():
            if (row, col) in seen:
                continue
            if not (r_min <= row <= r_max and c_min <= col <= c_max):
                continue
            if value is None and not include_empty:
                continue
            extra: dict[str, Any] = {
                "row": row,
                "column": col,
                "value": value,
                "data_type": _canonical_data_type(value),
            }
            # Mirror the patched-overlay branch: a formula string emits the
            # `formula` key (Rust-style, leading "=" stripped) so consumers
            # that pull `record["formula"]` for formula cells see the
            # expression for unsaved edits, not just on-disk records.
            if isinstance(value, str) and value.startswith("="):
                extra["formula"] = value[1:]
            if include_coordinate:
                extra["coordinate"] = rowcol_to_a1(row, col)
            yield extra

    def _collect_pending_overlay(self) -> dict[tuple[int, int], Any]:
        """Return ``{(row, col): value}`` for cells modified since the last save.

        Includes explicit cell edits (anything in ``_dirty``), the append
        buffer, and bulk-write grids. Returns an empty dict when nothing is
        pending — the Rust read path stays a hot, allocation-free loop.
        """
        overlay: dict[tuple[int, int], Any] = {}
        if self._dirty:
            for key in self._dirty:
                cell = self._cells.get(key)
                if cell is not None:
                    overlay[key] = cell.value
        if self._append_buffer:
            start = self._append_buffer_start
            for row_offset, row_values in enumerate(self._append_buffer):
                for col_offset, value in enumerate(row_values):
                    overlay[(start + row_offset, col_offset + 1)] = value
        for grid, start_row, start_col in self._bulk_writes:
            for row_offset, row_values in enumerate(grid):
                for col_offset, value in enumerate(row_values):
                    overlay[(start_row + row_offset, start_col + col_offset)] = value
        return overlay

    def cell_records(
        self,
        min_row: int | None = None,
        max_row: int | None = None,
        min_col: int | None = None,
        max_col: int | None = None,
        *,
        data_only: bool | None = None,
        include_format: bool = True,
        include_empty: bool = False,
        include_formula_blanks: bool = True,
        include_coordinate: bool = True,
        include_style_id: bool = True,
        include_extended_format: bool = True,
        include_cached_formula_value: bool = False,
    ) -> list[dict[str, Any]]:
        """Return ``iter_cell_records(...)`` as a list."""
        return list(
            self.iter_cell_records(
                min_row=min_row,
                max_row=max_row,
                min_col=min_col,
                max_col=max_col,
                data_only=data_only,
                include_format=include_format,
                include_empty=include_empty,
                include_formula_blanks=include_formula_blanks,
                include_coordinate=include_coordinate,
                include_style_id=include_style_id,
                include_extended_format=include_extended_format,
                include_cached_formula_value=include_cached_formula_value,
            ),
        )

    def cached_formula_values(self, *, qualified: bool = False) -> dict[str, Any]:
        """Return cached formula results for this sheet.

        Keys are A1 coordinates by default. Pass ``qualified=True`` to return
        ``"Sheet!A1"`` keys, matching :meth:`Workbook.cached_formula_values`.
        Only formula cells with saved cached values are included; uncached
        template formulas are omitted.
        """
        wb = self._workbook
        if wb._rust_reader is None:  # noqa: SLF001
            return {}
        values = dict(wb._rust_reader.read_cached_formula_values(self._title))  # noqa: SLF001
        if not qualified:
            return values
        return {f"{self._title}!{cell_ref}": value for cell_ref, value in values.items()}

    def sheet_visibility(self) -> dict[str, Any]:
        """Return hidden rows/columns and outline levels for this sheet.

        Row and column identifiers are 1-based to mirror openpyxl's dimension
        collections. The returned shape is:
        ``hidden_rows``, ``hidden_columns``, ``row_outline_levels``, and
        ``column_outline_levels``.
        """
        if self._sheet_visibility_cache is not None:
            return self._sheet_visibility_cache

        wb = self._workbook
        if wb._rust_reader is None:  # noqa: SLF001
            self._sheet_visibility_cache = {
                "hidden_rows": [],
                "hidden_columns": [],
                "row_outline_levels": {},
                "column_outline_levels": {},
            }
            return self._sheet_visibility_cache
        self._sheet_visibility_cache = dict(wb._rust_reader.read_sheet_visibility(self._title))  # noqa: SLF001
        return self._sheet_visibility_cache

    def _iter_cell_records_python(
        self,
        *,
        min_row: int | None,
        max_row: int | None,
        min_col: int | None,
        max_col: int | None,
        include_empty: bool,
        include_coordinate: bool = True,
    ) -> Iterator[dict[str, Any]]:
        r_min = min_row or 1
        r_max = max_row or self._max_row()
        c_min = min_col or 1
        c_max = max_col or self._max_col()
        for row in range(r_min, r_max + 1):
            for col in range(c_min, c_max + 1):
                cell = self._get_or_create_cell(row, col)
                value = cell.value
                if value is None and not include_empty:
                    continue
                record: dict[str, Any] = {
                    "row": row,
                    "column": col,
                    "value": value,
                    "data_type": _canonical_data_type(value),
                }
                if include_coordinate:
                    record["coordinate"] = rowcol_to_a1(row, col)
                yield record

    def calculate_dimension(self) -> str:
        """Return the used worksheet range in openpyxl's ``A1:C10`` form."""
        bounds = self._read_dimension_bounds()
        if bounds is None:
            return "A1:A1"
        min_row, min_col, max_row, max_col = bounds
        return f"{rowcol_to_a1(min_row, min_col)}:{rowcol_to_a1(max_row, max_col)}"

    def _read_dimension_bounds(self) -> tuple[int, int, int, int] | None:
        """Return 1-based ``(min_row, min_col, max_row, max_col)`` bounds.

        Modify mode has both a Rust reader (for the on-disk extents) and
        Python-side pending writes (cells/append buffer/bulk writes). The
        reported bounds must be the union, otherwise callers that derive
        ranges from ``calculate_dimension()`` miss unsaved edits.
        """
        wb = self._workbook
        rust_bounds: tuple[int, int, int, int] | None = None
        if wb._rust_reader is not None:  # noqa: SLF001
            raw = wb._rust_reader.read_sheet_bounds(self._title)  # noqa: SLF001
            if isinstance(raw, tuple) and len(raw) == 4:
                rust_bounds = tuple(int(value) for value in raw)  # type: ignore[assignment]

        pending = self._pending_writes_bounds()
        if rust_bounds is None and pending is None:
            return None
        if pending is None:
            return rust_bounds
        if rust_bounds is None:
            return pending
        return (
            min(rust_bounds[0], pending[0]),
            min(rust_bounds[1], pending[1]),
            max(rust_bounds[2], pending[2]),
            max(rust_bounds[3], pending[3]),
        )

    def _pending_writes_bounds(self) -> tuple[int, int, int, int] | None:
        """Bounds of unsaved Python-side writes: ``(min_row, min_col, max_row, max_col)``.

        Used by ``_read_dimension_bounds`` to fold modify-mode edits into the
        on-disk Rust bounds, and by ``_max_row``/``_max_col`` so write-mode
        ``ws.max_row`` reflects ``append()``/``write_rows()`` that haven't
        materialized yet.

        Iterates ``_dirty`` (set of actually-modified cell keys) rather than
        ``_cells`` — the cell map is populated by mere read access
        (``ws['Z999']`` materializes a Cell without modifying it), so reading
        a far cell would otherwise inflate dimension bounds and trigger
        oversized scans in downstream callers.
        """
        dirty = self._dirty
        buf = self._append_buffer
        bulk = self._bulk_writes
        if not dirty and not buf and not bulk:
            return None
        min_r = min_c = None
        max_r = max_c = 0
        for row, col in dirty:
            if min_r is None or row < min_r:
                min_r = row
            if min_c is None or col < min_c:
                min_c = col
            if row > max_r:
                max_r = row
            if col > max_c:
                max_c = col
        if buf:
            start = self._append_buffer_start
            buf_max_r = start + len(buf) - 1
            # An empty appended row (`ws.append([])`) still consumes a row
            # index but contributes no columns. Without `default=0`, a buf
            # of all-empty rows would be a max() over an empty generator;
            # with it but no >0 guard, the column-bounds branch would set
            # min_c=1 / max_c=0, which `_max_col()` would then emit as the
            # invalid 1-based column 0 to `rowcol_to_a1`.
            buf_max_c = max((len(row) for row in buf), default=0)
            if min_r is None or start < min_r:
                min_r = start
            if buf_max_r > max_r:
                max_r = buf_max_r
            if buf_max_c > 0:
                if min_c is None or 1 < min_c:
                    min_c = 1
                if buf_max_c > max_c:
                    max_c = buf_max_c
        for grid, start_row, start_col in bulk:
            if not grid:
                continue
            grid_max_r = start_row + len(grid) - 1
            # Same zero-width guard: a grid where every row is empty would
            # yield grid_max_c = start_col - 1, potentially below 1.
            grid_width = max((len(row) for row in grid), default=0)
            if grid_width == 0:
                if min_r is None or start_row < min_r:
                    min_r = start_row
                if grid_max_r > max_r:
                    max_r = grid_max_r
                continue
            grid_max_c = start_col + grid_width - 1
            if min_r is None or start_row < min_r:
                min_r = start_row
            if min_c is None or start_col < min_c:
                min_c = start_col
            if grid_max_r > max_r:
                max_r = grid_max_r
            if grid_max_c > max_c:
                max_c = grid_max_c
        if min_r is None or min_c is None:
            return None
        return min_r, min_c, max_r, max_c

    def _read_dimensions(self) -> tuple[int, int]:
        """Discover sheet dimensions from the Rust backend (read mode only)."""
        if self._dimensions is not None:
            return self._dimensions
        wb = self._workbook
        if wb._rust_reader is None:  # noqa: SLF001
            self._dimensions = (1, 1)
            return self._dimensions
        xml_dims = wb._rust_reader.read_sheet_dimensions(self._title)  # noqa: SLF001
        if isinstance(xml_dims, tuple) and len(xml_dims) == 2:
            self._dimensions = (int(xml_dims[0]), int(xml_dims[1]))
            return self._dimensions
        rows = wb._rust_reader.read_sheet_values(self._title, None, False)  # noqa: SLF001
        if not rows or not isinstance(rows, list):
            self._dimensions = (1, 1)
            return self._dimensions
        max_r = len(rows)
        max_c = max((len(row) for row in rows), default=1)
        self._dimensions = (max_r, max_c)
        return self._dimensions

    def _max_row(self) -> int:
        # Read mode honors the on-disk ``<dimension>`` tag (parity with
        # openpyxl, which trusts the tag even when trailing rows are empty).
        # Modify/write modes additionally union with Python-side pending
        # writes so ``ws.max_row`` reflects ``append()`` / ``write_rows()`` /
        # cell edits before save.
        if self._workbook._rust_reader is not None:  # noqa: SLF001
            disk_max_r = self._read_dimensions()[0]
        else:
            disk_max_r = max((k[0] for k in self._cells), default=0)
        pending = self._pending_writes_bounds()
        if pending is None:
            return disk_max_r if disk_max_r else 1
        return max(disk_max_r, pending[2])

    def _max_col(self) -> int:
        if self._workbook._rust_reader is not None:  # noqa: SLF001
            disk_max_c = self._read_dimensions()[1]
        else:
            disk_max_c = max((k[1] for k in self._cells), default=0)
        pending = self._pending_writes_bounds()
        if pending is None:
            return disk_max_c if disk_max_c else 1
        return max(disk_max_c, pending[3])

    # openpyxl exposes these as properties, not methods. Mirror that contract
    # so ``ws.max_row`` (no parens) works as a drop-in for openpyxl callers.
    # Pinned by ``tests/parity/test_read_parity.py`` (uses ``op_ws.max_row``).
    @property
    def max_row(self) -> int:
        return self._max_row()

    @property
    def max_column(self) -> int:
        return self._max_col()

    @property
    def min_row(self) -> int:
        """Always 1, matching openpyxl's contract (no real "first used" tracking)."""
        return 1

    @property
    def min_column(self) -> int:
        """Always 1, matching openpyxl's contract."""
        return 1

    @property
    def dimensions(self) -> str:
        """Used worksheet range in A1 form, e.g. ``"A1:C10"``."""
        return self.calculate_dimension()

    @property
    def parent(self) -> Workbook:
        """The containing Workbook."""
        return self._workbook

    @property
    def sheet_state(self) -> str:
        """Visibility state: ``"visible"``, ``"hidden"``, or ``"veryHidden"``.

        Defaults to ``"visible"``. wolfxl doesn't yet wire through the
        ``<sheet state="hidden">`` XML attribute; returning the default
        matches openpyxl's value for a freshly-created sheet.
        """
        return "visible"

    @property
    def _charts(self) -> list[Any]:
        """Sprint Μ Pod-β (RFC-046) — list of charts queued via ``add_chart``.

        Mirrors ``openpyxl.worksheet.worksheet.Worksheet._charts``. Returns
        the live pending list so mutations propagate to the next save and
        ``len(ws._charts)`` works in user code that mirrors openpyxl's
        read-side behaviour.
        """
        return self._pending_charts

    def add_chart(self, chart: Any, anchor: Any = None) -> None:
        """Sprint Μ Pod-β (RFC-046) — attach a chart to this worksheet.

        Mirrors :meth:`openpyxl.worksheet.worksheet.Worksheet.add_chart`.

        Parameters
        ----------
        chart : wolfxl.chart.ChartBase subclass
            The chart to embed (BarChart / LineChart / PieChart / …).
        anchor : str | None
            Where to anchor the chart. Accepts an A1-style cell ref
            (``"D2"``) or ``None`` (defaults to ``"E15"`` to match
            openpyxl's :class:`ChartBase` default).

        The chart is queued until ``Workbook.save()``, at which point
        the writer (write mode) or patcher (modify mode) emits the
        chart/drawing/rels parts.
        """
        from wolfxl.chart._chart import ChartBase as _ChartBase

        if not isinstance(chart, _ChartBase):
            raise TypeError(
                f"add_chart expected wolfxl.chart.ChartBase, got "
                f"{type(chart).__name__}"
            )

        if anchor is None:
            anchor = chart.anchor if chart.anchor is not None else "E15"

        # RFC-046 §10.11.2 — anchor must be a valid A1 cell ref or a
        # recognized anchor object (RFC-045 OneCellAnchor / TwoCellAnchor /
        # AbsoluteAnchor). Strings are validated via the A1 regex; non-str
        # values are accepted opaquely (the writer validates further).
        if isinstance(anchor, str):
            self._validate_a1_anchor(anchor)

        chart._anchor = anchor  # noqa: SLF001
        self._pending_charts.append(chart)

    def add_pivot_table(self, pivot_table: Any) -> None:
        """Sprint Ν Pod-γ (RFC-048) — anchor a pivot table on this sheet.

        The pivot table's ``cache`` MUST already be registered on the
        owning workbook via :meth:`Workbook.add_pivot_cache` (so the
        cache has a populated ``_cache_id``). The table is queued and
        drained at ``Workbook.save()`` time via
        ``_workbook._flush_pending_pivots_to_patcher``.

        Args:
            pivot_table: A :class:`wolfxl.pivot.PivotTable` instance.

        Raises:
            TypeError: If ``pivot_table`` is not a
                :class:`PivotTable`.
            ValueError: If the pivot table's cache has not been
                registered yet (its ``_cache_id`` is ``None``).
            RuntimeError: If the workbook is not in modify mode.
        """
        from wolfxl.pivot import PivotTable as _PivotTable

        if not isinstance(pivot_table, _PivotTable):
            raise TypeError(
                f"add_pivot_table expected wolfxl.pivot.PivotTable, "
                f"got {type(pivot_table).__name__}"
            )
        if self._workbook._rust_patcher is None:  # noqa: SLF001
            raise RuntimeError(
                "add_pivot_table requires modify mode — open the "
                "workbook with load_workbook(..., modify=True). "
                "Write-mode pivot table emission is not yet supported."
            )
        if pivot_table.cache._cache_id is None:  # noqa: SLF001
            raise ValueError(
                "PivotTable.cache has not been registered with the "
                "workbook yet. Call Workbook.add_pivot_cache(cache) "
                "before Worksheet.add_pivot_table(pt)."
            )
        # Compute the layout up-front so any field-axis errors surface
        # synchronously here rather than at save() time.
        if hasattr(pivot_table, "_compute_layout"):
            pivot_table._compute_layout()
        self._pending_pivot_tables.append(pivot_table)

    def add_slicer(self, slicer: Any, anchor: str) -> None:
        """RFC-061 §2.1 — anchor a slicer presentation on this sheet.

        The slicer's ``cache`` MUST already be registered on the
        workbook via :meth:`Workbook.add_slicer_cache`. The slicer
        is queued and drained at ``Workbook.save()`` time via
        ``_workbook._flush_pending_slicers_to_patcher``.

        Args:
            slicer: A :class:`wolfxl.pivot.Slicer` instance.
            anchor: A1-style anchor cell (top-left of the slicer's
                graphic frame), e.g. ``"H2"``.

        Raises:
            TypeError: If ``slicer`` is not a Slicer.
            ValueError: If the slicer's cache has not been registered
                or ``anchor`` is not a valid A1 string.
            RuntimeError: If the workbook is not in modify mode.
        """
        from wolfxl.pivot import Slicer as _Slicer

        if not isinstance(slicer, _Slicer):
            raise TypeError(
                f"add_slicer expected wolfxl.pivot.Slicer, got "
                f"{type(slicer).__name__}"
            )
        if self._workbook._rust_patcher is None:  # noqa: SLF001
            raise RuntimeError(
                "add_slicer requires modify mode — open the workbook "
                "with load_workbook(..., modify=True)."
            )
        if slicer.cache._slicer_cache_id is None:  # noqa: SLF001
            raise ValueError(
                "Slicer.cache has not been registered with the "
                "workbook yet. Call Workbook.add_slicer_cache(cache) "
                "before Worksheet.add_slicer(slicer, anchor)."
            )
        if not isinstance(anchor, str) or not anchor:
            raise ValueError(
                "Worksheet.add_slicer: anchor must be a non-empty A1 string"
            )
        self._validate_a1_anchor(anchor)
        slicer.anchor = anchor
        self._pending_slicers.append(slicer)

    @staticmethod
    def _validate_a1_anchor(anchor: str) -> None:
        """Raise :class:`ValueError` if *anchor* is not a valid A1 cell ref.

        Per RFC-046 §10.11.2: ``r"^[A-Z]+[0-9]+$"`` — single cell
        coordinates only (e.g. ``"E15"``, ``"AA200"``). Range refs and
        sheet-qualified refs are rejected; pass an anchor object for
        more complex placements. Excel's column max is ``XFD`` (16384)
        and row max is 1048576; refs outside those bounds raise.
        """
        import re
        if not anchor:
            raise ValueError("anchor must not be empty")
        m = re.match(r"^([A-Z]+)([0-9]+)$", anchor)
        if not m:
            raise ValueError(
                f"anchor={anchor!r} must be a single A1 cell ref like 'E15' "
                f"(regex ^[A-Z]+[0-9]+$); for ranged or absolute placement "
                f"pass an OneCellAnchor / TwoCellAnchor / AbsoluteAnchor"
            )
        col_letters, row_str = m.group(1), m.group(2)
        # Column letters → 1-based index.
        col_idx = 0
        for ch in col_letters:
            col_idx = col_idx * 26 + (ord(ch) - ord("A") + 1)
        if col_idx > 16384:
            raise ValueError(
                f"anchor={anchor!r}: column {col_letters!r} exceeds Excel max XFD (16384)"
            )
        row_idx = int(row_str)
        if row_idx < 1 or row_idx > 1_048_576:
            raise ValueError(
                f"anchor={anchor!r}: row {row_idx} out of Excel range [1, 1048576]"
            )

    def remove_chart(self, chart: Any) -> None:
        """Sprint Ξ (RFC-050) — remove a previously-added chart.

        Mirrors the openpyxl idiom ``ws._charts.remove(chart)``.

        Parameters
        ----------
        chart : wolfxl.chart.ChartBase subclass
            A chart instance previously passed to :meth:`add_chart`
            on this worksheet. Identity is matched by Python ``is``,
            not equality, so the caller must pass the same object.

        Raises
        ------
        ValueError
            If *chart* was never added to this worksheet (or has
            already been removed).

        Notes
        -----
        In **write mode** this removes the chart from the
        ``_pending_charts`` list; the writer simply does not emit
        the chart part on save.

        In **modify mode** (where the chart was already persisted on
        disk) this method currently only handles the
        not-yet-flushed case (chart is still in
        ``_pending_charts``); removing a chart that survives from
        the source workbook is tracked as a v1.8 follow-up
        (``Worksheet.delete_chart_persisted`` — needs the patcher to
        emit a chart-removal queue alongside ``queue_chart_add``).
        """
        try:
            self._pending_charts.remove(chart)
        except ValueError:
            raise ValueError(
                "chart was not added to this worksheet via add_chart() "
                "(or has already been removed). Removal of charts that "
                "survive from the source workbook is a v1.8 follow-up; "
                "see RFC-050 §6."
            ) from None

    def replace_chart(self, old: Any, new: Any) -> None:
        """Sprint Ξ (RFC-050) — replace one chart with another in place.

        Convenience for ``ws.remove_chart(old); ws.add_chart(new, old._anchor)``
        that preserves the anchor and the position in the chart list
        (so deterministic ID allocation matches the pre-replace layout).

        Parameters
        ----------
        old : wolfxl.chart.ChartBase subclass
            The chart to replace. Must have been added via
            :meth:`add_chart`.
        new : wolfxl.chart.ChartBase subclass
            The replacement. Inherits *old*'s anchor unless
            ``new._anchor`` was already set explicitly.

        Raises
        ------
        ValueError
            If *old* was never added to this worksheet.
        TypeError
            If *new* is not a :class:`ChartBase` instance.
        """
        from wolfxl.chart._chart import ChartBase as _ChartBase
        if not isinstance(new, _ChartBase):
            raise TypeError(
                f"replace_chart expected wolfxl.chart.ChartBase for new, got "
                f"{type(new).__name__}"
            )
        try:
            idx = self._pending_charts.index(old)
        except ValueError:
            raise ValueError(
                "old chart was not added to this worksheet via add_chart()"
            ) from None
        anchor = new._anchor if new._anchor is not None else old._anchor  # noqa: SLF001
        if anchor is None:
            anchor = "E15"
        if isinstance(anchor, str):
            self._validate_a1_anchor(anchor)
        new._anchor = anchor  # noqa: SLF001
        self._pending_charts[idx] = new

    @property
    def _images(self) -> list[Any]:
        """Sprint Λ Pod-β (RFC-045) — list of images queued via ``add_image``.

        Used by openpyxl-compat code that iterates ``ws._images`` (e.g.
        SynthGL utilities that mirror openpyxl's read-side behaviour).
        Returns the live list so mutations propagate to the next save.
        """
        return self._pending_images

    def add_image(self, img: Any, anchor: Any = None) -> None:
        """Sprint Λ Pod-β (RFC-045) — attach an image to this worksheet.

        Mirrors :meth:`openpyxl.worksheet.worksheet.Worksheet.add_image`.

        Parameters
        ----------
        img : wolfxl.drawing.image.Image
            The image to embed. Constructed from a path/BytesIO/bytes.
        anchor : str | TwoCellAnchor | AbsoluteAnchor | None
            Where to anchor the image. Accepts:

            - ``"B5"`` (A1 cell ref) — one-cell anchor, image extends
              naturally from its top-left corner. This is what
              openpyxl users overwhelmingly write.
            - :class:`wolfxl.drawing.spreadsheet_drawing.TwoCellAnchor`
              — image stretches between two cells.
            - :class:`wolfxl.drawing.spreadsheet_drawing.AbsoluteAnchor`
              — pure EMU coordinates, no cell binding.
            - ``None`` — defaults to ``"A1"`` (matches openpyxl).

        The image is queued until ``Workbook.save()``, at which point
        the writer (write mode) or the patcher (modify mode) emits the
        drawing/media/rels parts.
        """
        from wolfxl.drawing.image import Image as _Image
        from wolfxl.drawing.spreadsheet_drawing import (
            AbsoluteAnchor,
            OneCellAnchor,
            TwoCellAnchor,
        )

        if not isinstance(img, _Image):
            raise TypeError(
                f"add_image expected wolfxl.drawing.image.Image, got {type(img).__name__}"
            )

        if anchor is None:
            anchor = "A1"

        # Stash the resolved anchor on the image (openpyxl semantics)
        # AND on a local copy so this exact call's anchor is captured
        # even if the user reuses the Image object.
        img.anchor = anchor
        self._pending_images.append(img)

    # ------------------------------------------------------------------
    # End add_image
    # ------------------------------------------------------------------

    @property
    def merged_cells(self) -> _MergedCellsProxy:
        """openpyxl-shape merged-cells accessor.

        Lazy: on first access, pulls merged ranges from the Rust reader
        (read mode) or returns the in-memory write-mode set. Always exposes
        a ``.ranges`` iterable of range strings — matching openpyxl's
        ``MultiCellRange`` shape closely enough for SynthGL's needs.
        """
        return _MergedCellsProxy(self)

    # ------------------------------------------------------------------
    # Write-mode helpers
    # ------------------------------------------------------------------

    def merge_cells(self, range_string: str) -> None:
        """Merge cells (write mode only). Example: ``ws.merge_cells('A1:B2')``."""
        wb = self._workbook
        if wb._rust_writer is None:  # noqa: SLF001
            raise RuntimeError("merge_cells requires write mode")
        wb._rust_writer.merge_cells(self._title, range_string)  # noqa: SLF001
        self._merged_ranges.add(range_string)

    def unmerge_cells(self, range_string: str) -> None:
        """Unmerge a previously merged range.

        If *range_string* was not previously merged, silently does nothing
        (matches openpyxl behaviour).
        """
        self._merged_ranges.discard(range_string)

    # ------------------------------------------------------------------
    # Structural ops — scheduled, not yet implemented
    # ------------------------------------------------------------------
    #
    # Each stub raises NotImplementedError with an RFC pointer so users
    # see a discoverable roadmap entry instead of an AttributeError. The
    # workaround note targets the most common escape hatch: do the
    # structural shuffle in openpyxl, then read the result with
    # ``wolfxl.load_workbook`` for the heavy work.

    def insert_rows(self, idx: int, amount: int = 1) -> None:
        """Shift rows down to insert *amount* empty rows starting at *idx*.

        Implements RFC-030. Validates ``idx >= 1`` and ``amount >= 1``;
        raises ``ValueError`` otherwise. Queues a row-shift op on the
        owning workbook's ``_pending_axis_shifts`` list. The op is
        drained at ``Workbook.save()`` time and applied by the patcher
        (Phase 2.5i in ``src/wolfxl/mod.rs``).

        See ``Plans/rfcs/030-insert-delete-rows.md`` for full semantics
        (formula shift, hyperlink/table/DV/CF anchor shift, defined-name
        shift, comment/VML drawing anchor shift).
        """
        if not isinstance(idx, int) or idx < 1:
            raise ValueError(
                f"insert_rows: idx must be a positive integer (>=1), got {idx!r}"
            )
        if not isinstance(amount, int) or amount < 1:
            raise ValueError(
                f"insert_rows: amount must be a positive integer (>=1), got {amount!r}"
            )
        self._workbook._pending_axis_shifts.append(  # noqa: SLF001
            (self.title, "row", idx, amount)
        )

    def delete_rows(self, idx: int, amount: int = 1) -> None:
        """Delete *amount* rows starting at *idx*, shifting subsequent rows up.

        Implements RFC-030. Validates ``idx >= 1`` and ``amount >= 1``;
        raises ``ValueError`` otherwise. Refs that point INTO the
        deleted band become ``#REF!`` per OOXML semantics.
        """
        if not isinstance(idx, int) or idx < 1:
            raise ValueError(
                f"delete_rows: idx must be a positive integer (>=1), got {idx!r}"
            )
        if not isinstance(amount, int) or amount < 1:
            raise ValueError(
                f"delete_rows: amount must be a positive integer (>=1), got {amount!r}"
            )
        self._workbook._pending_axis_shifts.append(  # noqa: SLF001
            (self.title, "row", idx, -amount)
        )

    def insert_cols(self, idx: int | str, amount: int = 1) -> None:
        """Shift columns right to insert *amount* empty columns at *idx*.

        Implements RFC-031. ``idx`` may be a 1-based int or an Excel
        column letter (``"A"``, ``"AB"``, ...). Validates ``idx >= 1``
        and ``amount >= 0``; ``amount == 0`` is a noop. Queues a
        col-shift op on the owning workbook's ``_pending_axis_shifts``.

        See ``Plans/rfcs/031-insert-delete-cols.md`` for full semantics
        (formula shift, anchor shift, ``<col>`` span split).
        """
        idx_i = _coerce_col_idx(idx, "insert_cols")
        if not isinstance(amount, int) or amount < 0:
            raise ValueError(
                f"insert_cols: amount must be an integer >= 0, got {amount!r}"
            )
        if amount == 0:
            return
        self._workbook._pending_axis_shifts.append(  # noqa: SLF001
            (self.title, "col", idx_i, amount)
        )

    def delete_cols(self, idx: int | str, amount: int = 1) -> None:
        """Delete *amount* columns starting at *idx*, shifting subsequent columns left.

        Implements RFC-031. ``idx`` may be a 1-based int or an Excel
        column letter. Refs that point INTO the deleted band become
        ``#REF!`` per OOXML semantics. ``amount == 0`` is a noop.
        """
        idx_i = _coerce_col_idx(idx, "delete_cols")
        if not isinstance(amount, int) or amount < 0:
            raise ValueError(
                f"delete_cols: amount must be an integer >= 0, got {amount!r}"
            )
        if amount == 0:
            return
        self._workbook._pending_axis_shifts.append(  # noqa: SLF001
            (self.title, "col", idx_i, -amount)
        )

    def move_range(
        self,
        cell_range: Any,
        rows: int = 0,
        cols: int = 0,
        translate: bool = False,
    ) -> None:
        """Move a rectangular block of cells by *rows* / *cols*.

        Implements RFC-034. Paste-style relocation: every cell whose
        A1 coordinate falls inside *cell_range* is physically moved
        to ``(row + rows, col + cols)``. Existing cells at the
        destination are silently overwritten (matches openpyxl).

        Formulas inside the moved block are paste-translated:
        relative refs shift by ``(rows, cols)``; absolute refs (``$A$1``)
        DO NOT shift. With ``translate=True``, formulas in cells
        outside the moved block that reference cells inside the source
        rectangle are also re-anchored to the new location.

        ``cell_range`` accepts an A1 range string (``"C3:E10"``) or a
        single-cell string (``"A1"``). Validates ``rows`` and
        ``cols`` are ints (signed). Raises ``ValueError`` if the
        destination would land outside Excel's coordinate space
        (rows ``1..1_048_576``, cols ``1..16_384``).

        ``rows == 0 and cols == 0`` is a no-op (matches openpyxl).
        Empty queue → byte-identical save.
        """
        from wolfxl.utils.cell import range_boundaries

        if not isinstance(rows, int) or isinstance(rows, bool):
            raise TypeError(
                f"move_range: rows must be an int, got {type(rows).__name__}"
            )
        if not isinstance(cols, int) or isinstance(cols, bool):
            raise TypeError(
                f"move_range: cols must be an int, got {type(cols).__name__}"
            )
        if not isinstance(cell_range, str):
            # Best-effort coercion for openpyxl `CellRange`-style objects.
            cell_range = str(cell_range)
        try:
            min_col, min_row, max_col, max_row = range_boundaries(cell_range)
        except Exception as exc:
            raise ValueError(
                f"move_range: cell_range must be a valid A1 range string, "
                f"got {cell_range!r}: {exc}"
            ) from exc
        if min_col is None or min_row is None or max_col is None or max_row is None:
            raise ValueError(
                f"move_range: cell_range must have all four corners "
                f"(rows + cols), got {cell_range!r}"
            )
        if rows == 0 and cols == 0:
            return
        # Validate destination bounds. Excel: rows 1..1_048_576,
        # cols 1..16_384 (1-based, inclusive).
        dst_min_row = min_row + rows
        dst_max_row = max_row + rows
        dst_min_col = min_col + cols
        dst_max_col = max_col + cols
        if dst_min_row < 1 or dst_max_row > 1_048_576:
            raise ValueError(
                f"move_range: destination row range "
                f"[{dst_min_row}, {dst_max_row}] is out of bounds "
                f"(must be in [1, 1048576])"
            )
        if dst_min_col < 1 or dst_max_col > 16_384:
            raise ValueError(
                f"move_range: destination column range "
                f"[{dst_min_col}, {dst_max_col}] is out of bounds "
                f"(must be in [1, 16384])"
            )
        self._workbook._pending_range_moves.append(  # noqa: SLF001
            (
                self.title,
                int(min_col),
                int(min_row),
                int(max_col),
                int(max_row),
                int(rows),
                int(cols),
                bool(translate),
            )
        )

    # ------------------------------------------------------------------
    # T1 PR1 read maps — lazy + per-sheet cached
    # ------------------------------------------------------------------
    #
    # Each method caches the fully-materialized result in a per-worksheet
    # slot on first access. That single FFI hop populates the cache;
    # subsequent calls (including per-cell lookups like ``cell.comment``)
    # are O(1) dict probes in Python. The Rust reader memoizes too, so
    # even a cold cache is cheap — but avoiding PyObject conversion on
    # every cell-level access is the dominant win here.

    def _get_comments_map(self) -> dict[str, Any]:
        """Return ``{cell_ref: Comment}`` for this sheet, cached per instance."""
        if self._comments_cache is not None:
            return self._comments_cache
        from wolfxl.comments import Comment

        wb = self._workbook
        if wb._rust_reader is None:  # noqa: SLF001
            self._comments_cache = {}
            return self._comments_cache
        try:
            entries = wb._rust_reader.read_comments(self._title)  # noqa: SLF001
        except Exception:
            entries = []
        result: dict[str, Any] = {}
        for e in entries:
            cell_ref = e.get("cell")
            if not cell_ref:
                continue
            result[cell_ref] = Comment(
                text=e.get("text", ""),
                author=e.get("author") or None,
            )
        self._comments_cache = result
        return result

    def _get_hyperlinks_map(self) -> dict[str, Any]:
        """Return ``{cell_ref: Hyperlink}`` for this sheet, cached per instance."""
        if self._hyperlinks_cache is not None:
            return self._hyperlinks_cache
        from wolfxl.worksheet.hyperlink import Hyperlink

        wb = self._workbook
        if wb._rust_reader is None:  # noqa: SLF001
            self._hyperlinks_cache = {}
            return self._hyperlinks_cache
        try:
            entries = wb._rust_reader.read_hyperlinks(self._title)  # noqa: SLF001
        except Exception:
            entries = []
        result: dict[str, Any] = {}
        for e in entries:
            cell_ref = e.get("cell")
            if not cell_ref:
                continue
            # Rust marks intra-workbook refs with ``internal=True`` and
            # puts the destination in ``target``. openpyxl splits the two:
            # external links use ``target``, internal use ``location``.
            is_internal = bool(e.get("internal", False))
            raw_target = e.get("target")
            hl = Hyperlink(
                ref=cell_ref,
                target=None if is_internal else raw_target,
                location=raw_target if is_internal else None,
                display=e.get("display") or None,
                tooltip=e.get("tooltip") or None,
            )
            result[cell_ref] = hl
        self._hyperlinks_cache = result
        return result

    # ------------------------------------------------------------------
    # T1 PR2 — Worksheet collections (tables, DVs, conditional formats).
    # Same lazy-cache contract as comments/hyperlinks: one Rust call per
    # sheet, dict/list-shaped cache for subsequent reads.
    # ------------------------------------------------------------------

    @property
    def tables(self) -> dict[str, Any]:
        """Return ``{table_name: Table}`` for this sheet.

        Loaded from the Rust reader on first access. In write mode the
        dict starts empty and is populated by ``add_table()``.
        """
        if self._tables_cache is not None:
            return self._tables_cache
        from wolfxl.worksheet.table import Table, TableColumn, TableStyleInfo

        wb = self._workbook
        if wb._rust_reader is None:  # noqa: SLF001
            self._tables_cache = {}
            return self._tables_cache
        try:
            entries = wb._rust_reader.read_tables(self._title)  # noqa: SLF001
        except Exception:
            entries = []
        result: dict[str, Any] = {}
        for e in entries:
            name = e.get("name") or e.get("displayName")
            if not name:
                continue
            style_name = e.get("style") or e.get("style_name")
            tsi = TableStyleInfo(
                name=style_name,
                showRowStripes=bool(e.get("show_row_stripes", False)),
                showColumnStripes=bool(e.get("show_column_stripes", False)),
                showFirstColumn=bool(e.get("show_first_column", False)),
                showLastColumn=bool(e.get("show_last_column", False)),
            ) if style_name is not None else None
            cols_raw = e.get("columns") or []
            tcols = [
                TableColumn(id=i + 1, name=str(c))
                for i, c in enumerate(cols_raw)
            ]
            result[name] = Table(
                name=name,
                displayName=e.get("displayName") or name,
                ref=e.get("ref", ""),
                headerRowCount=1 if e.get("header_row", True) else 0,
                totalsRowCount=1 if e.get("totals_row", False) else 0,
                tableStyleInfo=tsi,
                tableColumns=tcols,
            )
        self._tables_cache = result
        return result

    def add_table(self, table: Any) -> None:
        """Attach a table to this worksheet.

        Both write mode (``Workbook()``) and modify mode
        (``load_workbook(path, modify=True)``) queue the table on
        ``self._pending_tables``; the ``Workbook.save`` coordinator
        routes the queue to the right backend. RFC-024 (modify mode)
        flushes through the patcher's ``queue_table`` PyO3 setter,
        which in turn allocates a workbook-unique ``id`` at save time
        — any explicit ``id`` set by the user on the ``Table`` dataclass
        is ignored to avoid cross-sheet collisions in the rare case
        where the user opens a multi-table file and re-uses a numeric
        id by accident.
        """
        from wolfxl.worksheet.table import Table

        if not isinstance(table, Table):
            raise TypeError(
                f"add_table() expects a wolfxl.worksheet.table.Table, got {type(table).__name__}"
            )
        # Make sure the cache exists so ws.tables[name] sees the queued one.
        if self._tables_cache is None:
            self._tables_cache = {}
        self._tables_cache[table.name] = table
        self._pending_tables.append(table)

    @property
    def data_validations(self) -> Any:
        """Return the ``DataValidationList`` for this sheet (lazy-loaded)."""
        if self._data_validations_cache is not None:
            return self._data_validations_cache
        from wolfxl.worksheet.datavalidation import DataValidation, DataValidationList

        wb = self._workbook
        dvl = DataValidationList(ws=self)
        if wb._rust_reader is None:  # noqa: SLF001
            self._data_validations_cache = dvl
            return dvl
        try:
            entries = wb._rust_reader.read_data_validations(self._title)  # noqa: SLF001
        except Exception:
            entries = []
        for e in entries:
            dvl.dataValidation.append(DataValidation(
                type=e.get("validation_type") or e.get("type"),
                operator=e.get("operator"),
                formula1=e.get("formula1"),
                formula2=e.get("formula2"),
                allowBlank=bool(e.get("allow_blank", False)),
                showErrorMessage=bool(e.get("show_error_message", False)),
                showInputMessage=bool(e.get("show_input_message", False)),
                error=e.get("error"),
                errorTitle=e.get("error_title"),
                prompt=e.get("prompt"),
                promptTitle=e.get("prompt_title"),
                sqref=e.get("range") or e.get("sqref") or "",
            ))
        self._data_validations_cache = dvl
        return dvl

    def add_data_validation(self, dv: Any) -> None:
        """openpyxl-style alias for ``ws.data_validations.append(dv)``."""
        self.data_validations.append(dv)

    @property
    def conditional_formatting(self) -> Any:
        """Return the ``ConditionalFormattingList`` for this sheet."""
        if self._conditional_formatting_cache is not None:
            return self._conditional_formatting_cache
        from wolfxl.formatting import ConditionalFormatting, ConditionalFormattingList
        from wolfxl.formatting.rule import Rule

        wb = self._workbook
        cfl = ConditionalFormattingList(ws=self)
        if wb._rust_reader is None:  # noqa: SLF001
            self._conditional_formatting_cache = cfl
            return cfl
        try:
            entries = wb._rust_reader.read_conditional_formats(self._title)  # noqa: SLF001
        except Exception:
            entries = []
        # Group by sqref the way openpyxl does — same range = one entry.
        grouped: dict[str, list[Rule]] = {}
        order: list[str] = []
        for e in entries:
            sqref = e.get("range") or e.get("sqref") or ""
            if sqref not in grouped:
                grouped[sqref] = []
                order.append(sqref)
            formula = e.get("formula")
            if formula is None:
                formula_list: list[str] = []
            elif isinstance(formula, list):
                formula_list = [str(f) for f in formula]
            else:
                formula_list = [str(formula)]
            grouped[sqref].append(Rule(
                type=e.get("rule_type") or e.get("type") or "expression",
                operator=e.get("operator"),
                formula=formula_list,
                stopIfTrue=bool(e.get("stop_if_true", False)),
                priority=int(e.get("priority", 1)),
            ))
        for sqref in order:
            cfl._append_entry(ConditionalFormatting(sqref=sqref, rules=grouped[sqref]))  # noqa: SLF001
        self._conditional_formatting_cache = cfl
        return cfl

    # ------------------------------------------------------------------
    # Flush pending writes to Rust
    # ------------------------------------------------------------------

    def _flush(self) -> None:
        """Write all dirty cells to the Rust backend. Called by Workbook.save()."""
        from wolfxl._cell import (
            alignment_to_format_dict,
            border_to_rust_dict,
            fill_to_format_dict,
            font_to_format_dict,
            python_value_to_payload,
        )

        wb = self._workbook
        patcher = wb._rust_patcher  # noqa: SLF001
        writer = wb._rust_writer  # noqa: SLF001

        # Flush openpyxl compat properties to writer
        if writer is not None:
            self._flush_compat_properties(writer)

        if patcher is not None:
            # Modify mode: materialize buffers first (patcher has no batch
            # API), then flush dirty cells individually.
            if self._append_buffer:
                self._materialize_append_buffer()
            if self._bulk_writes:
                self._materialize_bulk_writes()
            self._flush_to_patcher(patcher, python_value_to_payload,
                                   font_to_format_dict, fill_to_format_dict,
                                   alignment_to_format_dict, border_to_rust_dict)
        elif writer is not None:
            self._flush_to_writer(writer, python_value_to_payload,
                                  font_to_format_dict, fill_to_format_dict,
                                  alignment_to_format_dict, border_to_rust_dict)
            # Sprint Ο Pod 1B (RFC-056) — autoFilter must be flushed
            # AFTER cells so the evaluator sees the populated grid.
            self._flush_autofilter_post_cells(writer)

        self._dirty.clear()

    def _flush_to_writer(
        self, writer: Any, python_value_to_payload: Any,
        font_to_format_dict: Any, fill_to_format_dict: Any,
        alignment_to_format_dict: Any, border_to_rust_dict: Any,
    ) -> None:
        """Flush dirty cells to the NativeWorkbook backend (write mode).

        Values are batched into a single ``write_sheet_values()`` call when
        possible (int/float/str/None), eliminating per-cell FFI overhead.
        Booleans, dates, datetimes, and formulas fall through to per-cell
        ``write_cell_value()`` with type-preserving payload dicts.
        """
        from wolfxl._cell import _UNSET

        # -- Fast path: flush the append buffer directly ----------------------
        # Rows added via append() are stored as raw Python lists — no Cell
        # objects.  We write them in one shot via write_sheet_values(), then
        # handle any non-batchable values (bool/date/formula) per-cell.
        if self._append_buffer:
            buf = self._append_buffer
            start_row = self._append_buffer_start
            start_a1 = rowcol_to_a1(start_row, 1)

            indiv_from_buf = self._extract_non_batchable(buf, start_row, 1)

            writer.write_sheet_values(self._title, start_a1, buf)

            for r, c, val in indiv_from_buf:
                coord = rowcol_to_a1(r, c)
                payload = python_value_to_payload(val)
                writer.write_cell_value(self._title, coord, payload)

            self._append_buffer = []

        # -- Flush bulk writes (write_rows) -----------------------------------
        for grid, sr, sc in self._bulk_writes:
            start_a1 = rowcol_to_a1(sr, sc)
            indiv_from_bulk = self._extract_non_batchable(grid, sr, sc)
            writer.write_sheet_values(self._title, start_a1, grid)
            for r, c, val in indiv_from_bulk:
                coord = rowcol_to_a1(r, c)
                payload = python_value_to_payload(val)
                writer.write_cell_value(self._title, coord, payload)
        self._bulk_writes = []

        # -- Partition dirty cells into batch-eligible values vs individual ----
        #
        # "batchable" = value is int | float | str | None (not bool, not
        #               formula strings starting with "=").  These go into a 2-D
        #               grid for a single write_sheet_values() call.
        #
        # Everything else (booleans, dates, formulas, format-dirty cells) is
        # handled per-cell so that type semantics and formatting are preserved.

        batch_vals: list[tuple[int, int, Any]] = []   # (row, col, raw_value)
        indiv_vals: list[tuple[int, int, Any]] = []   # (row, col, cell)
        format_cells: list[tuple[int, int, Any]] = [] # (row, col, cell)

        for row, col in self._dirty:
            cell = self._cells.get((row, col))
            if cell is None:
                continue

            if cell._value_dirty:  # noqa: SLF001
                val = cell._value  # noqa: SLF001
                if val is None or (
                    isinstance(val, (int, float, str))
                    and not isinstance(val, bool)
                    and not (isinstance(val, str) and val.startswith("="))
                ):
                    batch_vals.append((row, col, val))
                else:
                    indiv_vals.append((row, col, cell))

            if cell._format_dirty:  # noqa: SLF001
                format_cells.append((row, col, cell))

        # -- Batch write values -----------------------------------------------
        if batch_vals:
            min_r = batch_vals[0][0]
            min_c = batch_vals[0][1]
            max_r = min_r
            max_c = min_c
            for r, c, _ in batch_vals:
                if r < min_r:
                    min_r = r
                if r > max_r:
                    max_r = r
                if c < min_c:
                    min_c = c
                if c > max_c:
                    max_c = c

            num_rows = max_r - min_r + 1
            num_cols = max_c - min_c + 1

            grid: list[list[Any]] = [
                [None] * num_cols for _ in range(num_rows)
            ]
            for r, c, v in batch_vals:
                grid[r - min_r][c - min_c] = v

            start = rowcol_to_a1(min_r, min_c)
            writer.write_sheet_values(self._title, start, grid)

        # -- Per-cell value writes for non-batchable types --------------------
        from wolfxl.cell.cell import ArrayFormula, DataTableFormula
        from wolfxl.cell.rich_text import CellRichText

        for _row, _col, cell in indiv_vals:
            coord = rowcol_to_a1(cell._row, cell._col)  # noqa: SLF001
            val = cell._value  # noqa: SLF001
            if isinstance(val, ArrayFormula):
                # RFC-057: write-mode array-formula — the native writer
                # exposes ``write_cell_array_formula`` that emits
                # ``<f t="array" ref="..."/>`` directly.
                if hasattr(writer, "write_cell_array_formula"):
                    writer.write_cell_array_formula(
                        self._title,
                        coord,
                        {
                            "kind": "array",
                            "ref": val.ref,
                            "text": val.text,
                        },
                    )
                else:
                    # Fallback: emit as a regular formula — Excel will
                    # treat it as a single-cell formula (still computes
                    # correctly for spilling 365 functions).
                    payload = python_value_to_payload(f"={val.text}")
                    writer.write_cell_value(self._title, coord, payload)
                continue
            if isinstance(val, DataTableFormula):
                if hasattr(writer, "write_cell_array_formula"):
                    writer.write_cell_array_formula(
                        self._title,
                        coord,
                        {
                            "kind": "data_table",
                            "ref": val.ref,
                            "ca": val.ca,
                            "dt2D": val.dt2D,
                            "dtr": val.dtr,
                            "r1": val.r1,
                            "r2": val.r2,
                        },
                    )
                continue
            if isinstance(val, CellRichText):
                # Sprint Ι Pod-α: write-mode rich-text.  The native
                # writer doesn't expose a structured rich-text API yet,
                # so we flatten to plain text here and surface the
                # structured form via a separate `add_rich_text_cell`
                # call once the writer's model gains a RichText cell
                # type.  Until then, write-mode rich-text round-trips
                # only via modify mode.
                if hasattr(writer, "write_cell_rich_text"):
                    runs_payload = _cellrichtext_to_runs_payload(val)
                    writer.write_cell_rich_text(self._title, coord, runs_payload)
                else:
                    payload = python_value_to_payload(str(val))
                    writer.write_cell_value(self._title, coord, payload)
                continue
            payload = python_value_to_payload(val)
            writer.write_cell_value(self._title, coord, payload)

        # -- Spill-child placeholders ---------------------------------
        # RFC-057: cells inside an array's spill range that aren't the
        # master need a placeholder ``<c r="..."/>`` so Excel sees the
        # spill area pre-populated.  These are not in ``self._dirty``
        # because no Cell object was ever instantiated.
        for (sr, sc), (kind, _payload) in self._pending_array_formulas.items():
            if kind != "spill_child":
                continue
            if (sr, sc) in self._dirty:
                # The user explicitly assigned a value to this child;
                # skip the placeholder.
                continue
            coord = rowcol_to_a1(sr, sc)
            if hasattr(writer, "write_cell_array_formula"):
                writer.write_cell_array_formula(
                    self._title,
                    coord,
                    {"kind": "spill_child"},
                )

        # -- Batch format / border writes -----------------------------------------
        if format_cells:
            # Build format and border dicts for each cell, then figure out if
            # we can batch them into write_sheet_formats / write_sheet_borders.
            fmt_entries: list[tuple[int, int, dict[str, Any]]] = []
            bdr_entries: list[tuple[int, int, dict[str, Any]]] = []

            for _r, _c, cell in format_cells:
                r = cell._row  # noqa: SLF001
                c = cell._col  # noqa: SLF001
                fmt: dict[str, Any] = {}

                if cell._font is not _UNSET and cell._font is not None:  # noqa: SLF001
                    fmt.update(font_to_format_dict(cell._font))  # noqa: SLF001
                if cell._fill is not _UNSET and cell._fill is not None:  # noqa: SLF001
                    fmt.update(fill_to_format_dict(cell._fill))  # noqa: SLF001
                if cell._alignment is not _UNSET and cell._alignment is not None:  # noqa: SLF001
                    fmt.update(alignment_to_format_dict(cell._alignment))  # noqa: SLF001
                if cell._number_format is not _UNSET and cell._number_format is not None:  # noqa: SLF001
                    fmt["number_format"] = cell._number_format  # noqa: SLF001

                if fmt:
                    fmt_entries.append((r, c, fmt))

                if cell._border is not _UNSET and cell._border is not None:  # noqa: SLF001
                    bdict = border_to_rust_dict(cell._border)  # noqa: SLF001
                    if bdict:
                        bdr_entries.append((r, c, bdict))

            # Use batch API if there are enough entries to justify grid
            # construction overhead; otherwise per-cell is fine.
            if len(fmt_entries) > 1:
                self._batch_write_dicts(writer.write_sheet_formats, fmt_entries)
            else:
                for r, c, fmt in fmt_entries:
                    writer.write_cell_format(self._title, rowcol_to_a1(r, c), fmt)

            if len(bdr_entries) > 1:
                self._batch_write_dicts(writer.write_sheet_borders, bdr_entries)
            else:
                for r, c, bdict in bdr_entries:
                    writer.write_cell_border(self._title, rowcol_to_a1(r, c), bdict)

    def _batch_write_dicts(
        self,
        batch_fn: Any,
        entries: list[tuple[int, int, dict[str, Any]]],
    ) -> None:
        """Build a bounding-box grid of dicts and call a batch Rust method."""
        min_r = entries[0][0]
        min_c = entries[0][1]
        max_r = min_r
        max_c = min_c
        for r, c, _ in entries:
            if r < min_r:
                min_r = r
            if r > max_r:
                max_r = r
            if c < min_c:
                min_c = c
            if c > max_c:
                max_c = c

        num_rows = max_r - min_r + 1
        num_cols = max_c - min_c + 1
        grid: list[list[Any]] = [[None] * num_cols for _ in range(num_rows)]
        for r, c, d in entries:
            grid[r - min_r][c - min_c] = d

        start = rowcol_to_a1(min_r, min_c)
        batch_fn(self._title, start, grid)

    def _flush_to_patcher(
        self, patcher: Any, python_value_to_payload: Any,
        font_to_format_dict: Any, fill_to_format_dict: Any,
        alignment_to_format_dict: Any, border_to_rust_dict: Any,
    ) -> None:
        """Flush dirty cells to the XlsxPatcher backend (modify mode)."""
        from wolfxl._cell import _UNSET
        from wolfxl.cell.rich_text import CellRichText

        from wolfxl.cell.cell import ArrayFormula, DataTableFormula

        # RFC-057: emit spill-child placeholders too — they are NOT in
        # ``self._dirty`` because no Cell object was ever instantiated
        # for them, only the master cell triggered the
        # ``_populate_spill_placeholders`` map population.
        spill_children: set[tuple[int, int]] = {
            key
            for key, (kind, _payload) in self._pending_array_formulas.items()
            if kind == "spill_child" and key not in self._dirty
        }

        for row, col in self._dirty:
            cell = self._cells.get((row, col))
            if cell is None:
                continue
            coord = rowcol_to_a1(row, col)

            if cell._value_dirty:  # noqa: SLF001
                val = cell._value  # noqa: SLF001
                if isinstance(val, ArrayFormula):
                    patcher.queue_array_formula(
                        self._title,
                        coord,
                        {
                            "kind": "array",
                            "ref": val.ref,
                            "text": val.text,
                        },
                    )
                elif isinstance(val, DataTableFormula):
                    patcher.queue_array_formula(
                        self._title,
                        coord,
                        {
                            "kind": "data_table",
                            "ref": val.ref,
                            "ca": val.ca,
                            "dt2D": val.dt2D,
                            "dtr": val.dtr,
                            "r1": val.r1,
                            "r2": val.r2,
                        },
                    )
                elif isinstance(val, CellRichText):
                    runs_payload = _cellrichtext_to_runs_payload(val)
                    patcher.queue_rich_text_value(self._title, coord, runs_payload)
                else:
                    payload = python_value_to_payload(val)
                    patcher.queue_value(self._title, coord, payload)

            if cell._format_dirty:  # noqa: SLF001
                fmt: dict[str, Any] = {}

                if cell._font is not _UNSET and cell._font is not None:  # noqa: SLF001
                    fmt.update(font_to_format_dict(cell._font))  # noqa: SLF001
                if cell._fill is not _UNSET and cell._fill is not None:  # noqa: SLF001
                    fmt.update(fill_to_format_dict(cell._fill))  # noqa: SLF001
                if cell._alignment is not _UNSET and cell._alignment is not None:  # noqa: SLF001
                    fmt.update(alignment_to_format_dict(cell._alignment))  # noqa: SLF001
                if cell._number_format is not _UNSET and cell._number_format is not None:  # noqa: SLF001
                    fmt["number_format"] = cell._number_format  # noqa: SLF001

                if fmt:
                    patcher.queue_format(self._title, coord, fmt)

                if cell._border is not _UNSET and cell._border is not None:  # noqa: SLF001
                    bdict = border_to_rust_dict(cell._border)  # noqa: SLF001
                    if bdict:
                        patcher.queue_border(self._title, coord, bdict)

        for (row, col) in spill_children:
            coord = rowcol_to_a1(row, col)
            patcher.queue_array_formula(
                self._title,
                coord,
                {"kind": "spill_child"},
            )

    def _flush_autofilter_post_cells(self, writer: Any) -> None:
        """Sprint Ο Pod 1B (RFC-056) — flush the autoFilter to the
        Rust writer AFTER cells have been populated, so the
        evaluator sees the real grid and can stamp `row.hidden`
        flags on filtered-out rows.

        Modify mode goes through
        `Workbook._flush_pending_autofilters_to_patcher` instead.
        """
        sheet = self._title
        af = self._auto_filter
        af_has_state = (
            af.ref is not None
            or bool(af.filter_columns)
            or af.sort_state is not None
        )
        if af_has_state and hasattr(writer, "set_autofilter_native"):
            try:
                writer.set_autofilter_native(sheet, af.to_rust_dict())
            except Exception:
                # Defensive: don't poison the save path on a malformed
                # autofilter spec.
                pass

    def _flush_compat_properties(self, writer: Any) -> None:
        """Flush openpyxl compat properties (freeze_panes, dimensions, etc.)."""
        sheet = self._title

        # Freeze panes
        if self._freeze_panes is not None:
            writer.set_freeze_panes(
                sheet, {"mode": "freeze", "top_left_cell": self._freeze_panes},
            )

        # Row heights
        for row_num, height in self._row_heights.items():
            if height is not None:
                writer.set_row_height(sheet, row_num, height)

        # Column widths
        for col_letter, width in self._col_widths.items():
            if width is not None:
                writer.set_column_width(sheet, col_letter, width)

        # Print area (flush only if the Rust writer supports it)
        if self._print_area is not None and hasattr(writer, "set_print_area"):
            writer.set_print_area(sheet, self._print_area)

        # Sprint Ο Pod 1A.5 (RFC-055) — sheet-setup blocks on write
        # mode. Cheap probe first; only call set_sheet_setup_native
        # when at least one slot has been mutated by the user.
        if hasattr(writer, "set_sheet_setup_native"):
            has_setup = (
                self._page_setup is not None
                or self._page_margins is not None
                or self._header_footer is not None
                or self._sheet_view is not None
                or self._protection is not None
                or getattr(self, "_print_title_rows", None) is not None
                or getattr(self, "_print_title_cols", None) is not None
            )
            if has_setup:
                try:
                    payload = self.to_rust_setup_dict()
                    if any(v is not None for v in payload.values()):
                        writer.set_sheet_setup_native(sheet, payload)
                except Exception:
                    # Defensive: don't poison the save path on a
                    # malformed setup spec; the Python class
                    # validators should already have caught it.
                    pass

        # Sprint Π Pod Π-α (RFC-062) — page breaks + sheetFormatPr on
        # write mode. Cheap probe first; only call when at least one
        # slot has been mutated by the user.
        if hasattr(writer, "set_page_breaks_native"):
            has_breaks = (
                self._row_breaks is not None
                or self._col_breaks is not None
                or self._sheet_format is not None
            )
            if has_breaks:
                try:
                    breaks_dict = self.to_rust_page_breaks_dict()
                    fmt_dict = self.to_rust_sheet_format_dict()
                    payload = {
                        "row_breaks": breaks_dict.get("row_breaks"),
                        "col_breaks": breaks_dict.get("col_breaks"),
                        "sheet_format": fmt_dict,
                    }
                    if any(v is not None for v in payload.values()):
                        writer.set_page_breaks_native(sheet, payload)
                except Exception:
                    # Defensive: don't poison the save path.
                    pass

        # Sprint Ο Pod 1B (RFC-056) — autoFilter on write-mode sheets.
        # Modify mode goes through `Workbook._flush_pending_autofilters_to_patcher`
        # instead.
        af = self._auto_filter
        af_has_state = (
            af.ref is not None
            or bool(af.filter_columns)
            or af.sort_state is not None
        )
        # AutoFilter is flushed separately AFTER cells so the
        # evaluator can see the populated grid; see
        # `_flush_autofilter_post_cells`.

        # T1 PR4: cell-level write features — hyperlinks, comments.
        # Setters populate ``_pending_hyperlinks`` / ``_pending_comments`` and
        # the Rust writer already has ``add_hyperlink`` / ``add_comment`` —
        # we just translate the openpyxl-shaped dataclasses into the dict
        # shapes those methods expect.
        if self._pending_hyperlinks:
            for coord, hl in self._pending_hyperlinks.items():
                if hl is None:
                    # Explicit-delete sentinel — there's nothing to flush
                    # in write mode (no prior hyperlink existed). Modify
                    # mode would honor this, but that's a T1.5 path.
                    continue
                target = hl.target
                internal = False
                if target is None and hl.location is not None:
                    target = hl.location
                    internal = True
                if not target:
                    continue
                writer.add_hyperlink(sheet, {
                    "cell": coord,
                    "target": target,
                    "display": hl.display,
                    "tooltip": hl.tooltip,
                    "internal": internal,
                })
            self._pending_hyperlinks.clear()

        if self._pending_comments:
            for coord, c in self._pending_comments.items():
                if c is None:
                    continue
                writer.add_comment(sheet, {
                    "cell": coord,
                    "text": c.text,
                    "author": c.author,
                })
            self._pending_comments.clear()

        # T1 PR5: worksheet-level writes — tables, DVs, conditional formats.
        if self._pending_tables:
            for t in self._pending_tables:
                style_name = t.tableStyleInfo.name if t.tableStyleInfo else None
                col_names = [c.name for c in t.tableColumns] if t.tableColumns else []
                writer.add_table(sheet, {
                    "name": t.name,
                    "ref": t.ref,
                    "style": style_name,
                    "columns": col_names,
                    "header_row": t.headerRowCount > 0,
                    "totals_row": t.totalsRowCount > 0,
                })
            self._pending_tables.clear()

        if self._pending_data_validations:
            for dv in self._pending_data_validations:
                writer.add_data_validation(sheet, {
                    "range": dv.sqref,
                    "validation_type": dv.type,
                    "operator": dv.operator,
                    "formula1": dv.formula1,
                    "formula2": dv.formula2,
                    "allow_blank": dv.allowBlank,
                    "error_title": dv.errorTitle,
                    "error": dv.error,
                })
            self._pending_data_validations.clear()

        if self._pending_conditional_formats:
            for range_string, rule in self._pending_conditional_formats:
                formula = rule.formula[0] if rule.formula else None
                writer.add_conditional_format(sheet, {
                    "range": range_string,
                    "rule_type": rule.type,
                    "operator": rule.operator,
                    "formula": formula,
                    "stop_if_true": rule.stopIfTrue,
                })
            self._pending_conditional_formats.clear()

        # Sprint Λ Pod-β (RFC-045) — drain pending images.
        if self._pending_images and hasattr(writer, "add_image"):
            from wolfxl._images import image_to_writer_payload

            for img in self._pending_images:
                payload = image_to_writer_payload(img)
                writer.add_image(sheet, payload)
            self._pending_images.clear()

        # Sprint Μ Pod-β (RFC-046) — drain pending charts.
        # Pod-α ships ``add_chart_native`` on the Rust writer; until that
        # binding lands the queue is silently dropped (with a warning) so
        # existing tests that don't construct charts don't regress.
        if self._pending_charts:
            if hasattr(writer, "add_chart_native"):
                for chart in self._pending_charts:
                    payload = chart.to_rust_dict()
                    writer.add_chart_native(sheet, payload, chart._anchor)  # noqa: SLF001
            else:
                import warnings

                warnings.warn(
                    "wolfxl.chart: native chart write requires Pod-α's "
                    "add_chart_native binding (not yet available). "
                    f"Dropping {len(self._pending_charts)} chart(s) on "
                    f"sheet {sheet!r}.",
                    RuntimeWarning,
                    stacklevel=2,
                )
            self._pending_charts.clear()

    # ------------------------------------------------------------------
    # wolfxl-core classifier bridge (delegates to the single Rust
    # classifier that `wolfxl schema --format json` also goes through —
    # so Python callers and the CLI can never drift in their answers).
    # ------------------------------------------------------------------

    def classify_format(self, fmt: str) -> str:
        """Classify an Excel number-format string (e.g. ``"$#,##0.00"``).

        Returns the same category string the CLI's ``schema`` subcommand
        emits in the per-column ``format`` field: ``"general"``,
        ``"currency"``, ``"percentage"``, ``"scientific"``, ``"date"``,
        ``"time"``, ``"datetime"``, ``"integer"``, ``"float"``, or
        ``"text"``. The method is an
        instance method for discoverability; it doesn't use any
        worksheet state.
        """
        from wolfxl._rust import classify_format as _classify_format

        return _classify_format(fmt)

    def schema(self) -> dict[str, Any]:
        """Infer this worksheet's schema via ``wolfxl_core::infer_sheet_schema``.

        Returns a dict shaped exactly like one entry of
        ``wolfxl schema <file> --format json``'s ``sheets`` array:

        .. code-block:: python

            {
                "name": "Sheet1",
                "rows": 50,
                "columns": [
                    {"name": "Account", "type": "string",
                     "format": "general", "null_count": 0,
                     "unique_count": 12, "unique_capped": false,
                     "cardinality": "categorical",
                     "samples": ["Revenue", "COGS", ...]},
                    ...
                ],
            }

        Builds two parallel grids — values and per-cell
        ``number_format`` strings — from ``cell_records()`` so the
        bridge sees the same format metadata the CLI does. Without the
        format grid, currency / percentage / date columns would
        classify as ``general`` and the Python answer would silently
        drift from the CLI's. Pending in-memory ``number_format`` edits
        are overlaid before inference so unsaved worksheet changes are
        included too.
        """
        from wolfxl._cell import _UNSET
        from wolfxl._rust import infer_sheet_schema as _infer_sheet_schema

        max_row = self._max_row()
        max_col = self._max_col()
        values: list[list[Any]] = [[None] * max_col for _ in range(max_row)]
        fmts: list[list[str | None]] = [[None] * max_col for _ in range(max_row)]
        for record in self.iter_cell_records(
            include_format=True,
            include_empty=False,
            include_coordinate=False,
        ):
            r = int(record["row"]) - 1
            c = int(record["column"]) - 1
            if r >= max_row or c >= max_col:
                continue
            values[r][c] = record.get("value")
            nf = record.get("number_format")
            if nf:
                fmts[r][c] = nf
        for (row, col), cell in self._cells.items():
            if row > max_row or col > max_col:
                continue
            nf = cell._number_format
            if nf is not _UNSET and nf:
                fmts[row - 1][col - 1] = nf
        return _infer_sheet_schema(values, self._title, fmts)

    def __repr__(self) -> str:
        return f"<Worksheet [{self._title}]>"
