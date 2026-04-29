"""Worksheet proxy — provides ``ws['A1']`` access and tracks dirty cells."""

from __future__ import annotations

from collections.abc import Iterable, Iterator
from typing import TYPE_CHECKING, Any

from wolfxl._cell import Cell
from wolfxl._worksheet_access import (
    get_col_tuple,
    get_item,
    get_rect,
    get_row_tuple,
    resolve_string_key,
)
from wolfxl._worksheet_bounds import (
    calculate_dimension as _calculate_dimension,
    max_col as _worksheet_max_col,
    max_row as _worksheet_max_row,
    read_dimension_bounds as _read_dimension_bounds,
    read_dimensions as _read_dimensions,
)
from wolfxl._worksheet_collections import AutoFilter as _AutoFilter
from wolfxl._worksheet_collections import MergedCellsProxy as _MergedCellsProxy
from wolfxl._worksheet_dimensions import ColumnDimensionProxy, RowDimensionProxy
from wolfxl._worksheet_features import (
    get_comments_map,
    get_conditional_formatting,
    get_data_validations,
    get_hyperlinks_map,
    get_tables_map,
)
from wolfxl._worksheet_flush import (
    flush_autofilter_post_cells,
    flush_compat_properties,
)
from wolfxl._worksheet_iteration import (
    iter_cols as _iter_cols,
    iter_cols_bulk as _iter_cols_bulk,
    iter_rows as _iter_rows,
    iter_rows_bulk as _iter_rows_bulk,
)
from wolfxl._worksheet_media import (
    add_chart as _add_chart,
    add_image as _add_image,
    add_pivot_table as _add_pivot_table,
    add_slicer as _add_slicer,
    remove_chart as _remove_chart,
    replace_chart as _replace_chart,
    validate_a1_anchor as _validate_a1_anchor,
)
from wolfxl._worksheet_pending import collect_pending_overlay, pending_writes_bounds
from wolfxl._worksheet_patcher_flush import flush_to_patcher
from wolfxl._worksheet_records import (
    cached_formula_values as _cached_formula_values,
    canonical_data_type as _canonical_data_type,  # noqa: F401 - legacy import path
    iter_cell_records as _iter_cell_records,
    iter_cell_records_python as _iter_cell_records_python,
    schema as _infer_worksheet_schema,
    sheet_visibility as _sheet_visibility,
)
from wolfxl._worksheet_setup import (
    get_col_breaks,
    get_dimension_holder,
    get_freeze_panes,
    get_header_footer,
    get_page_margins,
    get_page_setup,
    get_protection,
    get_row_breaks,
    get_sheet_format,
    get_sheet_view,
    set_freeze_panes,
    set_print_title_cols,
    set_print_title_rows,
    to_rust_page_breaks_dict as _to_rust_page_breaks_dict,
    to_rust_setup_dict as _to_rust_setup_dict,
    to_rust_sheet_format_dict as _to_rust_sheet_format_dict,
)
from wolfxl._worksheet_structural import (
    delete_cols as _delete_cols,
    delete_rows as _delete_rows,
    insert_cols as _insert_cols,
    insert_rows as _insert_rows,
    move_range as _move_range,
)
from wolfxl._worksheet_writer_flush import flush_to_writer
from wolfxl._worksheet_write_buffers import (
    batch_write_dicts,
    extract_non_batchable,
    materialize_append_buffer,
    materialize_bulk_writes,
)

def _cellrichtext_to_runs_payload(crt: Any) -> list[tuple[str, dict[str, Any] | None]]:
    """Sprint Ι Pod-α: convert a ``CellRichText`` into the Rust-side
    payload (a list of ``(text, font_dict_or_None)`` tuples).

    This lives at module scope so both the write-mode and modify-mode
    flush paths can share it.  The Rust side reconstructs runs via
    ``py_runs_to_rust`` in ``src/wolfxl/patcher_payload.rs``.
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


class Worksheet:
    """Openpyxl-shaped proxy for a worksheet in a :class:`Workbook`.

    The proxy presents cell access, sheet metadata, and feature collections
    while deferring reads and writes to the workbook's active backend. In
    modify mode, mutations are queued and flushed through the patcher on save.
    """

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
        """Worksheet title shown in the workbook tab list."""
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
        return get_freeze_panes(self)

    @freeze_panes.setter
    def freeze_panes(self, value: str | None) -> None:
        """Set the freeze panes cell reference.

        Args:
            value: The top-left scrollable cell, such as ``"B2"``, or
                ``None`` to clear frozen panes.
        """
        set_freeze_panes(self, value)

    # ------------------------------------------------------------------
    # Sprint Ο Pod 1A (RFC-055) — print / view / protection accessors
    # ------------------------------------------------------------------

    @property
    def page_setup(self) -> Any:
        """Lazy ``PageSetup`` accessor (RFC-055 §2.1)."""
        return get_page_setup(self)

    @page_setup.setter
    def page_setup(self, value: Any) -> None:
        """Replace the worksheet page setup block.

        Args:
            value: Page setup object compatible with WolfXL's page setup
                serializer.
        """
        self._page_setup = value

    @property
    def page_margins(self) -> Any:
        """Lazy ``PageMargins`` accessor (RFC-055 §2.2)."""
        return get_page_margins(self)

    @page_margins.setter
    def page_margins(self, value: Any) -> None:
        """Replace the worksheet page margins block.

        Args:
            value: Page margins object compatible with WolfXL's page margins
                serializer.
        """
        self._page_margins = value

    @property
    def HeaderFooter(self) -> Any:  # noqa: N802 - openpyxl alias
        """Openpyxl camel-case alias for :attr:`header_footer`."""
        return self.header_footer

    @property
    def header_footer(self) -> Any:
        """Lazy ``HeaderFooter`` accessor (RFC-055 §2.3)."""
        return get_header_footer(self)

    @header_footer.setter
    def header_footer(self, value: Any) -> None:
        """Replace the worksheet header/footer block.

        Args:
            value: Header/footer object compatible with WolfXL's
                header/footer serializer.
        """
        self._header_footer = value

    @property
    def sheet_view(self) -> Any:
        """Lazy ``SheetView`` accessor (RFC-055 §2.5).

        ``ws.freeze_panes`` mutations are mirrored into ``sheet_view.pane``
        on the setter side; on the getter side, if a sheet view has a
        non-None pane we surface that as ``freeze_panes`` for parity.
        """
        return get_sheet_view(self)

    @sheet_view.setter
    def sheet_view(self, value: Any) -> None:
        """Replace the worksheet view block.

        Args:
            value: Sheet view object compatible with WolfXL's sheet view
                serializer.
        """
        self._sheet_view = value

    @property
    def protection(self) -> Any:
        """Lazy ``SheetProtection`` accessor (RFC-055 §2.6)."""
        return get_protection(self)

    @protection.setter
    def protection(self, value: Any) -> None:
        """Replace the worksheet protection block.

        Args:
            value: Sheet protection object compatible with WolfXL's sheet
                protection serializer.
        """
        self._protection = value

    # ------------------------------------------------------------------
    # Sprint Π Pod Π-α (RFC-062) — page breaks + sheet format props
    # ------------------------------------------------------------------

    @property
    def row_breaks(self) -> Any:
        """Lazy ``PageBreakList`` of horizontal page breaks (RFC-062 §3)."""
        return get_row_breaks(self)

    @row_breaks.setter
    def row_breaks(self, value: Any) -> None:
        """Replace the horizontal page-break collection.

        Args:
            value: Page break collection for row breaks.
        """
        self._row_breaks = value

    @property
    def col_breaks(self) -> Any:
        """Lazy ``PageBreakList`` of vertical page breaks (RFC-062 §3)."""
        return get_col_breaks(self)

    @col_breaks.setter
    def col_breaks(self, value: Any) -> None:
        """Replace the vertical page-break collection.

        Args:
            value: Page break collection for column breaks.
        """
        self._col_breaks = value

    @property
    def page_breaks(self) -> Any:
        """openpyxl alias — ``ws.page_breaks`` is a row-breaks alias."""
        return self.row_breaks

    @page_breaks.setter
    def page_breaks(self, value: Any) -> None:
        """Set the openpyxl row-break alias.

        Args:
            value: Page break collection assigned to ``row_breaks``.
        """
        self._row_breaks = value

    @property
    def sheet_format(self) -> Any:
        """Lazy ``SheetFormatProperties`` accessor (RFC-062 §3)."""
        return get_sheet_format(self)

    @sheet_format.setter
    def sheet_format(self, value: Any) -> None:
        """Replace the sheet format properties block.

        Args:
            value: Sheet format properties object compatible with WolfXL's
                sheet format serializer.
        """
        self._sheet_format = value

    @property
    def dimension_holder(self) -> Any:
        """Return a fresh ``DimensionHolder`` view bound to this worksheet."""
        return get_dimension_holder(self)

    def to_rust_page_breaks_dict(self) -> dict[str, Any]:
        """Return the §10 dict shape for ``<rowBreaks>`` / ``<colBreaks>``.

        Each side is ``None`` when the corresponding ``PageBreakList``
        is un-touched OR carries zero breaks — the patcher / writer
        then knows to skip emitting the corresponding XML block.
        """
        return _to_rust_page_breaks_dict(self)

    def to_rust_sheet_format_dict(self) -> dict[str, Any] | None:
        """Return the §10 dict for ``<sheetFormatPr>`` or ``None``.

        Returns ``None`` when the wrapper is un-touched OR at all-default
        values — the writer then keeps the legacy hardcoded
        ``<sheetFormatPr defaultRowHeight="15"/>`` emit path.
        """
        return _to_rust_sheet_format_dict(self)

    @property
    def print_title_rows(self) -> str | None:
        """Repeat-rows for printing (RFC-055 §2.4)."""
        return self._print_title_rows

    @print_title_rows.setter
    def print_title_rows(self, value: str | None) -> None:
        """Set repeat rows for printed pages.

        Args:
            value: Row range string such as ``"1:3"``, or ``None`` to clear
                repeat rows.
        """
        set_print_title_rows(self, value)

    @property
    def print_title_cols(self) -> str | None:
        """Repeat-cols for printing (RFC-055 §2.4)."""
        return self._print_title_cols

    @print_title_cols.setter
    def print_title_cols(self, value: str | None) -> None:
        """Set repeat columns for printed pages.

        Args:
            value: Column range string such as ``"A:C"``, or ``None`` to clear
                repeat columns.
        """
        set_print_title_cols(self, value)

    def to_rust_setup_dict(self) -> dict[str, Any]:
        """Return the §10 dict contract for the Rust patcher / writer.

        Returns ``None`` for any sub-block whose Python wrapper is at
        its construction defaults — the Rust side then knows to skip
        emitting the corresponding XML.
        """
        return _to_rust_setup_dict(self)

    @property
    def auto_filter(self) -> _AutoFilter:
        """Worksheet auto-filter proxy."""
        return self._auto_filter

    @property
    def row_dimensions(self) -> RowDimensionProxy:
        """Dict-like row-dimension accessor keyed by 1-based row number."""
        return RowDimensionProxy(self)

    @property
    def column_dimensions(self) -> ColumnDimensionProxy:
        """Dict-like column-dimension accessor keyed by column letter."""
        return ColumnDimensionProxy(self)

    @property
    def print_area(self) -> str | None:
        """Get/set the print area range string (e.g. ``'A1:D10'``).

        Stored locally and flushed to the Rust writer on ``save()`` if the
        writer supports ``set_print_area()``.
        """
        return self._print_area

    @print_area.setter
    def print_area(self, value: str | None) -> None:
        """Set the worksheet print area.

        Args:
            value: A worksheet range string, or ``None`` to clear the print
                area.
        """
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
        return get_item(self, key)

    def _resolve_string_key(self, key: str) -> Any:
        """Resolve a string key to Cell / tuple / tuple-of-tuples."""
        return resolve_string_key(self, key)

    def _get_rect(
        self, min_row: int, min_col: int, max_row: int, max_col: int,
    ) -> tuple[tuple[Cell, ...], ...]:
        """Return a 2D tuple of Cells for the inclusive rectangle."""
        return get_rect(self, min_row, min_col, max_row, max_col)

    def _get_row_tuple(
        self, min_row: int, max_row: int,
    ) -> tuple[tuple[Cell, ...], ...]:
        """Return a tuple of row-tuples for rows min_row..max_row inclusive."""
        return get_row_tuple(self, min_row, max_row)

    def _get_col_tuple(
        self,
        min_col: int,
        max_col: int,
        min_row: int | None = None,
        max_row: int | None = None,
    ) -> tuple[tuple[Cell, ...], ...]:
        """Return a tuple of column-tuples for cols min_col..max_col inclusive."""
        return get_col_tuple(self, min_col, max_col, min_row, max_row)

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
        materialize_append_buffer(self)

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
        materialize_bulk_writes(self)

    @staticmethod
    def _extract_non_batchable(
        grid: list[list[Any]], start_row: int, start_col: int,
    ) -> list[tuple[int, int, Any]]:
        """Extract non-batchable values from grid, replacing them with None.

        Non-batchable: booleans, formulas (str starting with '='), and
        non-primitive types (dates, datetimes, etc.).  These require
        per-cell ``write_cell_value()`` calls with type-preserving payloads.
        """
        return extract_non_batchable(grid, start_row, start_col)

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
        yield from _iter_rows(
            self,
            min_row=min_row,
            max_row=max_row,
            min_col=min_col,
            max_col=max_col,
            values_only=values_only,
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
        yield from _iter_cols(
            self,
            min_col=min_col,
            max_col=max_col,
            min_row=min_row,
            max_row=max_row,
            values_only=values_only,
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
        yield from _iter_cols_bulk(self, min_col, max_col, min_row, max_row)

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
        yield from _iter_rows_bulk(self, min_row, max_row, min_col, max_col)

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
        yield from _iter_cell_records(
            self,
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
        )

    def _collect_pending_overlay(self) -> dict[tuple[int, int], Any]:
        """Return ``{(row, col): value}`` for cells modified since the last save.

        Includes explicit cell edits (anything in ``_dirty``), the append
        buffer, and bulk-write grids. Returns an empty dict when nothing is
        pending — the Rust read path stays a hot, allocation-free loop.
        """
        return collect_pending_overlay(self)

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
        return _cached_formula_values(self, qualified=qualified)

    def sheet_visibility(self) -> dict[str, Any]:
        """Return hidden rows/columns and outline levels for this sheet.

        Row and column identifiers are 1-based to mirror openpyxl's dimension
        collections. The returned shape is:
        ``hidden_rows``, ``hidden_columns``, ``row_outline_levels``, and
        ``column_outline_levels``.
        """
        return _sheet_visibility(self)

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
        yield from _iter_cell_records_python(
            self,
            min_row=min_row,
            max_row=max_row,
            min_col=min_col,
            max_col=max_col,
            include_empty=include_empty,
            include_coordinate=include_coordinate,
        )

    def calculate_dimension(self) -> str:
        """Return the used worksheet range in openpyxl's ``A1:C10`` form."""
        return _calculate_dimension(self)

    def _read_dimension_bounds(self) -> tuple[int, int, int, int] | None:
        """Return 1-based ``(min_row, min_col, max_row, max_col)`` bounds.

        Modify mode has both a Rust reader (for the on-disk extents) and
        Python-side pending writes (cells/append buffer/bulk writes). The
        reported bounds must be the union, otherwise callers that derive
        ranges from ``calculate_dimension()`` miss unsaved edits.
        """
        return _read_dimension_bounds(self)

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
        return pending_writes_bounds(self)

    def _read_dimensions(self) -> tuple[int, int]:
        """Discover sheet dimensions from the Rust backend (read mode only)."""
        return _read_dimensions(self)

    def _max_row(self) -> int:
        return _worksheet_max_row(self)

    def _max_col(self) -> int:
        return _worksheet_max_col(self)

    # openpyxl exposes these as properties, not methods. Mirror that contract
    # so ``ws.max_row`` (no parens) works as a drop-in for openpyxl callers.
    # Pinned by ``tests/parity/test_read_parity.py`` (uses ``op_ws.max_row``).
    @property
    def max_row(self) -> int:
        """Largest row index visible to openpyxl-style callers."""
        return self._max_row()

    @property
    def max_column(self) -> int:
        """Largest column index visible to openpyxl-style callers."""
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
        _add_chart(self, chart, anchor)

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
        _add_pivot_table(self, pivot_table)

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
        _add_slicer(self, slicer, anchor)

    @staticmethod
    def _validate_a1_anchor(anchor: str) -> None:
        """Raise :class:`ValueError` if *anchor* is not a valid A1 cell ref.

        Per RFC-046 §10.11.2: ``r"^[A-Z]+[0-9]+$"`` — single cell
        coordinates only (e.g. ``"E15"``, ``"AA200"``). Range refs and
        sheet-qualified refs are rejected; pass an anchor object for
        more complex placements. Excel's column max is ``XFD`` (16384)
        and row max is 1048576; refs outside those bounds raise.
        """
        _validate_a1_anchor(anchor)

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
        _remove_chart(self, chart)

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
        _replace_chart(self, old, new)

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
        _add_image(self, img, anchor)

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
        _insert_rows(self, idx, amount)

    def delete_rows(self, idx: int, amount: int = 1) -> None:
        """Delete *amount* rows starting at *idx*, shifting subsequent rows up.

        Implements RFC-030. Validates ``idx >= 1`` and ``amount >= 1``;
        raises ``ValueError`` otherwise. Refs that point INTO the
        deleted band become ``#REF!`` per OOXML semantics.
        """
        _delete_rows(self, idx, amount)

    def insert_cols(self, idx: int | str, amount: int = 1) -> None:
        """Shift columns right to insert *amount* empty columns at *idx*.

        Implements RFC-031. ``idx`` may be a 1-based int or an Excel
        column letter (``"A"``, ``"AB"``, ...). Validates ``idx >= 1``
        and ``amount >= 0``; ``amount == 0`` is a noop. Queues a
        col-shift op on the owning workbook's ``_pending_axis_shifts``.

        See ``Plans/rfcs/031-insert-delete-cols.md`` for full semantics
        (formula shift, anchor shift, ``<col>`` span split).
        """
        _insert_cols(self, idx, amount)

    def delete_cols(self, idx: int | str, amount: int = 1) -> None:
        """Delete *amount* columns starting at *idx*, shifting subsequent columns left.

        Implements RFC-031. ``idx`` may be a 1-based int or an Excel
        column letter. Refs that point INTO the deleted band become
        ``#REF!`` per OOXML semantics. ``amount == 0`` is a noop.
        """
        _delete_cols(self, idx, amount)

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
        _move_range(self, cell_range, rows=rows, cols=cols, translate=translate)

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
        return get_comments_map(self)

    def _get_hyperlinks_map(self) -> dict[str, Any]:
        """Return ``{cell_ref: Hyperlink}`` for this sheet, cached per instance."""
        return get_hyperlinks_map(self)

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
        return get_tables_map(self)

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
        return get_data_validations(self)

    def add_data_validation(self, dv: Any) -> None:
        """openpyxl-style alias for ``ws.data_validations.append(dv)``."""
        self.data_validations.append(dv)

    @property
    def conditional_formatting(self) -> Any:
        """Return the ``ConditionalFormattingList`` for this sheet."""
        return get_conditional_formatting(self)

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
        flush_to_writer(
            self,
            writer,
            python_value_to_payload,
            font_to_format_dict,
            fill_to_format_dict,
            alignment_to_format_dict,
            border_to_rust_dict,
            _cellrichtext_to_runs_payload,
        )

    def _batch_write_dicts(
        self,
        batch_fn: Any,
        entries: list[tuple[int, int, dict[str, Any]]],
    ) -> None:
        """Build a bounding-box grid of dicts and call a batch Rust method."""
        batch_write_dicts(self, batch_fn, entries)

    def _flush_to_patcher(
        self, patcher: Any, python_value_to_payload: Any,
        font_to_format_dict: Any, fill_to_format_dict: Any,
        alignment_to_format_dict: Any, border_to_rust_dict: Any,
    ) -> None:
        """Flush dirty cells to the XlsxPatcher backend (modify mode)."""
        flush_to_patcher(
            self,
            patcher,
            python_value_to_payload,
            font_to_format_dict,
            fill_to_format_dict,
            alignment_to_format_dict,
            border_to_rust_dict,
            _cellrichtext_to_runs_payload,
        )

    def _flush_autofilter_post_cells(self, writer: Any) -> None:
        """Flush write-mode auto-filter metadata after cell values."""
        flush_autofilter_post_cells(self, writer)

    def _flush_compat_properties(self, writer: Any) -> None:
        """Flush openpyxl compatibility metadata to the write-mode backend."""
        flush_compat_properties(self, writer)

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
        return _infer_worksheet_schema(self)

    def __repr__(self) -> str:
        """Return a compact debug representation for this worksheet.

        Returns:
            A string containing the worksheet title.
        """
        return f"<Worksheet [{self._title}]>"
