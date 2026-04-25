"""Worksheet proxy — provides ``ws['A1']`` access and tracks dirty cells."""

from __future__ import annotations

import datetime as _dt
from collections.abc import Iterable, Iterator
from typing import TYPE_CHECKING, Any

from wolfxl._cell import Cell
from wolfxl._utils import a1_to_rowcol, column_index, rowcol_to_a1
from wolfxl.utils.cell import column_index_from_string, range_boundaries


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
    """Proxy for ``ws.auto_filter.ref = 'A1:D10'``."""

    __slots__ = ("_ref",)

    def __init__(self) -> None:
        self._ref: str | None = None

    @property
    def ref(self) -> str | None:
        return self._ref

    @ref.setter
    def ref(self, value: str | None) -> None:
        self._ref = value


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
        self._pending_tables: list[Any] = []
        self._pending_data_validations: list[Any] = []
        self._pending_conditional_formats: list[tuple[str, Any]] = []

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
        """Iterate over rows in a range. Matches openpyxl's iter_rows API."""
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
        """Empty list - matches openpyxl's default for sheets without charts.

        Some code paths (e.g. ``len(ws._charts)``) read this even on
        sheets that weren't constructed from a chart-bearing file.
        """
        return []

    @property
    def _images(self) -> list[Any]:
        """Empty list - same rationale as ``_charts``."""
        return []

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

        Tracked by RFC-030 (Phase 4 / WolfXL 1.1). See
        ``Plans/rfcs/030-insert-delete-rows.md`` for the implementation plan.
        """
        raise NotImplementedError(
            "Worksheet.insert_rows is scheduled for WolfXL 1.1 (RFC-030). "
            "See Plans/rfcs/030-insert-delete-rows.md for the implementation plan. "
            "Workaround: use openpyxl for structural ops, then load the result "
            "with wolfxl.load_workbook() to do the heavy reads."
        )

    def delete_rows(self, idx: int, amount: int = 1) -> None:
        """Delete *amount* rows starting at *idx*, shifting subsequent rows up.

        Tracked by RFC-030 (Phase 4 / WolfXL 1.1). See
        ``Plans/rfcs/030-insert-delete-rows.md`` for the implementation plan.
        """
        raise NotImplementedError(
            "Worksheet.delete_rows is scheduled for WolfXL 1.1 (RFC-030). "
            "See Plans/rfcs/030-insert-delete-rows.md for the implementation plan. "
            "Workaround: use openpyxl for structural ops, then load the result "
            "with wolfxl.load_workbook() to do the heavy reads."
        )

    def insert_cols(self, idx: int, amount: int = 1) -> None:
        """Shift columns right to insert *amount* empty columns at *idx*.

        Tracked by RFC-031 (Phase 4 / WolfXL 1.1). See
        ``Plans/rfcs/031-insert-delete-cols.md`` for the implementation plan.
        """
        raise NotImplementedError(
            "Worksheet.insert_cols is scheduled for WolfXL 1.1 (RFC-031). "
            "See Plans/rfcs/031-insert-delete-cols.md for the implementation plan. "
            "Workaround: use openpyxl for structural ops, then load the result "
            "with wolfxl.load_workbook() to do the heavy reads."
        )

    def delete_cols(self, idx: int, amount: int = 1) -> None:
        """Delete *amount* columns starting at *idx*, shifting subsequent columns left.

        Tracked by RFC-031 (Phase 4 / WolfXL 1.1). See
        ``Plans/rfcs/031-insert-delete-cols.md`` for the implementation plan.
        """
        raise NotImplementedError(
            "Worksheet.delete_cols is scheduled for WolfXL 1.1 (RFC-031). "
            "See Plans/rfcs/031-insert-delete-cols.md for the implementation plan. "
            "Workaround: use openpyxl for structural ops, then load the result "
            "with wolfxl.load_workbook() to do the heavy reads."
        )

    def move_range(
        self,
        cell_range: Any,
        rows: int = 0,
        cols: int = 0,
        translate: bool = False,
    ) -> None:
        """Move a rectangular block of cells by *rows* / *cols*.

        Tracked by RFC-034 (Phase 4 / WolfXL 1.1). See
        ``Plans/rfcs/034-move-range.md`` for the implementation plan.
        """
        raise NotImplementedError(
            "Worksheet.move_range is scheduled for WolfXL 1.1 (RFC-034). "
            "See Plans/rfcs/034-move-range.md for the implementation plan. "
            "Workaround: use openpyxl for structural ops, then load the result "
            "with wolfxl.load_workbook() to do the heavy reads."
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
        for _row, _col, cell in indiv_vals:
            coord = rowcol_to_a1(cell._row, cell._col)  # noqa: SLF001
            payload = python_value_to_payload(cell._value)  # noqa: SLF001
            writer.write_cell_value(self._title, coord, payload)

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

        for row, col in self._dirty:
            cell = self._cells.get((row, col))
            if cell is None:
                continue
            coord = rowcol_to_a1(row, col)

            if cell._value_dirty:  # noqa: SLF001
                payload = python_value_to_payload(cell._value)  # noqa: SLF001
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

        # T1 PR4: cell-level write features — hyperlinks, comments.
        # Setters populate ``_pending_hyperlinks`` / ``_pending_comments`` and
        # the Rust writer already has ``add_hyperlink`` / ``add_comment`` —
        # we just translate the openpyxl-shaped dataclasses into the dict
        # shapes those methods expect.
        if self._pending_hyperlinks:
            for coord, hl in self._pending_hyperlinks.items():
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
