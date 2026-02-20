"""Worksheet proxy — provides ``ws['A1']`` access and tracks dirty cells."""

from __future__ import annotations

from collections.abc import Iterable, Iterator
from typing import TYPE_CHECKING, Any

from wolfxl._cell import Cell
from wolfxl._utils import a1_to_rowcol, rowcol_to_a1

if TYPE_CHECKING:
    from wolfxl._workbook import Workbook


class Worksheet:
    """Proxy for a single worksheet in a Workbook."""

    __slots__ = (
        "_workbook", "_title", "_cells", "_dirty", "_dimensions",
        "_max_col_idx", "_next_append_row",
        "_append_buffer", "_append_buffer_start", "_bulk_writes",
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
    # Cell access
    # ------------------------------------------------------------------

    def __getitem__(self, key: str) -> Cell:
        """``ws['A1']`` -> Cell."""
        row, col = a1_to_rowcol(key)
        return self._get_or_create_cell(row, col)

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

        # Build an A1:B2-style range string for Rust.
        r_min = min_row or 1
        r_max = max_row or self._max_row()
        c_min = min_col or 1
        c_max = max_col or self._max_col()
        range_str = f"{rowcol_to_a1(r_min, c_min)}:{rowcol_to_a1(r_max, c_max)}"

        # Prefer plain-value read (no dict overhead) if available.
        use_plain = hasattr(reader, "read_sheet_values_plain")
        if use_plain:
            rows = reader.read_sheet_values_plain(sheet, range_str)
        else:
            rows = reader.read_sheet_values(sheet, range_str)

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

    def _read_dimensions(self) -> tuple[int, int]:
        """Discover sheet dimensions from the Rust backend (read mode only)."""
        if self._dimensions is not None:
            return self._dimensions
        wb = self._workbook
        if wb._rust_reader is None:  # noqa: SLF001
            self._dimensions = (1, 1)
            return self._dimensions
        rows = wb._rust_reader.read_sheet_values(self._title, None)  # noqa: SLF001
        if not rows or not isinstance(rows, list):
            self._dimensions = (1, 1)
            return self._dimensions
        max_r = len(rows)
        max_c = max((len(row) for row in rows), default=1)
        self._dimensions = (max_r, max_c)
        return self._dimensions

    def _max_row(self) -> int:
        if self._workbook._rust_reader is not None:  # noqa: SLF001
            return self._read_dimensions()[0]
        if not self._cells:
            return 1
        return max(k[0] for k in self._cells)

    def _max_col(self) -> int:
        if self._workbook._rust_reader is not None:  # noqa: SLF001
            return self._read_dimensions()[1]
        if not self._cells:
            return 1
        return max(k[1] for k in self._cells)

    # ------------------------------------------------------------------
    # Write-mode helpers
    # ------------------------------------------------------------------

    def merge_cells(self, range_string: str) -> None:
        """Merge cells (write mode only). Example: ``ws.merge_cells('A1:B2')``."""
        wb = self._workbook
        if wb._rust_writer is None:  # noqa: SLF001
            raise RuntimeError("merge_cells requires write mode")
        wb._rust_writer.merge_cells(self._title, range_string)  # noqa: SLF001

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
        """Flush dirty cells to the RustXlsxWriterBook backend (write mode).

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

    def __repr__(self) -> str:
        return f"<Worksheet [{self._title}]>"
