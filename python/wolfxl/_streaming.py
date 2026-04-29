"""Streaming read path — Sprint Ι Pod-β.

Public entry-point: :func:`stream_iter_rows`. Activated by
``load_workbook(path, read_only=True)`` or auto-trigger when a sheet has
more than ``AUTO_STREAM_ROW_THRESHOLD`` rows. Wraps the Rust
``StreamingSheetReader`` and converts its row tuples into either:

- ``StreamingCell`` instances (mutation-rejected proxies that lazily look
  up styles via the existing ``CalamineStyledBook`` reader), or
- plain value tuples (when ``values_only=True``), padded by the configured
  column bounds.

The streaming path bypasses ``calamine-styles``' eager sheet
materialization for the value scan but still uses the
``CalamineStyledBook`` style table for ``StreamingCell.font`` /
``.fill`` / etc. — that table is loaded once on first style access and
shared across every cell in the iteration.
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from wolfxl._utils import column_letter as _column_letter
from wolfxl._utils import rowcol_to_a1
from wolfxl.utils.datetime import from_excel
from wolfxl.utils.numbers import is_date_format

if TYPE_CHECKING:
    from wolfxl._styles import Alignment, Border, Font, PatternFill
    from wolfxl._worksheet import Worksheet


#: Row count above which ``iter_rows`` transparently uses the streaming
#: path even when the workbook was opened without ``read_only=True``.
#: Tuned to keep the eager path for small/medium workbooks (where the
#: bulk-read FFI is fastest) while still scaling to multi-million-cell
#: sheets without exhausting RSS.
AUTO_STREAM_ROW_THRESHOLD = 50_000


def _streaming_value(payload: Any) -> Any:
    """Normalize a Rust streaming payload into the same Python types
    that the eager value-reader emits.

    The Rust streaming layer surfaces:
      - plain ``str``/``int``/``float``/``bool`` for the common types,
      - a ``dict`` with ``type=formula|error`` for formulas/errors,
      - the raw ISO string for ``t="d"`` cells (Excel rarely emits these).
    """
    if isinstance(payload, dict):
        t = payload.get("type")
        if t == "formula":
            # openpyxl returns the formula string from `Cell.value` in
            # read mode, so do the same here.
            return payload.get("formula")
        if t == "error":
            return payload.get("value")
    return payload


def _maybe_datetime_from_serial(value: Any, num_fmt: str | None) -> Any:
    """Convert an Excel serial to a ``datetime``/``date``/``time`` when
    ``num_fmt`` is a date-typed format string (Sprint Λ Pod-γ).

    Mirrors openpyxl's read-only path: numeric cells with a date number
    format surface as Python datetimes rather than raw serials. Non-date
    formats and non-numeric values pass through unchanged.
    """
    if not isinstance(value, (int, float)) or isinstance(value, bool):
        return value
    if not is_date_format(num_fmt):
        return value
    try:
        return from_excel(value)
    except Exception:
        # Out-of-range serial / unrepresentable timedelta — fall back to
        # the raw serial so the row keeps flowing rather than erroring
        # mid-iteration. Matches openpyxl's behavior on bad serials.
        return value


class StreamingCell:
    """Read-only cell proxy yielded by streaming ``iter_rows()``.

    Holds a snapshot of the value (parsed by the Rust SAX scanner) plus
    the originating row/column and a reference back to the Worksheet so
    style lookups can defer to the existing
    ``CalamineStyledBook.read_cell_format`` path. Style attributes are
    fully featured (font, fill, border, alignment, number_format) and
    behave identically to the eager ``Cell`` properties — the difference
    is mutation: every setter raises ``RuntimeError``.

    The class is ``__slots__``-based so a 1M-cell scan doesn't blow the
    Python heap with per-cell ``__dict__`` storage.
    """

    __slots__ = ("_ws", "_row", "_col", "_value", "_style_id", "_cell_type")

    def __init__(
        self,
        ws: Worksheet,
        row: int,
        col: int,
        value: Any,
        style_id: int | None,
        cell_type: str,
    ) -> None:
        self._ws = ws
        self._row = row
        self._col = col
        self._value = _streaming_value(value)
        self._style_id = style_id
        self._cell_type = cell_type

    # ------------------------------------------------------------------
    # Coordinate accessors — match openpyxl's read_only Cell API.
    # ------------------------------------------------------------------

    @property
    def value(self) -> Any:
        """Return the streaming cell value, converting date serials when needed."""
        # Sprint Λ Pod-γ: numeric cells whose number format is date-typed
        # surface as Python datetime, mirroring openpyxl's read_only path
        # (and the eager Cell.value path, which converts via calamine).
        if isinstance(self._value, (int, float)) and not isinstance(self._value, bool):
            return _maybe_datetime_from_serial(self._value, self.number_format)
        return self._value

    @property
    def row(self) -> int:
        """Return this cell's 1-based row index."""
        return self._row

    @property
    def column(self) -> int:
        """Return this cell's 1-based column index."""
        return self._col

    @property
    def column_letter(self) -> str:
        """Return this cell's Excel column letter."""
        return _column_letter(self._col)

    @property
    def coordinate(self) -> str:
        """Return this cell's A1-style coordinate."""
        return rowcol_to_a1(self._row, self._col)

    @property
    def parent(self) -> Worksheet:
        """Return the containing worksheet."""
        return self._ws

    @property
    def data_type(self) -> str:
        """Map the SAX ``t=`` token onto openpyxl's ``data_type`` letters."""
        return {
            "n": "n",
            "s": "s",
            "str": "s",
            "inlineStr": "s",
            "b": "b",
            "e": "e",
            "d": "d",
            "formula": "f",
            "blank": "n",
        }.get(self._cell_type, "n")

    # ------------------------------------------------------------------
    # Style lookups — defer to the eager CalamineStyledBook reader.
    # The streaming path does NOT re-implement xl/styles.xml parsing;
    # the styles table is small (~KB) regardless of sheet size and the
    # eager reader caches it after first access.
    # ------------------------------------------------------------------

    @property
    def font(self) -> Font:
        """Return the resolved cell font."""
        from wolfxl._cell import _format_to_font

        wb = self._ws._workbook  # noqa: SLF001
        reader = wb._rust_reader  # noqa: SLF001
        if reader is None:
            from wolfxl._styles import Font as _Font

            return _Font()
        payload = reader.read_cell_format(self._ws.title, self.coordinate)
        return _format_to_font(payload)

    @property
    def fill(self) -> PatternFill:
        """Return the resolved cell fill."""
        from wolfxl._cell import _format_to_fill

        wb = self._ws._workbook  # noqa: SLF001
        reader = wb._rust_reader  # noqa: SLF001
        if reader is None:
            from wolfxl._styles import PatternFill as _PF

            return _PF()
        payload = reader.read_cell_format(self._ws.title, self.coordinate)
        return _format_to_fill(payload)

    @property
    def border(self) -> Border:
        """Return the resolved cell border."""
        from wolfxl._cell import _border_payload_to_border

        wb = self._ws._workbook  # noqa: SLF001
        reader = wb._rust_reader  # noqa: SLF001
        if reader is None:
            from wolfxl._styles import Border as _B

            return _B()
        payload = reader.read_cell_border(self._ws.title, self.coordinate)
        return _border_payload_to_border(payload)

    @property
    def alignment(self) -> Alignment:
        """Return the resolved cell alignment."""
        from wolfxl._cell import _format_to_alignment

        wb = self._ws._workbook  # noqa: SLF001
        reader = wb._rust_reader  # noqa: SLF001
        if reader is None:
            from wolfxl._styles import Alignment as _A

            return _A()
        payload = reader.read_cell_format(self._ws.title, self.coordinate)
        return _format_to_alignment(payload)

    @property
    def number_format(self) -> str | None:
        """Return the resolved number format string."""
        wb = self._ws._workbook  # noqa: SLF001
        reader = wb._rust_reader  # noqa: SLF001
        if reader is None:
            return None
        payload = reader.read_cell_format(self._ws.title, self.coordinate)
        if isinstance(payload, dict):
            return payload.get("number_format")
        return None

    # ------------------------------------------------------------------
    # Mutation — strictly rejected. Sprint Ι Pod-β contract.
    #
    # We don't bother defining `@value.setter` etc.; ``__setattr__``
    # below is the single chokepoint that rejects every assignment and
    # carries a uniform error message. The trade-off: a user who
    # registers a property setter via inheritance can't override the
    # rejection — by design, since StreamingCell is final-by-contract.
    # ------------------------------------------------------------------

    def _reject(self, what: str) -> None:
        raise RuntimeError(
            f"read_only=True: cannot set {what} on streaming cell "
            f"{self.coordinate}; reload without read_only or use modify=True"
        )

    def __setattr__(self, name: str, value: Any) -> None:
        # __slots__ permits assignment to declared slot names; reject
        # any others up-front so a typo'd attribute doesn't silently
        # succeed via setter resolution. The constructor is the only
        # legitimate caller for slot writes.
        if name in StreamingCell.__slots__:
            object.__setattr__(self, name, value)
            return
        self._reject(name)

    def __repr__(self) -> str:
        """Return a compact debug representation for this streaming cell."""
        return f"<StreamingCell {self.coordinate} value={self._value!r}>"


def _resolve_bounds(
    ws: Worksheet,
    min_row: int | None,
    max_row: int | None,
    min_col: int | None,
    max_col: int | None,
) -> tuple[int | None, int | None, int | None, int | None]:
    """Pass-through that just clamps user ``None``s to ``None`` (the Rust
    layer treats ``None`` as 'unbounded'). Kept as a seam so future
    sheet-dimension probing can land here without churning callers.
    """
    return min_row, max_row, min_col, max_col


def stream_iter_rows(
    ws: Worksheet,
    min_row: int | None = None,
    max_row: int | None = None,
    min_col: int | None = None,
    max_col: int | None = None,
    values_only: bool = False,
) -> Iterator[tuple[Any, ...]]:
    """SAX-streaming generator backing ``iter_rows`` in read-only / large-sheet mode.

    Yields:
        - When ``values_only=True``: plain tuples of cell values padded to
          ``(min_col..max_col)`` width. Missing cells become ``None``.
          When no column bound is supplied, the tuple is sized by the
          actual cells present in that row.
        - Otherwise: tuples of :class:`StreamingCell` instances covering
          the same range.
    """
    from wolfxl import _rust

    wb = ws._workbook  # noqa: SLF001
    path = getattr(wb, "_source_path", None)
    if path is None:
        raise RuntimeError(
            "stream_iter_rows requires a source-path-bearing workbook "
            "(use load_workbook to obtain one)"
        )
    mn_r, mx_r, mn_c, mx_c = _resolve_bounds(ws, min_row, max_row, min_col, max_col)
    reader = _rust.StreamingSheetReader.open(
        path, ws.title, mn_r, mx_r, mn_c, mx_c
    )

    # Sprint Λ Pod-γ: cache style_id → (number_format, is_date) so we
    # resolve a date format once per distinct style rather than once per
    # cell. Sentinel `_NO_STYLE` covers the (style_id is None) case.
    style_date_cache: dict[int | None, tuple[str | None, bool]] = {}
    rust_reader = wb._rust_reader  # noqa: SLF001

    def _is_date_style(style_id: int | None, row_idx: int, col: int) -> bool:
        cached = style_date_cache.get(style_id)
        if cached is not None:
            return cached[1]
        if style_id is None or rust_reader is None:
            style_date_cache[style_id] = (None, False)
            return False
        # The Rust reader exposes number_format only via cell coordinate.
        # The lookup is keyed off the cell's style_id internally, so any
        # cell that shares this style_id will resolve to the same format
        # — caching by style_id keeps this O(unique-styles) rather than
        # O(cells).
        try:
            payload = rust_reader.read_cell_format(
                ws.title, rowcol_to_a1(row_idx, col)
            )
        except Exception:
            style_date_cache[style_id] = (None, False)
            return False
        num_fmt = payload.get("number_format") if isinstance(payload, dict) else None
        is_date = is_date_format(num_fmt)
        style_date_cache[style_id] = (num_fmt, is_date)
        return is_date

    try:
        if values_only:
            # Switch to ``read_next_row`` rather than ``read_next_values``
            # so we have access to each cell's style_id. The slight extra
            # work (vs Rust-side tuple padding) is offset by skipping a
            # second materialization of the tuple in Python.
            while True:
                row = reader.read_next_row()
                if row is None:
                    break
                row_idx, cells = row
                if mn_c is not None and mx_c is not None:
                    cmin, cmax = mn_c, mx_c
                else:
                    if not cells:
                        yield ()
                        continue
                    cols = [c[0] for c in cells]
                    cmin = mn_c if mn_c is not None else min(cols)
                    cmax = mx_c if mx_c is not None else max(cols)
                if cmax < cmin:
                    yield ()
                    continue
                by_col = {c[0]: c for c in cells}
                row_out: list[Any] = []
                for col in range(cmin, cmax + 1):
                    rec = by_col.get(col)
                    if rec is None:
                        row_out.append(None)
                        continue
                    _, value, style_id, _cell_type = rec
                    py_val = _streaming_value(value) if isinstance(value, dict) else value
                    if (
                        isinstance(py_val, (int, float))
                        and not isinstance(py_val, bool)
                        and _is_date_style(style_id, row_idx, col)
                    ):
                        py_val = _maybe_datetime_from_serial(
                            py_val, style_date_cache[style_id][0]
                        )
                    row_out.append(py_val)
                yield tuple(row_out)
        else:
            while True:
                row = reader.read_next_row()
                if row is None:
                    break
                row_idx, cells = row
                # Determine emitted column bounds: explicit min/max wins,
                # else span observed cells.
                if mn_c is not None and mx_c is not None:
                    cmin, cmax = mn_c, mx_c
                else:
                    if not cells:
                        yield ()
                        continue
                    cols = [c[0] for c in cells]
                    cmin = mn_c if mn_c is not None else min(cols)
                    cmax = mx_c if mx_c is not None else max(cols)
                if cmax < cmin:
                    yield ()
                    continue
                by_col = {c[0]: c for c in cells}
                row_out: list[Any] = []
                for col in range(cmin, cmax + 1):
                    rec = by_col.get(col)
                    if rec is None:
                        row_out.append(
                            StreamingCell(ws, row_idx, col, None, None, "blank")
                        )
                    else:
                        _, value, style_id, cell_type = rec
                        row_out.append(
                            StreamingCell(ws, row_idx, col, value, style_id, cell_type)
                        )
                yield tuple(row_out)
    finally:
        reader.close()


def should_auto_stream(ws: Worksheet) -> bool:
    """Heuristic: auto-engage streaming for ``iter_rows`` on huge sheets.

    Triggers when the sheet's ``max_row`` exceeds
    :data:`AUTO_STREAM_ROW_THRESHOLD`. Cheap to evaluate because
    ``Worksheet._max_row`` either consults the dimension-tag fast path
    (single XML head probe) or returns a cached value.
    """
    try:
        rows = ws._max_row()  # noqa: SLF001
    except Exception:
        return False
    return rows is not None and rows > AUTO_STREAM_ROW_THRESHOLD
