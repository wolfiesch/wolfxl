"""Streaming read path — Sprint Ι Pod-β.

Public entry-point: :func:`stream_iter_rows`. Activated by
``load_workbook(path, read_only=True)`` or auto-trigger when a sheet has
more than ``AUTO_STREAM_ROW_THRESHOLD`` rows. Wraps the Rust
``StreamingSheetReader`` and converts its row tuples into either:

- ``StreamingCell`` instances (mutation-rejected proxies that lazily look
  up styles via the eager workbook reader), or
- plain value tuples (when ``values_only=True``), padded by the configured
  column bounds.

The streaming path bypasses eager sheet materialization for the value scan but
still uses the eager workbook reader's style table for ``StreamingCell.font`` /
``.fill`` / etc. That table is loaded once on first style access and shared
across every cell in the iteration.
"""

from __future__ import annotations

from collections.abc import Iterator
import posixpath
import re
from typing import TYPE_CHECKING, Any
from xml.etree import ElementTree as ET
import zipfile

from wolfxl._utils import a1_to_rowcol
from wolfxl._utils import column_letter as _column_letter
from wolfxl._utils import rowcol_to_a1
from wolfxl._zip_safety import read_entry, validate_zipfile
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

_DIMENSION_REF_RE = re.compile(rb"<(?:[A-Za-z0-9_]+:)?dimension\b[^>]*\bref=[\"']([^\"']+)[\"']")
_ROW_TAG_RE = re.compile(rb"<(?:[A-Za-z0-9_]+:)?row(?:\s|>)")


def _local_name(tag: str) -> str:
    if "}" in tag:
        return tag.rsplit("}", 1)[1]
    return tag.rsplit(":", 1)[-1]


def _resolve_package_target(base_dir: str, target: str) -> str:
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(posixpath.join(base_dir, target))


def _sheet_path_from_workbook(zf: zipfile.ZipFile, sheet_title: str) -> str | None:
    try:
        workbook_xml = read_entry(zf, "xl/workbook.xml")
        rels_xml = read_entry(zf, "xl/_rels/workbook.xml.rels")
    except (KeyError, OSError, zipfile.BadZipFile):
        return None

    rid_by_sheet: dict[str, str] = {}
    try:
        root = ET.fromstring(workbook_xml)
        for elem in root.iter():
            if _local_name(elem.tag) != "sheet":
                continue
            if elem.attrib.get("name") != sheet_title:
                continue
            for key, value in elem.attrib.items():
                if _local_name(key) == "id":
                    rid_by_sheet[sheet_title] = value
                    break
    except ET.ParseError:
        return None
    rid = rid_by_sheet.get(sheet_title)
    if rid is None:
        return None

    try:
        rels_root = ET.fromstring(rels_xml)
    except ET.ParseError:
        return None
    for rel in rels_root.iter():
        if _local_name(rel.tag) != "Relationship":
            continue
        if rel.attrib.get("Id") != rid:
            continue
        target = rel.attrib.get("Target")
        if not target:
            return None
        return _resolve_package_target("xl", target)
    return None


def _dimension_ref_from_sheet_head(zf: zipfile.ZipFile, sheet_path: str) -> str | None:
    try:
        with zf.open(sheet_path) as fh:
            head = bytearray()
            while len(head) < 131_072:
                chunk = fh.read(8192)
                if not chunk:
                    break
                head.extend(chunk)
                match = _DIMENSION_REF_RE.search(head)
                if match is not None:
                    return match.group(1).decode("ascii", errors="ignore")
                if b"<sheetData" in head or b"<sheetData>" in head:
                    break
    except (KeyError, OSError, zipfile.BadZipFile):
        return None
    return None


def _row_count_exceeds_threshold(zf: zipfile.ZipFile, sheet_path: str) -> bool:
    try:
        with zf.open(sheet_path) as fh:
            tail = b""
            count = 0
            while True:
                chunk = fh.read(65_536)
                if not chunk:
                    break
                buf = tail + chunk
                count += len(_ROW_TAG_RE.findall(buf))
                if count > AUTO_STREAM_ROW_THRESHOLD:
                    return True
                tail = buf[-32:]
    except (KeyError, OSError, zipfile.BadZipFile):
        return False
    return False


def _max_row_from_dimension_ref(ref: str) -> int | None:
    cell = ref.split(":", 1)[-1].replace("$", "")
    try:
        return a1_to_rowcol(cell)[0]
    except ValueError:
        return None


def _bounds_from_dimension_ref(ref: str) -> tuple[int, int, int, int] | None:
    parts = ref.replace("$", "").split(":", 1)
    if len(parts) == 1:
        parts = [parts[0], parts[0]]
    try:
        min_row, min_col = a1_to_rowcol(parts[0])
        max_row, max_col = a1_to_rowcol(parts[1])
    except ValueError:
        return None
    return (
        min(min_row, max_row),
        min(min_col, max_col),
        max(min_row, max_row),
        max(min_col, max_col),
    )


def _source_dimension_bounds(path: str, sheet_title: str) -> tuple[int, int, int, int] | None:
    try:
        with zipfile.ZipFile(path) as zf:
            validate_zipfile(zf)
            sheet_path = _sheet_path_from_workbook(zf, sheet_title)
            if sheet_path is None:
                return None
            ref = _dimension_ref_from_sheet_head(zf, sheet_path)
            return _bounds_from_dimension_ref(ref) if ref is not None else None
    except (OSError, zipfile.BadZipFile):
        return None


def _source_dimension_max_row(path: str, sheet_title: str) -> int | None:
    try:
        with zipfile.ZipFile(path) as zf:
            validate_zipfile(zf)
            sheet_path = _sheet_path_from_workbook(zf, sheet_title)
            if sheet_path is None:
                return None
            ref = _dimension_ref_from_sheet_head(zf, sheet_path)
            bounds = _bounds_from_dimension_ref(ref) if ref is not None else None
            max_row = bounds[2] if bounds is not None else None
            if max_row is not None:
                return max_row
            if _row_count_exceeds_threshold(zf, sheet_path):
                return AUTO_STREAM_ROW_THRESHOLD + 1
            return None
    except (OSError, zipfile.BadZipFile):
        return None


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
    style lookups can defer to the workbook reader's ``read_cell_format`` path.
    Style attributes are
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
    # Style lookups — defer to the eager workbook reader.
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
            return payload.get("number_format") or "General"
        return "General"

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

    def __eq__(self, other: object) -> bool:
        return (
            getattr(other, "coordinate", None) == self.coordinate
            and getattr(other, "value", None) == self.value
        )


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
    source_bounds = _source_dimension_bounds(path, ws.title)
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

    def _resolved_cmax(cells: list[tuple[Any, ...]]) -> int:
        if mx_c is not None:
            return mx_c
        if source_bounds is not None:
            return source_bounds[3]
        return max((int(cell[0]) for cell in cells), default=(mn_c or 1) - 1)

    try:
        if values_only:
            counter = mn_r if mn_r is not None else 1
            # Switch to ``read_next_row`` rather than ``read_next_values``
            # so we have access to each cell's style_id. The slight extra
            # work (vs Rust-side tuple padding) is offset by skipping a
            # second materialization of the tuple in Python.
            while True:
                row = reader.read_next_row()
                if row is None:
                    break
                row_idx, cells = row
                cmin = mn_c if mn_c is not None else 1
                cmax = _resolved_cmax(cells)
                empty_row = (None,) * max(0, cmax + 1 - cmin)
                while counter < row_idx:
                    yield empty_row
                    counter += 1
                if cmax < cmin:
                    yield ()
                    counter = row_idx + 1
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
                counter = row_idx + 1
            if mx_r is not None:
                cmin = mn_c if mn_c is not None else 1
                cmax = mx_c if mx_c is not None else (
                    source_bounds[3] if source_bounds is not None else cmin - 1
                )
                empty_row = (None,) * max(0, cmax + 1 - cmin)
                while counter <= mx_r:
                    yield empty_row
                    counter += 1
        else:
            counter = mn_r if mn_r is not None else 1
            while True:
                row = reader.read_next_row()
                if row is None:
                    break
                row_idx, cells = row
                # Determine emitted column bounds: explicit min/max wins,
                # else span observed cells.
                cmin = mn_c if mn_c is not None else 1
                cmax = _resolved_cmax(cells)
                while counter < row_idx:
                    empty_row = tuple(
                        StreamingCell(ws, counter, col, None, None, "blank")
                        for col in range(cmin, cmax + 1)
                    )
                    yield empty_row
                    counter += 1
                if cmax < cmin:
                    yield ()
                    counter = row_idx + 1
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
                counter = row_idx + 1
            if mx_r is not None:
                cmin = mn_c if mn_c is not None else 1
                cmax = mx_c if mx_c is not None else (
                    source_bounds[3] if source_bounds is not None else cmin - 1
                )
                while counter <= mx_r:
                    yield tuple(
                        StreamingCell(ws, counter, col, None, None, "blank")
                        for col in range(cmin, cmax + 1)
                    )
                    counter += 1
    finally:
        reader.close()


def should_auto_stream(ws: Worksheet) -> bool:
    """Heuristic: auto-engage streaming for ``iter_rows`` on huge sheets.

    Triggers when the sheet's ``max_row`` exceeds
    :data:`AUTO_STREAM_ROW_THRESHOLD`. This deliberately consults the
    worksheet ``<dimension ref=...>`` tag directly from the ZIP head instead
    of ``Worksheet._max_row()``, because the normal dimensions API may parse
    and cache the full sheet before streaming can engage.
    """
    wb = ws._workbook  # noqa: SLF001
    path = getattr(wb, "_source_path", None)
    if not path:
        return False
    rows = _source_dimension_max_row(path, ws.title)
    return rows is not None and rows > AUTO_STREAM_ROW_THRESHOLD
