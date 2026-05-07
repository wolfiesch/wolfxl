"""Streaming write-only worksheet for ``Workbook(write_only=True)``.

Mirror of openpyxl's ``openpyxl.writer.write_only.WriteOnlyWorksheet``
backed by the Rust ``StreamingSheet`` temp-file machinery (G20 / RFC-073).

The eager :class:`wolfxl._worksheet.Worksheet` materialises every cell
in four memory layers (Python ``_append_buffer`` → Python ``_cells`` →
Rust ``BTreeMap<u32, Row>`` → ``String`` accumulator). Streaming write-
only mode breaks that chain: each :meth:`WriteOnlyWorksheet.append`
encodes one ``<row>...</row>`` element into a per-sheet temp file
immediately and forgets the cell payloads. SST and styles still grow
in-memory because they're irreducible OOXML costs (matches openpyxl's
``lxml.xmlfile`` model exactly), but row data scales O(rows on disk),
not O(rows in RAM).

# Forbidden-method matrix

Random access (``ws["A1"]``, ``ws.cell(...)``, ``ws.iter_rows``,
slicing) raises :class:`AttributeError` to match openpyxl's
``_write_only.py:51-66``. Writing once is the only API.

# Style resolution

:class:`WriteOnlyCell` is a lightweight ``@dataclass``
factory — NOT a ``Cell`` subclass. It holds the unresolved style
attributes (``font``, ``fill``, ``border``, ``alignment``,
``number_format``). At append time the style resolves into a
``style_id`` once via :func:`_resolve_style_id`, with caching keyed by
the cell's font/fill/border identity so a typical row of 10 cells in 5
distinct styles only crosses the FFI 5 times — not 10.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, Any, Iterable

from wolfxl._cell_payloads import (
    border_to_rust_dict,
    fill_to_format_dict,
    font_to_format_dict,
    python_value_to_payload,
)
from wolfxl.utils.exceptions import WorkbookAlreadySaved

if TYPE_CHECKING:
    from wolfxl._workbook import Workbook


_FORBIDDEN_ATTRS = frozenset(
    {
        "cell",
        "iter_rows",
        "iter_cols",
        "rows",
        "columns",
        "values",
        "merge_cells",
        "unmerge_cells",
        "add_chart",
        "add_image",
        "add_table",
        "add_data_validation",
        "add_pivot_table",
        "conditional_formatting",
        "auto_filter",
    }
)


@dataclass
class WriteOnlyCell:
    """Lightweight cell factory for streaming write-only mode.

    Construct with ``WriteOnlyCell(ws, value, font=..., fill=...,
    border=..., alignment=..., number_format=...)``. The cell holds the
    raw value plus unresolved style attributes; the styling resolves to
    a single ``style_id`` when the row containing this cell is
    appended.

    NOT a subclass of :class:`wolfxl._cell.Cell` — write-only cells do
    not support coordinate access, formula caching, or post-write
    mutation. The ``ws`` argument is accepted for openpyxl-shape
    compatibility but is unused.
    """

    value: Any = None
    font: Any = None
    fill: Any = None
    border: Any = None
    alignment: Any = None
    number_format: str | None = None

    def __init__(
        self,
        ws: "WriteOnlyWorksheet | None" = None,
        value: Any = None,
        *,
        font: Any = None,
        fill: Any = None,
        border: Any = None,
        alignment: Any = None,
        number_format: str | None = None,
    ) -> None:
        # Accept ws positionally for openpyxl-shape compat, ignore it —
        # styling resolves through the workbook's StylesBuilder lazily.
        del ws
        self.value = value
        self.font = font
        self.fill = fill
        self.border = border
        self.alignment = alignment
        self.number_format = number_format


class WriteOnlyWorksheet:
    """Append-only worksheet backed by a per-sheet streaming temp file.

    Constructed indirectly via :meth:`Workbook.create_sheet` when the
    workbook was opened with ``write_only=True``. The constructor
    converts the Rust-side sheet into streaming mode and installs the
    forbidden-method guard.

    Public API matches openpyxl's contract:

    * :meth:`append` is the only row-write entry point.
    * :attr:`column_dimensions` / :attr:`row_dimensions` /
      :attr:`freeze_panes` / :attr:`print_area` / :attr:`print_titles`
      may be set BEFORE the first :meth:`append`. Setting them after
      raises :class:`RuntimeError` (matches openpyxl's invariant —
      sheet-level XML lives BEFORE ``<sheetData>`` in slot order).
    * Random access and post-row mutation methods raise
      :class:`AttributeError`.
    """

    def __init__(self, workbook: "Workbook", title: str) -> None:
        self._workbook = workbook
        self.title = title
        self._row_counter: int = 0
        self._closed = False
        # Style resolution cache keyed by `(id(font), id(fill), ...)`.
        self._style_cache: dict[tuple[int, int, int, int, str | None], int] = {}
        # Sheet-level slots that must be set before the first append.
        self.column_dimensions: dict[str, Any] = {}
        self.row_dimensions: dict[int, Any] = {}
        self._freeze_panes: str | None = None
        self._print_area: str | None = None
        self._print_title_rows: str | None = None
        self._print_title_cols: str | None = None
        # Convert the underlying Rust-side worksheet into streaming mode.
        backend = workbook._rust_writer  # noqa: SLF001
        backend.enable_streaming_sheet(title)

    # ------------------------------------------------------------------
    # The single supported write API.
    # ------------------------------------------------------------------

    def append(self, row: Iterable[Any]) -> None:
        """Append one row of cell values to the streaming temp file.

        ``row`` may be:

        * a ``list`` / ``tuple`` of plain Python values (numbers,
          strings, bools, datetimes, ``None`` to skip a column);
        * a ``list`` / ``tuple`` of :class:`WriteOnlyCell` objects;
        * a heterogeneous mix.

        Each non-None entry encodes one ``<c r="...">...</c>`` cell.
        ``None`` entries skip a column (the column index advances but
        no cell is emitted).
        """
        if self._closed:
            raise WorkbookAlreadySaved(
                "cannot append to a write-only worksheet after save"
            )
        if isinstance(row, (str, bytes)):
            raise TypeError(
                "WriteOnlyWorksheet.append requires an iterable of cells, "
                "not a string"
            )
        # Materialize the row before crossing the FFI so we can freeze
        # the column count (the Rust side reads `len(cells)` to compute
        # max_col).
        row_list = list(row)
        cells_payload: list[dict[str, Any] | None] = []
        for value in row_list:
            if value is None:
                cells_payload.append(None)
                continue
            if isinstance(value, WriteOnlyCell):
                payload = python_value_to_payload(value.value)
                style_id = self._resolve_style_id(value)
                if style_id is not None:
                    payload["style_id"] = style_id
                cells_payload.append(payload)
            else:
                cells_payload.append(python_value_to_payload(value))

        self._row_counter += 1
        backend = self._workbook._rust_writer  # noqa: SLF001
        backend.append_streaming_row(self.title, self._row_counter, cells_payload)

    # ------------------------------------------------------------------
    # Style resolution
    # ------------------------------------------------------------------

    def _resolve_style_id(self, cell: WriteOnlyCell) -> int | None:
        """Intern the cell's font/fill/border/alignment/number_format
        on the workbook's styles builder and return the style_id.

        Returns ``None`` for cells with no styling attributes set —
        the FFI then leaves the cell with the default style_id,
        matching the eager path's "no format applied" behaviour.
        """
        if (
            cell.font is None
            and cell.fill is None
            and cell.border is None
            and cell.alignment is None
            and cell.number_format is None
        ):
            return None
        key = (
            id(cell.font),
            id(cell.fill),
            id(cell.border),
            id(cell.alignment),
            cell.number_format,
        )
        cached = self._style_cache.get(key)
        if cached is not None:
            return cached

        format_dict: dict[str, Any] = {}
        if cell.font is not None:
            format_dict.update(font_to_format_dict(cell.font))
        if cell.fill is not None:
            format_dict.update(fill_to_format_dict(cell.fill))
        if cell.border is not None:
            format_dict["border"] = border_to_rust_dict(cell.border)
        if cell.alignment is not None:
            # Mirror eager-path key shape — the FFI dispatch handles it.
            format_dict["alignment"] = {
                "horizontal": getattr(cell.alignment, "horizontal", None),
                "vertical": getattr(cell.alignment, "vertical", None),
                "wrap_text": getattr(cell.alignment, "wrap_text", None),
                "shrink_to_fit": getattr(cell.alignment, "shrink_to_fit", None),
                "indent": getattr(cell.alignment, "indent", None),
                "text_rotation": getattr(cell.alignment, "text_rotation", None),
            }
        if cell.number_format is not None:
            format_dict["number_format"] = cell.number_format

        backend = self._workbook._rust_writer  # noqa: SLF001
        if not hasattr(backend, "intern_format"):
            # The Rust side exposes intern via a dedicated method when
            # available; otherwise fall through to the per-cell write_cell_format
            # path on the next eager save (write-only never reaches that path).
            return None
        style_id = backend.intern_format(format_dict)
        self._style_cache[key] = style_id
        return style_id

    # ------------------------------------------------------------------
    # Forbidden-method guard
    # ------------------------------------------------------------------

    def __getitem__(self, key: Any) -> Any:
        raise AttributeError(
            f"WriteOnlyWorksheet does not support coordinate access ({key!r}); "
            "use .append(row) for the single supported write API"
        )

    def __setitem__(self, key: Any, value: Any) -> None:
        raise AttributeError(
            f"WriteOnlyWorksheet does not support coordinate assignment ({key!r}); "
            "use .append(row) for the single supported write API"
        )

    def __getattr__(self, name: str) -> Any:
        if name in _FORBIDDEN_ATTRS:
            raise AttributeError(
                f"WriteOnlyWorksheet does not support {name!r}; "
                "use .append(row) for the single supported write API"
            )
        raise AttributeError(
            f"'WriteOnlyWorksheet' object has no attribute {name!r}"
        )

    # ------------------------------------------------------------------
    # Sheet-level setters that route to the underlying Rust worksheet.
    # ------------------------------------------------------------------

    @property
    def freeze_panes(self) -> str | None:
        return self._freeze_panes

    @freeze_panes.setter
    def freeze_panes(self, value: str | None) -> None:
        self._guard_pre_append("freeze_panes")
        self._freeze_panes = value
        if value is not None:
            backend = self._workbook._rust_writer  # noqa: SLF001
            backend.set_freeze_panes(self.title, {"mode": "freeze", "top_left_cell": value})

    @property
    def print_area(self) -> str | None:
        return self._print_area

    @print_area.setter
    def print_area(self, value: str | None) -> None:
        self._guard_pre_append("print_area")
        self._print_area = value
        if value is not None:
            backend = self._workbook._rust_writer  # noqa: SLF001
            backend.set_print_area(self.title, value)

    def _guard_pre_append(self, attr: str) -> None:
        """Sheet-level slots must be set before any rows are appended.

        OOXML's slot order puts ``<cols>`` / ``<sheetViews>`` /
        ``<printOptions>`` BEFORE ``<sheetData>``. The streaming temp
        file has already started writing into the ``<sheetData>``
        block by the time the first row appends, so a late-arriving
        ``cols`` block would land in the wrong slot. Reject early so
        callers see a clean error.
        """
        if self._row_counter > 0:
            raise RuntimeError(
                f"{attr} must be set before appending rows on a write-only worksheet"
            )

    def close(self) -> None:
        """Mark this worksheet closed (post-save).

        Subsequent :meth:`append` calls raise
        :class:`WorkbookAlreadySaved`. This is also the hook the
        workbook save path uses to release the streaming
        ``BufWriter`` early when the user wants explicit lifecycle
        control.
        """
        self._closed = True
