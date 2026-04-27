"""Cell proxy — dispatches property access to the Rust backend."""

from __future__ import annotations

from datetime import date, datetime
from typing import TYPE_CHECKING, Any

from wolfxl._styles import Alignment, Border, Color, Font, PatternFill, Side
from wolfxl._utils import column_letter as _column_letter
from wolfxl._utils import rowcol_to_a1
from wolfxl.utils.numbers import is_date_format

if TYPE_CHECKING:
    from wolfxl._worksheet import Worksheet


class Cell:
    """Lightweight proxy for a single cell.

    In read mode, properties call into the Rust backend on first access.
    In write mode, assignments queue pending state flushed on ``save()``.
    """

    __slots__ = (
        "_ws",
        "_row",
        "_col",
        "_value",
        "_font",
        "_fill",
        "_border",
        "_alignment",
        "_number_format",
        "_value_dirty",
        "_format_dirty",
    )

    def __init__(self, ws: Worksheet, row: int, col: int) -> None:
        self._ws = ws
        self._row = row
        self._col = col
        # Sentinel — None is a valid value so we use a special marker.
        self._value: Any = _UNSET
        self._font: Font | None | _Sentinel = _UNSET
        self._fill: PatternFill | None | _Sentinel = _UNSET
        self._border: Border | None | _Sentinel = _UNSET
        self._alignment: Alignment | None | _Sentinel = _UNSET
        self._number_format: str | None | _Sentinel = _UNSET
        self._value_dirty = False
        self._format_dirty = False

    @property
    def coordinate(self) -> str:
        return rowcol_to_a1(self._row, self._col)

    @property
    def row(self) -> int:
        return self._row

    @property
    def column(self) -> int:
        return self._col

    @property
    def column_letter(self) -> str:
        """Column letter (e.g. ``"A"``, ``"AA"``) — openpyxl alias."""
        return _column_letter(self._col)

    @property
    def parent(self) -> Worksheet:
        """The containing Worksheet — openpyxl alias."""
        return self._ws

    def offset(self, row: int = 0, column: int = 0) -> Cell:
        """Return the cell ``row`` rows down and ``column`` columns right.

        Matches openpyxl's ``Cell.offset(row=0, column=0)`` signature. Negative
        offsets are allowed as long as the target row/col stays within Excel's
        1-based address space.
        """
        return self._ws._get_or_create_cell(self._row + row, self._col + column)  # noqa: SLF001

    @property
    def data_type(self) -> str:
        """openpyxl-compatible single-char type code.

        Maps to openpyxl's tags:
        - ``"s"``: string
        - ``"n"``: number (openpyxl also uses this for blank cells)
        - ``"b"``: boolean
        - ``"d"``: date / datetime
        - ``"f"``: formula
        - ``"e"``: error
        """
        from wolfxl._worksheet import _canonical_data_type

        canon = _canonical_data_type(self.value)
        mapping = {
            "string": "s",
            "number": "n",
            "boolean": "b",
            "datetime": "d",
            "date": "d",
            "formula": "f",
            "error": "e",
            "blank": "n",
        }
        return mapping.get(canon, "n")

    @property
    def has_style(self) -> bool:
        """True if any style attribute has been explicitly set on this cell.

        In read mode, checks whether the on-disk format carries any non-default
        style. In write mode, checks the dirty-flag sentinels so an unset cell
        reads as False and a cell with ``font = Font(bold=True)`` reads as True.
        """
        if self._format_dirty:
            return True
        font = self._font if self._font is not _UNSET else None
        fill = self._fill if self._fill is not _UNSET else None
        border = self._border if self._border is not _UNSET else None
        align = self._alignment if self._alignment is not _UNSET else None
        nfmt = self._number_format if self._number_format is not _UNSET else None
        if font and font != Font():
            return True
        if fill and fill != PatternFill():
            return True
        if border and border != Border():
            return True
        if align and align != Alignment():
            return True
        if nfmt and nfmt != "General":
            return True
        return False

    @property
    def is_date(self) -> bool:
        """True if the value is a date/datetime or the number format is a date."""
        value = self.value
        if hasattr(value, "year") and hasattr(value, "month"):
            return True
        # Sprint Κ Pod-β: xlsb / xls workbooks don't expose number_format
        # because the binary backends don't carry per-cell style records.
        # Fall back to the value-type check above and return False rather
        # than raise out of an introspection accessor.
        wb_format = getattr(self._ws._workbook, "_format", "xlsx")  # noqa: SLF001
        if wb_format != "xlsx":
            return False
        return is_date_format(self.number_format)

    @property
    def style(self) -> None:
        """Named style name. Not yet supported; always returns None.

        openpyxl uses this to look up ``NamedStyle`` objects in
        ``wb.named_styles``. wolfxl hasn't implemented named styles yet
        (tracked for T2). Setting it raises ``NotImplementedError``.
        """
        return None

    @style.setter
    def style(self, value: Any) -> None:  # noqa: ARG002
        raise NotImplementedError(
            "Named styles are not yet supported by wolfxl. "
            "See https://github.com/SynthGL/wolfxl#openpyxl-compatibility for tracking."
        )

    # ------------------------------------------------------------------
    # T1 PR1: hyperlink / comment read access (write-mode setters land in PR4)
    #
    # Reads pull from per-worksheet lazy maps populated on first access.
    # Cells without a hyperlink/comment return None (matches openpyxl).
    # Setters raise NotImplementedError with a T1.5 pointer when the file
    # was opened via load_workbook(...) (no rust writer); write-mode
    # implementations land in PR4.
    # ------------------------------------------------------------------

    @property
    def hyperlink(self) -> Any:
        # Pre-save visibility: a queued hyperlink shows up immediately
        # without waiting for ``save()`` to flush to the writer.
        ws = self._ws
        coord = self.coordinate
        pending = ws._pending_hyperlinks.get(coord, _UNSET)  # noqa: SLF001
        if pending is None:
            return None  # explicit-delete sentinel
        if pending is not _UNSET:
            return pending
        return ws._get_hyperlinks_map().get(coord)  # noqa: SLF001

    @hyperlink.setter
    def hyperlink(self, value: Any) -> None:
        from wolfxl.worksheet.hyperlink import Hyperlink

        ws = self._ws
        wb = ws._workbook  # noqa: SLF001
        # RFC-022: cell.hyperlink rounds-trips in BOTH write and modify
        # mode. Both backends consume the same _pending_hyperlinks dict;
        # the workbook flush dispatches to writer.add_hyperlink (write)
        # or patcher.queue_hyperlink (modify). None is the explicit-delete
        # sentinel per INDEX decision #5 — never use pop().
        if wb._rust_writer is None and wb._rust_patcher is None:  # noqa: SLF001
            raise RuntimeError("cell.hyperlink requires write or modify mode")
        coord = self.coordinate
        if value is None:
            ws._pending_hyperlinks[coord] = None  # noqa: SLF001
            return
        if isinstance(value, str):
            value = Hyperlink(target=value)
        if not isinstance(value, Hyperlink):
            raise TypeError(
                f"hyperlink must be a Hyperlink or str, got {type(value).__name__}"
            )
        ws._pending_hyperlinks[coord] = value  # noqa: SLF001
        # openpyxl parity: if the cell has no value yet, surface the
        # target/location as the visible cell value so a freshly-set
        # hyperlink is also visibly clickable text.
        if self.value is None:
            display_value = value.display or value.target or value.location
            if display_value is not None:
                self.value = display_value

    @property
    def comment(self) -> Any:
        ws = self._ws
        coord = self.coordinate
        pending = ws._pending_comments.get(coord, _UNSET)  # noqa: SLF001
        if pending is None:
            return None
        if pending is not _UNSET:
            return pending
        return ws._get_comments_map().get(coord)  # noqa: SLF001

    @comment.setter
    def comment(self, value: Any) -> None:
        from wolfxl.comments import Comment

        ws = self._ws
        wb = ws._workbook  # noqa: SLF001
        # RFC-023: cell.comment round-trips in both write and modify
        # mode. Both backends consume the same _pending_comments dict;
        # the workbook flush dispatches to writer.add_comment (write)
        # or patcher.queue_comment (modify). None is the explicit-
        # delete sentinel.
        if wb._rust_writer is None and wb._rust_patcher is None:  # noqa: SLF001
            raise RuntimeError("cell.comment requires write or modify mode")
        coord = self.coordinate
        if value is None:
            ws._pending_comments[coord] = None  # noqa: SLF001
            return
        if not isinstance(value, Comment):
            raise TypeError(
                f"comment must be a Comment, got {type(value).__name__}"
            )
        ws._pending_comments[coord] = value  # noqa: SLF001

    @property
    def protection(self) -> None:
        """Read-only default (None). Cell protection is not supported."""
        return None

    # ------------------------------------------------------------------
    # Value
    # ------------------------------------------------------------------

    @property
    def value(self) -> Any:
        if self._value is _UNSET:
            self._value = self._read_value()
        # Sprint Ι Pod-α: when the workbook was opened with
        # ``rich_text=True``, surface ``CellRichText`` for cells whose
        # backing string carries `<r>` runs.  Default load mode mirrors
        # openpyxl 3.x, which flattens to plain ``str`` unless the user
        # opts in via the same flag.
        if isinstance(self._value, str):
            wb = self._ws._workbook  # noqa: SLF001
            if getattr(wb, "_rich_text", False):
                rt = self.rich_text
                if rt is not None:
                    return rt
        return self._value

    @value.setter
    def value(self, val: Any) -> None:
        # Accept CellRichText pass-through: if the user assigns a
        # CellRichText, defer rich-text serialization to the writer.
        # Plain strings keep the existing fast path.
        self._value = val
        self._value_dirty = True
        self._ws._mark_dirty(self._row, self._col)  # noqa: SLF001
        # Pod-α: when a CellRichText is assigned, also stash it on the
        # worksheet's pending-rich-text map so the flush layer can pick
        # it up (write-mode and modify-mode both consume the same map).
        from wolfxl.cell.rich_text import CellRichText  # local import — avoids cycles

        ws = self._ws
        if isinstance(val, CellRichText):
            ws._pending_rich_text[(self._row, self._col)] = val  # noqa: SLF001
        else:
            # Clearing or replacing with plain — drop any prior rich entry.
            ws._pending_rich_text.pop((self._row, self._col), None)  # noqa: SLF001

    # ------------------------------------------------------------------
    # Rich text
    # ------------------------------------------------------------------

    @property
    def rich_text(self) -> Any:
        """Structured rich-text runs for this cell, or ``None``.

        Returns a :class:`wolfxl.cell.rich_text.CellRichText` object
        when the on-disk cell carries `<r>` runs (either via the SST
        or as inline-string runs).  Returns ``None`` for plain-text
        cells, non-string types, and brand-new cells with no on-disk
        backing.

        Parity with openpyxl: openpyxl exposes the same data via
        ``Cell.value`` *only* when the workbook is loaded with
        ``rich_text=True``.  WolfXL goes one step further and always
        surfaces the structured representation through this side
        channel — defaulting ``Cell.value`` to flattened ``str`` so
        existing user code keeps working unchanged.
        """
        from wolfxl.cell.rich_text import CellRichText, InlineFont, TextBlock

        ws = self._ws
        # Pre-save visibility for write/modify-mode setters.
        pending = ws._pending_rich_text.get((self._row, self._col))  # noqa: SLF001
        if pending is not None:
            return pending

        wb = ws._workbook  # noqa: SLF001
        reader = getattr(wb, "_rust_reader", None)
        if reader is None:
            return None
        payload = reader.read_cell_rich_text(ws.title, self.coordinate)
        if payload is None:
            return None
        out = CellRichText()
        for entry in payload:
            text, font_dict = entry[0], entry[1]
            if font_dict is None:
                out.append(text)
            else:
                out.append(
                    TextBlock(
                        InlineFont(
                            rFont=font_dict.get("rFont"),
                            charset=font_dict.get("charset"),
                            family=font_dict.get("family"),
                            b=font_dict.get("b"),
                            i=font_dict.get("i"),
                            strike=font_dict.get("strike"),
                            color=font_dict.get("color"),
                            sz=font_dict.get("sz"),
                            u=font_dict.get("u"),
                            vertAlign=font_dict.get("vertAlign"),
                            scheme=font_dict.get("scheme"),
                        ),
                        text,
                    )
                )
        return out

    @rich_text.setter
    def rich_text(self, val: Any) -> None:
        """Setter alias for ``cell.value = CellRichText(...)``.

        Lets users round-trip via ``cell.rich_text = ...`` even if they
        never touch ``cell.value`` directly — handy in code that wants
        to add/edit runs without disturbing other state.
        """
        self.value = val

    # ------------------------------------------------------------------
    # Style guard (Sprint Κ Pod-β)
    # ------------------------------------------------------------------

    def _require_xlsx_for_style(self, attr: str) -> None:
        """Raise NotImplementedError if the workbook isn't xlsx.

        xlsb / xls workbooks load via calamine's binary readers which
        don't expose the per-cell style records that ``cell.font``,
        ``cell.fill`` etc. need.  Surface a clear error at the Python
        layer so callers don't get a confusing Rust panic / empty
        Font object.
        """
        wb_format = getattr(self._ws._workbook, "_format", "xlsx")  # noqa: SLF001
        if wb_format != "xlsx":
            raise NotImplementedError(
                f"cell.{attr} is xlsx-only; this workbook is .{wb_format}. "
                "Use .xlsx for style-aware reads."
            )

    # ------------------------------------------------------------------
    # Font
    # ------------------------------------------------------------------

    @property
    def font(self) -> Font:
        self._require_xlsx_for_style("font")
        if self._font is _UNSET:
            self._font = self._read_font()
        return self._font  # type: ignore[return-value]

    @font.setter
    def font(self, val: Font) -> None:
        self._require_xlsx_for_style("font")
        self._font = val
        self._format_dirty = True
        self._ws._mark_dirty(self._row, self._col)  # noqa: SLF001

    # ------------------------------------------------------------------
    # Fill
    # ------------------------------------------------------------------

    @property
    def fill(self) -> PatternFill:
        self._require_xlsx_for_style("fill")
        if self._fill is _UNSET:
            self._fill = self._read_fill()
        return self._fill  # type: ignore[return-value]

    @fill.setter
    def fill(self, val: PatternFill) -> None:
        self._require_xlsx_for_style("fill")
        self._fill = val
        self._format_dirty = True
        self._ws._mark_dirty(self._row, self._col)  # noqa: SLF001

    # ------------------------------------------------------------------
    # Border
    # ------------------------------------------------------------------

    @property
    def border(self) -> Border:
        self._require_xlsx_for_style("border")
        if self._border is _UNSET:
            self._border = self._read_border()
        return self._border  # type: ignore[return-value]

    @border.setter
    def border(self, val: Border) -> None:
        self._require_xlsx_for_style("border")
        self._border = val
        self._format_dirty = True
        self._ws._mark_dirty(self._row, self._col)  # noqa: SLF001

    # ------------------------------------------------------------------
    # Alignment
    # ------------------------------------------------------------------

    @property
    def alignment(self) -> Alignment:
        self._require_xlsx_for_style("alignment")
        if self._alignment is _UNSET:
            self._alignment = self._read_alignment()
        return self._alignment  # type: ignore[return-value]

    @alignment.setter
    def alignment(self, val: Alignment) -> None:
        self._require_xlsx_for_style("alignment")
        self._alignment = val
        self._format_dirty = True
        self._ws._mark_dirty(self._row, self._col)  # noqa: SLF001

    # ------------------------------------------------------------------
    # Number format
    # ------------------------------------------------------------------

    @property
    def number_format(self) -> str | None:
        self._require_xlsx_for_style("number_format")
        if self._number_format is _UNSET:
            self._number_format = self._read_number_format()
        return self._number_format  # type: ignore[return-value]

    @number_format.setter
    def number_format(self, val: str | None) -> None:
        self._require_xlsx_for_style("number_format")
        self._number_format = val
        self._format_dirty = True
        self._ws._mark_dirty(self._row, self._col)  # noqa: SLF001

    # ------------------------------------------------------------------
    # Read helpers (dispatch to Rust via workbook)
    # ------------------------------------------------------------------

    def _read_value(self) -> Any:
        wb = self._ws._workbook  # noqa: SLF001
        if wb._rust_reader is None:  # noqa: SLF001
            return None
        payload = wb._rust_reader.read_cell_value(  # noqa: SLF001
            self._ws.title, self.coordinate, getattr(wb, "_data_only", False),
        )
        return _payload_to_python(payload)

    def _read_font(self) -> Font:
        wb = self._ws._workbook  # noqa: SLF001
        if wb._rust_reader is None:  # noqa: SLF001
            return Font()
        payload = wb._rust_reader.read_cell_format(  # noqa: SLF001
            self._ws.title, self.coordinate,
        )
        return _format_to_font(payload)

    def _read_fill(self) -> PatternFill:
        wb = self._ws._workbook  # noqa: SLF001
        if wb._rust_reader is None:  # noqa: SLF001
            return PatternFill()
        payload = wb._rust_reader.read_cell_format(  # noqa: SLF001
            self._ws.title, self.coordinate,
        )
        return _format_to_fill(payload)

    def _read_border(self) -> Border:
        wb = self._ws._workbook  # noqa: SLF001
        if wb._rust_reader is None:  # noqa: SLF001
            return Border()
        payload = wb._rust_reader.read_cell_border(  # noqa: SLF001
            self._ws.title, self.coordinate,
        )
        return _border_payload_to_border(payload)

    def _read_alignment(self) -> Alignment:
        wb = self._ws._workbook  # noqa: SLF001
        if wb._rust_reader is None:  # noqa: SLF001
            return Alignment()
        payload = wb._rust_reader.read_cell_format(  # noqa: SLF001
            self._ws.title, self.coordinate,
        )
        return _format_to_alignment(payload)

    def _read_number_format(self) -> str | None:
        wb = self._ws._workbook  # noqa: SLF001
        if wb._rust_reader is None:  # noqa: SLF001
            return None
        payload = wb._rust_reader.read_cell_format(  # noqa: SLF001
            self._ws.title, self.coordinate,
        )
        if isinstance(payload, dict):
            return payload.get("number_format")
        return None

    def __repr__(self) -> str:
        return f"<Cell {self.coordinate}>"


# ======================================================================
# Sentinel type for lazy-init detection
# ======================================================================

class _Sentinel:
    """Marker to distinguish 'not yet loaded' from None."""

    _instance: _Sentinel | None = None

    def __new__(cls) -> _Sentinel:
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __repr__(self) -> str:
        return "<UNSET>"

    def __bool__(self) -> bool:
        return False


_UNSET = _Sentinel()


# ======================================================================
# Payload <-> Python conversion helpers
# ======================================================================

def _payload_to_python(payload: Any) -> Any:
    """Convert a Rust cell-value payload dict to a plain Python value."""
    if not isinstance(payload, dict):
        return payload
    t = payload.get("type", "blank")
    v = payload.get("value")
    if t == "blank":
        return None
    if t == "string":
        return v
    if t == "number":
        return v
    if t == "boolean":
        return bool(v)
    if t == "error":
        return v
    if t == "formula":
        return payload.get("formula", v)
    if t == "date":
        if isinstance(v, str):
            return datetime.fromisoformat(v)
        if isinstance(v, date) and not isinstance(v, datetime):
            return datetime.combine(v, datetime.min.time())
        return v
    if t == "datetime":
        if isinstance(v, str):
            return datetime.fromisoformat(v)
        return v
    return v


def _format_to_font(payload: Any) -> Font:
    """Extract Font fields from a Rust format dict."""
    if not isinstance(payload, dict) or not payload:
        return Font()
    color_raw = payload.get("font_color")
    color: Color | str | None = None
    if color_raw:
        color = color_raw if isinstance(color_raw, str) else str(color_raw)
    return Font(
        name=payload.get("font_name"),
        size=payload.get("font_size"),
        bold=bool(payload.get("bold", False)),
        italic=bool(payload.get("italic", False)),
        underline=payload.get("underline"),
        strike=bool(payload.get("strikethrough", False)),
        color=color,
    )


def _format_to_fill(payload: Any) -> PatternFill:
    """Extract PatternFill fields from a Rust format dict."""
    if not isinstance(payload, dict) or not payload:
        return PatternFill()
    bg = payload.get("bg_color")
    if bg:
        return PatternFill(patternType="solid", fgColor=bg)
    return PatternFill()


def _format_to_alignment(payload: Any) -> Alignment:
    """Extract Alignment fields from a Rust format dict."""
    if not isinstance(payload, dict) or not payload:
        return Alignment()
    return Alignment(
        horizontal=payload.get("h_align"),
        vertical=payload.get("v_align"),
        wrap_text=bool(payload.get("wrap", False)),
        text_rotation=int(payload.get("rotation", 0)),
        indent=int(payload.get("indent", 0)),
    )


def _border_payload_to_border(payload: Any) -> Border:
    """Convert a Rust border dict to a Border dataclass."""
    if not isinstance(payload, dict) or not payload:
        return Border()
    return Border(
        left=_edge_to_side(payload.get("left")),
        right=_edge_to_side(payload.get("right")),
        top=_edge_to_side(payload.get("top")),
        bottom=_edge_to_side(payload.get("bottom")),
    )


def _edge_to_side(edge: Any) -> Side:
    if not isinstance(edge, dict):
        return Side()
    return Side(
        style=edge.get("style"),
        color=edge.get("color"),
    )


# ======================================================================
# Python -> Rust payload converters (for write mode)
# ======================================================================

def python_value_to_payload(value: Any) -> dict[str, Any]:
    """Convert a plain Python value to a Rust cell-value payload dict."""
    if value is None:
        return {"type": "blank"}
    if isinstance(value, bool):
        return {"type": "boolean", "value": value}
    if isinstance(value, (int, float)):
        return {"type": "number", "value": value}
    if isinstance(value, datetime):
        return {"type": "datetime", "value": value.replace(microsecond=0).isoformat()}
    if isinstance(value, date):
        return {"type": "date", "value": value.isoformat()}
    if isinstance(value, str) and value.startswith("="):
        return {"type": "formula", "formula": value, "value": value}
    return {"type": "string", "value": str(value)}


def font_to_format_dict(font: Font) -> dict[str, Any]:
    """Convert a Font to a Rust format dict."""
    d: dict[str, Any] = {}
    if font.bold:
        d["bold"] = True
    if font.italic:
        d["italic"] = True
    if font.underline:
        d["underline"] = font.underline
    if font.strike:
        d["strikethrough"] = True
    if font.name:
        d["font_name"] = font.name
    if font.size is not None:
        d["font_size"] = font.size
    color_hex = font._color_hex()  # noqa: SLF001
    if color_hex:
        d["font_color"] = color_hex
    return d


def fill_to_format_dict(fill: PatternFill) -> dict[str, Any]:
    """Convert a PatternFill to a Rust format dict."""
    d: dict[str, Any] = {}
    fg = fill._fg_hex()  # noqa: SLF001
    if fg:
        d["bg_color"] = fg
    return d


def alignment_to_format_dict(alignment: Alignment) -> dict[str, Any]:
    """Convert an Alignment to a Rust format dict."""
    d: dict[str, Any] = {}
    if alignment.horizontal:
        d["h_align"] = alignment.horizontal
    if alignment.vertical:
        d["v_align"] = alignment.vertical
    if alignment.wrap_text:
        d["wrap"] = True
    if alignment.text_rotation:
        d["rotation"] = alignment.text_rotation
    if alignment.indent:
        d["indent"] = alignment.indent
    return d


def border_to_rust_dict(border: Border) -> dict[str, Any]:
    """Convert a Border to a Rust border dict."""
    d: dict[str, Any] = {}
    for edge_name in ("left", "right", "top", "bottom"):
        side: Side = getattr(border, edge_name)
        if side.style:
            edge: dict[str, str] = {"style": side.style}
            color = side._color_hex()  # noqa: SLF001
            if color:
                edge["color"] = color
            else:
                edge["color"] = "#000000"
            d[edge_name] = edge
    return d
