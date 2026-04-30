"""Cell proxy — dispatches property access to the Rust backend."""

from __future__ import annotations

import re
from typing import TYPE_CHECKING, Any

from wolfxl._cell_annotations import (
    get_comment,
    get_hyperlink,
    set_comment,
    set_hyperlink,
)
from wolfxl._cell_payloads import (
    alignment_to_format_dict as alignment_to_format_dict,
    border_payload_to_border as _border_payload_to_border,
    border_to_rust_dict as border_to_rust_dict,
    fill_to_format_dict as fill_to_format_dict,
    font_to_format_dict as font_to_format_dict,
    format_to_alignment as _format_to_alignment,
    format_to_fill as _format_to_fill,
    format_to_font as _format_to_font,
    payload_to_python as _payload_to_python,
    python_value_to_payload as python_value_to_payload,
)
from wolfxl._styles import Alignment, Border, Font, PatternFill
from wolfxl._utils import column_letter as _column_letter
from wolfxl._utils import rowcol_to_a1
from wolfxl._worksheet_rich_text import runs_payload_to_cellrichtext
from wolfxl.utils.exceptions import IllegalCharacterError
from wolfxl.utils.numbers import is_date_format

if TYPE_CHECKING:
    from wolfxl._worksheet import Worksheet


# RFC-059 (Sprint Ο Pod-1E): OOXML-illegal control characters.
# The C0 controls 0x00–0x08, 0x0B, 0x0C, 0x0E–0x1F plus 0x7F are
# rejected by Excel's serializer.  Tab (0x09), newline (0x0A), and
# carriage return (0x0D) are allowed and pass through unchanged.
# Mirrors openpyxl's ``ILLEGAL_CHARACTERS_RE``.
ILLEGAL_CHARACTERS_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")


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
        # Array / data-table formula metadata. Populated when ``cell.value`` is
        # assigned an
        # :class:`ArrayFormula` / :class:`DataTableFormula` instance,
        # or when an existing cell parses back as one of those types.
        # ``_formula_type`` is one of: ``None`` (plain), ``"array"``,
        # ``"dataTable"``.
        "_formula_type",
        "_array_ref",
        "_formula_text",
        "_dt_ca",
        "_dt_2d",
        "_dt_r",
        "_dt_r1",
        "_dt_r2",
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
        # None until the cell is identified as array / data-table either via
        # setter or on read-back.
        self._formula_type: str | None = None
        self._array_ref: str | None = None
        self._formula_text: str | None = None
        self._dt_ca: bool = False
        self._dt_2d: bool = False
        self._dt_r: bool = False
        self._dt_r1: str | None = None
        self._dt_r2: str | None = None

    @property
    def coordinate(self) -> str:
        """Return this cell's A1-style coordinate."""
        return rowcol_to_a1(self._row, self._col)

    @property
    def row(self) -> int:
        """Return this cell's 1-based row index."""
        return self._row

    @property
    def column(self) -> int:
        """Return this cell's 1-based column index."""
        return self._col

    @property
    def col_idx(self) -> int:
        """Return this cell's 1-based column index."""
        return self._col

    @property
    def column_letter(self) -> str:
        """Column letter (e.g. ``"A"``, ``"AA"``) — openpyxl alias."""
        return _column_letter(self._col)

    @property
    def parent(self) -> Worksheet:
        """The containing Worksheet — openpyxl alias."""
        return self._ws

    @property
    def base_date(self) -> Any:
        """Workbook epoch used for Excel serial date conversion."""
        return self._ws._workbook.excel_base_date  # noqa: SLF001

    @property
    def encoding(self) -> str:
        """Cell text encoding marker for openpyxl compatibility."""
        return "utf-8"

    @property
    def internal_value(self) -> Any:
        """Internal cell value; WolfXL stores the openpyxl-facing value."""
        return self.value

    @property
    def pivotButton(self) -> bool:  # noqa: N802 - openpyxl public alias
        """Whether the cell displays a pivot-table field button."""
        return False

    @property
    def quotePrefix(self) -> bool:  # noqa: N802 - openpyxl public alias
        """Whether the cell has Excel's quote-prefix flag."""
        return False

    @property
    def style_id(self) -> int:
        """Return the workbook style identifier for this cell."""
        if self._format_dirty:
            return 1
        return self._read_style_id()

    def offset(self, row: int = 0, column: int = 0) -> Cell:
        """Return the cell ``row`` rows down and ``column`` columns right.

        Matches openpyxl's ``Cell.offset(row=0, column=0)`` signature. Negative
        offsets are allowed as long as the target row/col stays within Excel's
        1-based address space.
        """
        return self._ws._get_or_create_cell(self._row + row, self._col + column)  # noqa: SLF001

    def check_string(self, value: Any) -> str | None:
        """Validate a worksheet string using openpyxl's public helper rules."""
        if value is None:
            return None
        if not isinstance(value, str):
            value = str(value, self.encoding)
        value = str(value)[:32767]
        if ILLEGAL_CHARACTERS_RE.search(value):
            raise IllegalCharacterError(f"{value} cannot be used in worksheets.")
        return value

    def check_error(self, value: Any) -> str:
        """Convert an error-like value to openpyxl's string representation."""
        try:
            return str(value)
        except UnicodeDecodeError:
            return "#N/A"

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
        # Binary formats may not expose style metadata. Fall back to the
        # value-type check above rather than raise from an introspection
        # accessor.
        wb_format = getattr(self._ws._workbook, "_format", "xlsx")  # noqa: SLF001
        if wb_format != "xlsx":
            return False
        return is_date_format(self.number_format)

    @property
    def style(self) -> None:
        """Return the named style assigned to this cell, if any.

        WolfXL currently preserves many style attributes through the
        explicit ``font``, ``fill``, ``border``, ``alignment``, and
        ``number_format`` accessors. Named-style objects are not exposed
        through this compatibility property, so the getter returns ``None``.
        """
        return None

    @style.setter
    def style(self, value: Any) -> None:  # noqa: ARG002
        """Reject named-style assignment.

        Args:
            value: Named style requested by the caller.
        """
        raise NotImplementedError(
            "Named styles are not yet supported by wolfxl. "
            "See https://github.com/SynthGL/wolfxl#openpyxl-compatibility "
            "for compatibility notes."
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
        """Return the cell hyperlink, including pending unsaved edits."""
        return get_hyperlink(self, _UNSET)

    @hyperlink.setter
    def hyperlink(self, value: Any) -> None:
        """Set or clear the cell hyperlink.

        Args:
            value: ``Hyperlink`` instance, URL string, or ``None`` to delete
                the hyperlink on the next save.
        """
        set_hyperlink(self, value)

    @property
    def comment(self) -> Any:
        """Return the cell comment, including pending unsaved edits."""
        return get_comment(self, _UNSET)

    @comment.setter
    def comment(self, value: Any) -> None:
        """Set or clear the cell comment.

        Args:
            value: ``Comment`` instance, or ``None`` to delete the comment on
                the next save.
        """
        set_comment(self, value)

    @property
    def protection(self) -> None:
        """Read-only default (None). Cell protection is not supported."""
        return None

    # ------------------------------------------------------------------
    # Value
    # ------------------------------------------------------------------

    @property
    def value(self) -> Any:
        """Return the cell value using openpyxl-compatible Python types."""
        # RFC-057: surface array / data-table formulas as the typed
        # instance regardless of what's been cached in ``_value``.
        # The metadata is populated either by the setter or by the
        # read-back path (parse_cell in the calamine backend tags the
        # cell post-read; pending-array map carries write-side state).
        pending_value = self._value_from_pending_formula()
        if pending_value is not _UNSET:
            return pending_value

        formula_value = self._value_from_formula_metadata()
        if formula_value is not _UNSET:
            return formula_value

        if self._value is _UNSET:
            self._value = self._read_value()
            # _read_value may have populated the formula metadata —
            # re-check after the read.
            formula_value = self._value_from_formula_metadata()
            if formula_value is not _UNSET:
                return formula_value
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
        """Set the cell value and queue it for the next workbook save.

        Args:
            val: Scalar value, formula string, rich text object, array formula,
                data-table formula, or ``None``.
        """
        # Accept CellRichText pass-through: if the user assigns a
        # CellRichText, defer rich-text serialization to the writer.
        # Plain strings keep the existing fast path.
        # RFC-059: reject OOXML-illegal control characters before
        # they hit the writer.  ``IllegalCharacterError`` subclasses
        # ``ValueError`` so existing ``except ValueError`` callsites
        # keep working unchanged.
        if isinstance(val, str) and ILLEGAL_CHARACTERS_RE.search(val):
            raise IllegalCharacterError(
                f"Cell value {val!r} contains characters that are not allowed in "
                "OOXML strings (control chars 0x00-0x08, 0x0B, 0x0C, 0x0E-0x1F, 0x7F)"
            )
        from wolfxl.cell.cell import ArrayFormula, DataTableFormula
        from wolfxl.cell.rich_text import CellRichText  # local import — avoids cycles

        ws = self._ws

        # RFC-057 — array / data-table formula assignment.
        if isinstance(val, ArrayFormula):
            self._queue_array_formula(val)
            return

        if isinstance(val, DataTableFormula):
            self._queue_data_table_formula(val)
            return

        # Plain assignment — clear any previous array / data-table
        # state so a former master cell can be replaced cleanly.
        self._clear_formula_metadata()

        self._value = val
        self._value_dirty = True
        ws._mark_dirty(self._row, self._col)  # noqa: SLF001

        # Pod-α: when a CellRichText is assigned, also stash it on the
        # worksheet's pending-rich-text map so the flush layer can pick
        # it up (write-mode and modify-mode both consume the same map).
        if isinstance(val, CellRichText):
            ws._pending_rich_text[(self._row, self._col)] = val  # noqa: SLF001
        else:
            # Clearing or replacing with plain — drop any prior rich entry.
            ws._pending_rich_text.pop((self._row, self._col), None)  # noqa: SLF001

    def _value_from_pending_formula(self) -> Any:
        """Return pending array/data-table formula value or ``_UNSET``."""
        from wolfxl.cell.cell import ArrayFormula, DataTableFormula

        pending = self._ws._pending_array_formulas.get((self._row, self._col))  # noqa: SLF001
        if pending is None:
            return _UNSET
        kind, payload = pending
        if kind == "spill_child":
            return None
        if kind == "array":
            return ArrayFormula(payload["ref"], payload["text"])
        if kind == "data_table":
            return DataTableFormula(
                ref=payload["ref"],
                ca=payload.get("ca", False),
                dt2D=payload.get("dt2D", False),
                dtr=payload.get("dtr", False),
                r1=payload.get("r1"),
                r2=payload.get("r2"),
            )
        return _UNSET

    def _value_from_formula_metadata(self) -> Any:
        """Return read-side formula metadata value or ``_UNSET``."""
        from wolfxl.cell.cell import ArrayFormula, DataTableFormula

        if self._formula_type == "array":
            return ArrayFormula(self._array_ref or "", self._formula_text or "")
        if self._formula_type == "dataTable":
            return DataTableFormula(
                ref=self._array_ref or "",
                ca=self._dt_ca,
                dt2D=self._dt_2d,
                dtr=self._dt_r,
                r1=self._dt_r1,
                r2=self._dt_r2,
            )
        if self._formula_type == "array_child":
            return None
        return _UNSET

    def _clear_formula_metadata(self) -> None:
        """Clear array/data-table metadata and pending formula state."""
        self._formula_type = None
        self._array_ref = None
        self._formula_text = None
        self._dt_ca = False
        self._dt_2d = False
        self._dt_r = False
        self._dt_r1 = None
        self._dt_r2 = None
        self._ws._pending_array_formulas.pop((self._row, self._col), None)  # noqa: SLF001

    def _queue_array_formula(self, val: Any) -> None:
        """Queue an array formula assignment for save."""
        ws = self._ws
        self._formula_type = "array"
        self._array_ref = val.ref
        self._formula_text = val.text
        self._value = val
        self._value_dirty = True
        ws._mark_dirty(self._row, self._col)  # noqa: SLF001
        ws._pending_array_formulas[(self._row, self._col)] = (  # noqa: SLF001
            "array",
            {"ref": val.ref, "text": val.text},
        )
        # Populate placeholder entries for cells inside the spill range
        # (excluding the master). These show up as ``<c r="..."/>``
        # placeholders so Excel sees the spill area pre-populated.
        self._populate_spill_placeholders(val.ref)
        ws._pending_rich_text.pop((self._row, self._col), None)  # noqa: SLF001

    def _queue_data_table_formula(self, val: Any) -> None:
        """Queue a data-table formula assignment for save."""
        ws = self._ws
        self._formula_type = "dataTable"
        self._array_ref = val.ref
        self._dt_ca = val.ca
        self._dt_2d = val.dt2D
        self._dt_r = val.dtr
        self._dt_r1 = val.r1
        self._dt_r2 = val.r2
        self._value = val
        self._value_dirty = True
        ws._mark_dirty(self._row, self._col)  # noqa: SLF001
        ws._pending_array_formulas[(self._row, self._col)] = (  # noqa: SLF001
            "data_table",
            {
                "ref": val.ref,
                "ca": val.ca,
                "dt2D": val.dt2D,
                "dtr": val.dtr,
                "r1": val.r1,
                "r2": val.r2,
            },
        )
        ws._pending_rich_text.pop((self._row, self._col), None)  # noqa: SLF001

    def _populate_spill_placeholders(self, ref: str) -> None:
        """Mark every non-master cell in ``ref`` as a spill child.

        RFC-057: when the user assigns ``cell.value = ArrayFormula(...)``,
        every cell inside the spill range becomes a placeholder so the
        ``cell.value`` getter on those cells returns ``None`` (mirroring
        openpyxl/Excel).  Only the master cell carries the actual
        formula text.
        """
        from wolfxl._utils import a1_to_rowcol  # noqa: SLF001

        ws = self._ws
        # Parse the ref ("A1:A10") into a 2-tuple of cells.
        if ":" not in ref:
            return  # single-cell array — nothing else to mark
        try:
            top_left, bottom_right = ref.split(":", 1)
            r1, c1 = a1_to_rowcol(top_left)
            r2, c2 = a1_to_rowcol(bottom_right)
        except Exception:  # noqa: BLE001
            return
        top, bottom = sorted((r1, r2))
        left, right = sorted((c1, c2))
        master_key = (self._row, self._col)
        for r in range(top, bottom + 1):
            for c in range(left, right + 1):
                if (r, c) == master_key:
                    continue
                ws._pending_array_formulas[(r, c)] = ("spill_child", {})  # noqa: SLF001

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
        return runs_payload_to_cellrichtext(payload)

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
        """Raise NotImplementedError if this format cannot expose styles."""
        wb_format = getattr(self._ws._workbook, "_format", "xlsx")  # noqa: SLF001
        if wb_format == "xlsb" and attr in {
            "font",
            "fill",
            "border",
            "alignment",
            "number_format",
        }:
            return
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
        """Return the resolved cell font."""
        self._require_xlsx_for_style("font")
        if self._font is _UNSET:
            self._font = self._read_font()
        return self._font  # type: ignore[return-value]

    @font.setter
    def font(self, val: Font) -> None:
        """Set the cell font.

        Args:
            val: Font object to apply to the cell.
        """
        self._set_style_value("font", "_font", val)

    # ------------------------------------------------------------------
    # Fill
    # ------------------------------------------------------------------

    @property
    def fill(self) -> PatternFill:
        """Return the resolved cell fill."""
        self._require_xlsx_for_style("fill")
        if self._fill is _UNSET:
            self._fill = self._read_fill()
        return self._fill  # type: ignore[return-value]

    @fill.setter
    def fill(self, val: PatternFill) -> None:
        """Set the cell fill.

        Args:
            val: Pattern fill object to apply to the cell.
        """
        self._set_style_value("fill", "_fill", val)

    # ------------------------------------------------------------------
    # Border
    # ------------------------------------------------------------------

    @property
    def border(self) -> Border:
        """Return the resolved cell border."""
        self._require_xlsx_for_style("border")
        if self._border is _UNSET:
            self._border = self._read_border()
        return self._border  # type: ignore[return-value]

    @border.setter
    def border(self, val: Border) -> None:
        """Set the cell border.

        Args:
            val: Border object to apply to the cell.
        """
        self._set_style_value("border", "_border", val)

    # ------------------------------------------------------------------
    # Alignment
    # ------------------------------------------------------------------

    @property
    def alignment(self) -> Alignment:
        """Return the resolved cell alignment."""
        self._require_xlsx_for_style("alignment")
        if self._alignment is _UNSET:
            self._alignment = self._read_alignment()
        return self._alignment  # type: ignore[return-value]

    @alignment.setter
    def alignment(self, val: Alignment) -> None:
        """Set the cell alignment.

        Args:
            val: Alignment object to apply to the cell.
        """
        self._set_style_value("alignment", "_alignment", val)

    # ------------------------------------------------------------------
    # Number format
    # ------------------------------------------------------------------

    @property
    def number_format(self) -> str | None:
        """Return the resolved number format string."""
        self._require_xlsx_for_style("number_format")
        if self._number_format is _UNSET:
            self._number_format = self._read_number_format()
        return self._number_format  # type: ignore[return-value]

    @number_format.setter
    def number_format(self, val: str | None) -> None:
        """Set the cell number format.

        Args:
            val: Number format code, or ``None`` to clear the cached format.
        """
        self._set_style_value("number_format", "_number_format", val)

    def _set_style_value(self, public_attr: str, storage_attr: str, value: Any) -> None:
        """Set a cached style value and mark the cell dirty."""
        if getattr(self._ws._workbook, "_format", "xlsx") == "xlsb":  # noqa: SLF001
            raise NotImplementedError(
                f"cell.{public_attr} assignment is xlsx-only; .xlsb workbooks "
                "are read-only in WolfXL. Transcribe to .xlsx before editing styles."
            )
        self._require_xlsx_for_style(public_attr)
        setattr(self, storage_attr, value)
        self._format_dirty = True
        self._ws._mark_dirty(self._row, self._col)  # noqa: SLF001

    # ------------------------------------------------------------------
    # Read helpers (dispatch to Rust via workbook)
    # ------------------------------------------------------------------

    def _read_value(self) -> Any:
        wb = self._ws._workbook  # noqa: SLF001
        if wb._rust_reader is None:  # noqa: SLF001
            return None
        # RFC-057: tag the cell with array / data-table metadata
        # before falling through to the regular payload read.  The
        # reader returns ``None`` for plain cells so the cost is one
        # extra dict-lookup at most.
        try:
            af_payload = wb._rust_reader.read_cell_array_formula(  # noqa: SLF001
                self._ws.title, self.coordinate,
            )
        except AttributeError:
            af_payload = None
        if af_payload is not None:
            kind = af_payload.get("kind")
            if kind == "array":
                self._formula_type = "array"
                self._array_ref = af_payload.get("ref")
                self._formula_text = af_payload.get("text", "")
            elif kind == "data_table":
                self._formula_type = "dataTable"
                self._array_ref = af_payload.get("ref")
                self._dt_ca = bool(af_payload.get("ca", False))
                self._dt_2d = bool(af_payload.get("dt2D", False))
                self._dt_r = bool(af_payload.get("dtr", False))
                self._dt_r1 = af_payload.get("r1")
                self._dt_r2 = af_payload.get("r2")
            elif kind == "spill_child":
                self._formula_type = "array_child"
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

    def _read_style_id(self) -> int:
        wb = self._ws._workbook  # noqa: SLF001
        if wb._rust_reader is None:  # noqa: SLF001
            return 0
        try:
            records = wb._rust_reader.read_sheet_records(  # noqa: SLF001
                self._ws.title,
                f"{self.coordinate}:{self.coordinate}",
                getattr(wb, "_data_only", False),
                True,
                True,
                True,
                False,
                True,
                False,
                False,
            )
        except AttributeError:
            return 0
        for record in records:
            style_id = record.get("style_id")
            return int(style_id or 0)
        return 0

    def __repr__(self) -> str:
        """Return a compact debug representation for this cell."""
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
        """Return the sentinel's debug label."""
        return "<UNSET>"

    def __bool__(self) -> bool:
        return False


_UNSET = _Sentinel()
