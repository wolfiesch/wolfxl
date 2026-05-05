""":class:`PivotTableHandle` — modify-mode proxy for an existing pivot table.

Returned from ``Worksheet.pivot_tables`` for a pivot table that was
parsed off disk. Carries the minimal metadata required to round-trip
a source-range edit:

- ``name`` (read-only) — the ``<pivotTableDefinition name="...">``.
- ``location`` (read-only) — the ``<location ref="A1:E20">`` string.
- ``cache_id`` (read-only) — the workbook-scope cache id.
- ``source`` (read/write) — a :class:`Reference` over the underlying
  ``<cacheSource><worksheetSource>`` element. Setting this stamps a
  new ref + flips ``_dirty``; the actual XML rewrite happens at
  :meth:`Workbook.save` time via ``apply_pivot_source_edits_phase``.

The handle is intentionally thin: full pivot mutation
(field placement, filters, aggregation) is out of scope.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from wolfxl.chart.reference import Reference

if TYPE_CHECKING:
    from wolfxl._workbook import Workbook
    from wolfxl._worksheet import Worksheet


class PivotTableHandle:
    """Modify-mode proxy for an existing pivot table.

    Constructed lazily by :attr:`Worksheet.pivot_tables`. Users do not
    instantiate this directly.
    """

    __slots__ = (
        "_workbook",
        "_owner_sheet",
        "_name",
        "_location",
        "_cache_id",
        "_cache_part_path",
        "_records_part_path",
        "_table_part_path",
        "_orig_source_range",
        "_orig_source_sheet",
        "_orig_field_count",
        "_dirty",
        "_new_source",
    )

    def __init__(
        self,
        *,
        workbook: Workbook,
        owner_sheet: Worksheet,
        name: str,
        location: str,
        cache_id: int,
        cache_part_path: str,
        records_part_path: str,
        table_part_path: str,
        orig_source_range: str,
        orig_source_sheet: str,
        orig_field_count: int,
    ) -> None:
        self._workbook = workbook
        self._owner_sheet = owner_sheet
        self._name = name
        self._location = location
        self._cache_id = cache_id
        self._cache_part_path = cache_part_path
        self._records_part_path = records_part_path
        self._table_part_path = table_part_path
        self._orig_source_range = orig_source_range
        self._orig_source_sheet = orig_source_sheet
        self._orig_field_count = orig_field_count
        self._dirty = False
        self._new_source: Reference | None = None

    # ------------------------------------------------------------------
    # Read-only props
    # ------------------------------------------------------------------

    @property
    def name(self) -> str:
        """``<pivotTableDefinition name="...">`` from the source XML."""
        return self._name

    @property
    def location(self) -> str:
        """``<location ref="A1:E20">`` from the source XML.

        Source-range mutation does NOT alter the pivot's drawn
        location; that is determined by the table part, which v1.0
        passes through unchanged.
        """
        return self._location

    @property
    def cache_id(self) -> int:
        """Workbook-scope ``cacheId`` linking the table to its cache."""
        return self._cache_id

    # ------------------------------------------------------------------
    # source — the only mutator in v1.0
    # ------------------------------------------------------------------

    @property
    def source(self) -> Reference:
        """Current source range. Returns the *pending* range when the
        handle has been mutated this session; otherwise the original
        on-disk range as a synthetic :class:`Reference`.
        """
        if self._new_source is not None:
            return self._new_source
        return self._build_orig_reference()

    @source.setter
    def source(self, value: Reference) -> None:
        """Stamp a new source range. The actual XML rewrite happens at
        save time. Raises :class:`RuntimeError` when the workbook is
        not in modify mode (read-only or write-only contexts cannot
        round-trip pivot edits).
        """
        wb = self._workbook
        if getattr(wb, "_read_only", False):
            raise RuntimeError(
                "PivotTableHandle.source = ... requires modify mode; "
                "this workbook was opened with read_only=True"
            )
        if getattr(wb, "_rust_patcher", None) is None:
            raise RuntimeError(
                "PivotTableHandle.source = ... requires modify mode; "
                "open the workbook with load_workbook(..., modify=True)"
            )
        if not isinstance(value, Reference):
            raise TypeError(
                f"PivotTableHandle.source must be a Reference, got "
                f"{type(value).__name__}"
            )
        self._new_source = value
        self._dirty = True

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _build_orig_reference(self) -> Reference:
        """Build a :class:`Reference` mirroring the on-disk source."""
        from wolfxl.chart.reference import _DummyWorksheet

        ws_obj: Any = None
        if self._orig_source_sheet:
            ws_obj = self._workbook._sheets.get(  # noqa: SLF001
                self._orig_source_sheet
            )
        if ws_obj is None:
            ws_obj = self._owner_sheet
        if ws_obj is None:
            ws_obj = _DummyWorksheet(self._orig_source_sheet)
        rng = self._orig_source_range or "A1"
        bounds = _parse_a1_range(rng)
        if bounds is None:
            # Cannot round-trip an unparseable original; surface a
            # synthetic single-cell reference rather than crash.
            return Reference(ws_obj, min_col=1, min_row=1, max_col=1, max_row=1)
        min_col, min_row, max_col, max_row = bounds
        return Reference(
            ws_obj,
            min_col=min_col,
            min_row=min_row,
            max_col=max_col,
            max_row=max_row,
        )

    def _new_source_to_a1(self) -> str:
        """Render :attr:`_new_source` as an A1 range string for emit."""
        from wolfxl.chart.reference import _index_to_col

        ref = self._new_source
        assert ref is not None  # caller checked _dirty
        if ref.min_col == ref.max_col and ref.min_row == ref.max_row:
            return f"{_index_to_col(ref.min_col)}{ref.min_row}"
        return (
            f"{_index_to_col(ref.min_col)}{ref.min_row}:"
            f"{_index_to_col(ref.max_col)}{ref.max_row}"
        )

    def _new_source_sheet_name(self) -> str:
        """Resolve the new source's sheet name (for ``sheet=``)."""
        ref = self._new_source
        assert ref is not None
        ws_obj = ref.worksheet
        if ws_obj is None:
            return self._orig_source_sheet
        title = getattr(ws_obj, "title", None)
        return str(title) if title is not None else self._orig_source_sheet

    def _column_count(self) -> int:
        """Column span of :attr:`_new_source`."""
        ref = self._new_source
        assert ref is not None
        return int(ref.max_col) - int(ref.min_col) + 1

    def __repr__(self) -> str:
        sfx = " *dirty*" if self._dirty else ""
        return f"<PivotTableHandle name={self._name!r} location={self._location!r}{sfx}>"


def _parse_a1_range(rng: str) -> tuple[int, int, int, int] | None:
    """Parse an A1 range string into 1-based column/row bounds."""
    if not rng:
        return None
    parts = rng.split(":")
    if len(parts) == 1:
        c, r = _parse_a1_cell(parts[0])
        if c is None or r is None:
            return None
        return (c, r, c, r)
    if len(parts) == 2:
        c1, r1 = _parse_a1_cell(parts[0])
        c2, r2 = _parse_a1_cell(parts[1])
        if c1 is None or r1 is None or c2 is None or r2 is None:
            return None
        return (c1, r1, c2, r2)
    return None


def _parse_a1_cell(cell: str) -> tuple[int | None, int | None]:
    """Parse a single A1 cell (e.g. ``$B$5``) into ``(col, row)``."""
    s = cell.lstrip("$")
    col = 0
    i = 0
    while i < len(s) and s[i].isalpha():
        col = col * 26 + (ord(s[i].upper()) - ord("A") + 1)
        i += 1
    if col == 0:
        return (None, None)
    if i < len(s) and s[i] == "$":
        i += 1
    row_str = s[i:]
    if not row_str.isdigit():
        return (None, None)
    return (col, int(row_str))


__all__ = ["PivotTableHandle"]
