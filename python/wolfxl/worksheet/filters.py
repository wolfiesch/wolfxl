"""AutoFilter + filter-class implementations (RFC-056 Sprint Ο Pod 1B).

Replaces the Sprint Ζ stub with the full openpyxl-shaped surface. The
class names + field names match openpyxl's
``openpyxl.worksheet.filters`` so ``from openpyxl.worksheet.filters
import …`` swaps to ``from wolfxl.worksheet.filters import …``
mechanically (Pod 2 wires the re-export shim — this module is the
implementation site).

The 11 filter classes here are PyO3-free dataclasses; they marshal
into the §10 dict shape via :class:`AutoFilter.to_rust_dict` and the
Rust crate ``wolfxl-autofilter`` does both XML emit and filter
evaluation. See ``Plans/rfcs/056-autofilter-eval.md``.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Literal, Optional, Union


# ---------------------------------------------------------------------------
# Filter primitives — mirror RFC-056 §2.1 + openpyxl ``filters.py``.
# ---------------------------------------------------------------------------


@dataclass
class BlankFilter:
    """Pass iff the cell is blank.

    openpyxl emits this as ``<filters blank="1"/>``. Wolfxl emits the
    same shape (RFC-056 §3.1).
    """


@dataclass
class ColorFilter:
    """Pass iff the cell's ``dxfId`` matches.

    ``cell_color = True`` (default) → match against the cell fill.
    ``cell_color = False`` → match against the font colour. Wolfxl's
    evaluator currently accepts every row for ``ColorFilter`` (defers
    to Excel on open) — KNOWN_GAPS entry. The XML round-trips
    losslessly.
    """

    dxf_id: int = 0
    cell_color: bool = True


@dataclass
class CustomFilter:
    """One ``<customFilter operator val>`` row.

    ``operator`` ∈ {"equal", "lessThan", "lessThanOrEqual",
    "notEqual", "greaterThanOrEqual", "greaterThan"}.
    """

    operator: Literal[
        "equal",
        "lessThan",
        "lessThanOrEqual",
        "notEqual",
        "greaterThanOrEqual",
        "greaterThan",
    ] = "equal"
    val: str = ""


@dataclass
class CustomFilters:
    """Group of ``CustomFilter`` joined by AND or OR.

    ``and_ = False`` (default) → logical OR.
    ``and_ = True`` → logical AND.
    """

    customFilter: list[CustomFilter] = field(default_factory=list)  # noqa: N815 — openpyxl name
    and_: bool = False

    @property
    def filters(self) -> list[CustomFilter]:
        """Alias for ``customFilter`` (the snake_case is more readable)."""
        return self.customFilter

    @filters.setter
    def filters(self, value: list[CustomFilter]) -> None:
        self.customFilter = value


@dataclass
class DateGroupItem:
    """Date-component matcher inside a ``<filters>`` group.

    Components below ``date_time_grouping`` are typically ``None``.
    """

    year: int = 0
    month: Optional[int] = None
    day: Optional[int] = None
    hour: Optional[int] = None
    minute: Optional[int] = None
    second: Optional[int] = None
    date_time_grouping: Literal[
        "year", "month", "day", "hour", "minute", "second"
    ] = "year"


@dataclass
class DynamicFilter:
    """Date- or aggregate-driven dynamic filter.

    Type values (subset shown) include
    ``"today"``, ``"yesterday"``, ``"thisWeek"``, ``"Q1".."Q4"``,
    ``"M1".."M12"``, ``"aboveAverage"``, ``"belowAverage"``,
    ``"yearToDate"``. Full list in RFC-056 §2.1.
    """

    type: str = "null"
    val: Optional[float] = None
    val_iso: Optional[str] = None
    max_val_iso: Optional[str] = None


@dataclass
class IconFilter:
    """Filter on conditional-format icon set + index.

    ``icon_set`` is the named CF icon set (e.g. ``"3Arrows"``,
    ``"5Quarters"``); ``icon_id`` is the 0-based index into it.
    Like ``ColorFilter``, evaluation is best-effort and defers to
    Excel.
    """

    icon_set: str = "3Arrows"
    icon_id: int = 0


@dataclass
class NumberFilter:
    """Numeric whitelist with optional blank pass-through."""

    filters: list[float] = field(default_factory=list)
    blank: bool = False
    calendar_type: Optional[str] = None


@dataclass
class StringFilter:
    """String whitelist (case-insensitive per Excel)."""

    values: list[str] = field(default_factory=list)


@dataclass
class Top10:
    """Top-N or bottom-N filter, optionally as percent.

    ``top = True`` → top N. ``top = False`` → bottom N.
    ``percent = True`` → N is a percentage of the column.
    ``filter_val`` is set by Excel after evaluation; readers may
    inspect it but writers should leave it ``None`` (the patcher
    does not back-fill it).
    """

    top: bool = True
    percent: bool = False
    val: float = 10.0
    filter_val: Optional[float] = None


# ---------------------------------------------------------------------------
# FilterColumn + Sort
# ---------------------------------------------------------------------------


# Public discriminated union of every filter kind a FilterColumn carries.
FilterT = Union[
    BlankFilter,
    ColorFilter,
    CustomFilters,
    DynamicFilter,
    IconFilter,
    NumberFilter,
    StringFilter,
    Top10,
]


@dataclass
class FilterColumn:
    """One ``<filterColumn colId>`` entry inside ``<autoFilter>``.

    ``col_id`` is **0-based** relative to ``auto_filter.ref``'s left
    edge — Excel's convention for this attribute (every other column
    index in OOXML is 1-based).
    """

    col_id: int = 0
    hidden_button: bool = False
    show_button: bool = True
    filter: Optional[FilterT] = None
    date_group_items: list[DateGroupItem] = field(default_factory=list)


@dataclass
class SortCondition:
    """One ``<sortCondition>`` inside ``<sortState>``."""

    ref: str = ""
    descending: bool = False
    sort_by: Literal["value", "cellColor", "fontColor", "icon"] = "value"
    custom_list: Optional[str] = None
    dxf_id: Optional[int] = None
    icon_set: Optional[str] = None
    icon_id: Optional[int] = None


@dataclass
class SortState:
    """``<sortState>`` block.

    .. note::

       v2.0 ships XML round-trip + ``sort_order`` permutation only.
       Physical row reordering is **deferred to v2.1** (RFC-056 §8).
    """

    sort_conditions: list[SortCondition] = field(default_factory=list)
    column_sort: bool = False
    case_sensitive: bool = False
    ref: Optional[str] = None


# ---------------------------------------------------------------------------
# Top-level AutoFilter — the dataclass version. The thin proxy on
# Worksheet (``ws.auto_filter``) wraps an instance of this class.
# ---------------------------------------------------------------------------


@dataclass
class AutoFilter:
    """``ws.auto_filter`` — sheet-scoped filter state.

    ``ref`` is the A1-notation range. ``filter_columns`` is the list
    of ``FilterColumn`` entries; ``sort_state`` is optional.

    See ``Worksheet.auto_filter`` for the friendlier in-place builder
    methods (``add_filter_column``, ``add_sort_condition``).
    """

    ref: Optional[str] = None
    filter_columns: list[FilterColumn] = field(default_factory=list)
    sort_state: Optional[SortState] = None

    # ------------------------------------------------------------------
    # Builders matching openpyxl-style fluency.
    # ------------------------------------------------------------------

    def add_filter_column(
        self,
        col_id: int,
        filter: Optional[FilterT] = None,
        *,
        hidden_button: bool = False,
        show_button: bool = True,
        date_group_items: Optional[list[DateGroupItem]] = None,
    ) -> FilterColumn:
        """Add a ``FilterColumn``. Returns the new entry for chaining."""
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
        custom_list: Optional[str] = None,
        dxf_id: Optional[int] = None,
        icon_set: Optional[str] = None,
        icon_id: Optional[int] = None,
    ) -> SortCondition:
        """Add a ``SortCondition``. Auto-creates ``sort_state`` if absent."""
        if self.sort_state is None:
            self.sort_state = SortState(ref=ref)
        sc = SortCondition(
            ref=ref,
            descending=descending,
            sort_by=sort_by,  # type: ignore[arg-type]
            custom_list=custom_list,
            dxf_id=dxf_id,
            icon_set=icon_set,
            icon_id=icon_id,
        )
        self.sort_state.sort_conditions.append(sc)
        return sc

    # ------------------------------------------------------------------
    # §10 dict marshalling.
    # ------------------------------------------------------------------

    def to_rust_dict(self) -> dict[str, Any]:
        """Serialise to the RFC-056 §10 dict shape consumed by the Rust
        ``wolfxl-autofilter`` crate via the PyO3 boundary.
        """
        return {
            "ref": self.ref,
            "filter_columns": [_filter_column_to_dict(fc) for fc in self.filter_columns],
            "sort_state": _sort_state_to_dict(self.sort_state) if self.sort_state else None,
        }


# ---------------------------------------------------------------------------
# §10 dict helpers (private).
# ---------------------------------------------------------------------------


def _filter_column_to_dict(fc: FilterColumn) -> dict[str, Any]:
    return {
        "col_id": int(fc.col_id),
        "hidden_button": bool(fc.hidden_button),
        "show_button": bool(fc.show_button),
        "filter": _filter_to_dict(fc.filter) if fc.filter is not None else None,
        "date_group_items": [_date_group_item_to_dict(d) for d in fc.date_group_items],
    }


def _filter_to_dict(f: FilterT) -> dict[str, Any]:
    if isinstance(f, BlankFilter):
        return {"kind": "blank"}
    if isinstance(f, ColorFilter):
        return {
            "kind": "color",
            "dxf_id": int(f.dxf_id),
            "cell_color": bool(f.cell_color),
        }
    if isinstance(f, CustomFilters):
        return {
            "kind": "custom",
            "and_": bool(f.and_),
            "filters": [
                {"operator": cf.operator, "val": str(cf.val)} for cf in f.customFilter
            ],
        }
    if isinstance(f, CustomFilter):
        # Lone CustomFilter — wrap as a one-element CustomFilters to
        # keep the §10 shape uniform.
        return {
            "kind": "custom",
            "and_": False,
            "filters": [{"operator": f.operator, "val": str(f.val)}],
        }
    if isinstance(f, DynamicFilter):
        return {
            "kind": "dynamic",
            "type": f.type,
            "val": f.val,
            "val_iso": f.val_iso,
            "max_val_iso": f.max_val_iso,
        }
    if isinstance(f, IconFilter):
        return {
            "kind": "icon",
            "icon_set": f.icon_set,
            "icon_id": int(f.icon_id),
        }
    if isinstance(f, NumberFilter):
        return {
            "kind": "number",
            "filters": [float(v) for v in f.filters],
            "blank": bool(f.blank),
            "calendar_type": f.calendar_type,
        }
    if isinstance(f, StringFilter):
        return {
            "kind": "string",
            "values": [str(v) for v in f.values],
        }
    if isinstance(f, Top10):
        return {
            "kind": "top10",
            "top": bool(f.top),
            "percent": bool(f.percent),
            "val": float(f.val),
            "filter_val": f.filter_val,
        }
    raise TypeError(f"unknown filter type: {type(f)!r}")


def _date_group_item_to_dict(d: DateGroupItem) -> dict[str, Any]:
    return {
        "year": int(d.year),
        "month": int(d.month) if d.month is not None else None,
        "day": int(d.day) if d.day is not None else None,
        "hour": int(d.hour) if d.hour is not None else None,
        "minute": int(d.minute) if d.minute is not None else None,
        "second": int(d.second) if d.second is not None else None,
        "date_time_grouping": d.date_time_grouping,
    }


def _sort_state_to_dict(s: SortState) -> dict[str, Any]:
    return {
        "sort_conditions": [_sort_condition_to_dict(sc) for sc in s.sort_conditions],
        "column_sort": bool(s.column_sort),
        "case_sensitive": bool(s.case_sensitive),
        "ref": s.ref,
    }


def _sort_condition_to_dict(sc: SortCondition) -> dict[str, Any]:
    return {
        "ref": sc.ref,
        "descending": bool(sc.descending),
        "sort_by": sc.sort_by,
        "custom_list": sc.custom_list,
        "dxf_id": int(sc.dxf_id) if sc.dxf_id is not None else None,
        "icon_set": sc.icon_set,
        "icon_id": int(sc.icon_id) if sc.icon_id is not None else None,
    }


# Aliased name kept for openpyxl compatibility — its module exports
# `Filters` (the inner `<filters>` element); we expose it as an alias
# to NumberFilter since they share XML shape.
Filters = NumberFilter


__all__ = [
    "AutoFilter",
    "BlankFilter",
    "ColorFilter",
    "CustomFilter",
    "CustomFilters",
    "DateGroupItem",
    "DynamicFilter",
    "FilterColumn",
    "Filters",
    "IconFilter",
    "NumberFilter",
    "SortCondition",
    "SortState",
    "StringFilter",
    "Top10",
]
