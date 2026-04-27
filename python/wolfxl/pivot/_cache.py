"""``PivotCache`` — schema + records snapshot for a pivot's source range.

Mirrors :class:`openpyxl.pivot.cache.CacheDefinition` plus the
records part. See RFC-047 §10 for the authoritative dict contract
returned by :meth:`PivotCache.to_rust_dict` and the helper
:meth:`PivotCache.to_rust_records_dict`.

Construction-time validation matches RFC-047 §10.8.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Any, Iterable

from wolfxl.chart.reference import Reference

# RFC-047 §10.9 inference thresholds. Tunable but documented.
_INFER_THRESHOLDS = {
    "max_string_unique_for_enumeration": 2000,
    "max_number_unique_for_enumeration": 200,
}


# ---------------------------------------------------------------------------
# CacheValue — the variant emitted to the §10.5 contract.
# ---------------------------------------------------------------------------


class CacheValue:
    """Tagged variant matching RFC-047 §10.5.

    Use the classmethod constructors (``CacheValue.string("North")``)
    rather than the dataclass init for clarity at call sites.
    """

    __slots__ = ("kind", "value")

    def __init__(self, kind: str, value: Any = None) -> None:
        self.kind = kind
        self.value = value

    @classmethod
    def string(cls, s: str) -> "CacheValue":
        return cls("string", s)

    @classmethod
    def number(cls, n: float) -> "CacheValue":
        return cls("number", float(n))

    @classmethod
    def boolean(cls, b: bool) -> "CacheValue":
        return cls("boolean", bool(b))

    @classmethod
    def date(cls, d: str | date | datetime) -> "CacheValue":
        if isinstance(d, datetime):
            iso = d.isoformat()
        elif isinstance(d, date):
            iso = f"{d.isoformat()}T00:00:00"
        else:
            iso = str(d)
        return cls("date", iso)

    @classmethod
    def missing(cls) -> "CacheValue":
        return cls("missing", None)

    @classmethod
    def error(cls, s: str) -> "CacheValue":
        return cls("error", s)

    def to_rust_dict(self) -> dict:
        if self.value is None:
            return {"kind": self.kind}
        return {"kind": self.kind, "value": self.value}

    def __repr__(self) -> str:
        if self.value is None:
            return f"CacheValue.{self.kind}()"
        return f"CacheValue.{self.kind}({self.value!r})"

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, CacheValue):
            return NotImplemented
        return self.kind == other.kind and self.value == other.value

    def __hash__(self) -> int:
        return hash((self.kind, self.value))


# ---------------------------------------------------------------------------
# WorksheetSource
# ---------------------------------------------------------------------------


@dataclass
class WorksheetSource:
    """RFC-047 §10.2.

    Either ``(sheet, ref)`` or ``name``, mutually exclusive.
    """

    sheet: str = ""
    ref: str = ""
    name: str | None = None

    def __post_init__(self) -> None:
        has_sheet_ref = bool(self.sheet) or bool(self.ref)
        has_name = self.name is not None and self.name != ""
        if not has_sheet_ref and not has_name:
            raise ValueError(
                "WorksheetSource requires sheet+ref or name"
            )
        if has_sheet_ref and has_name:
            raise ValueError(
                "WorksheetSource: sheet+ref and name are mutually exclusive"
            )

    def to_rust_dict(self) -> dict:
        return {"sheet": self.sheet, "ref": self.ref, "name": self.name}


# ---------------------------------------------------------------------------
# SharedItems
# ---------------------------------------------------------------------------


@dataclass
class SharedItems:
    """RFC-047 §10.4.

    ``items=None`` + numeric attrs (min/max) → numeric-only no-enumeration
    form (``<sharedItems containsNumber="1" minValue=… maxValue=…/>``).

    ``items=[…]`` → enumeration form (``<s>``/``<n>``/``<d>``/etc.
    children).
    """

    count: int | None = None
    items: list[CacheValue] | None = None
    contains_blank: bool = False
    contains_mixed_types: bool = False
    contains_semi_mixed_types: bool = True
    contains_string: bool = True
    contains_number: bool = False
    contains_integer: bool = False
    contains_date: bool = False
    contains_non_date: bool = True
    min_value: float | None = None
    max_value: float | None = None
    min_date: str | None = None
    max_date: str | None = None
    long_text: bool = False

    def to_rust_dict(self) -> dict:
        return {
            "count": self.count,
            "items": (
                [v.to_rust_dict() for v in self.items]
                if self.items is not None
                else None
            ),
            "contains_blank": self.contains_blank,
            "contains_mixed_types": self.contains_mixed_types,
            "contains_semi_mixed_types": self.contains_semi_mixed_types,
            "contains_string": self.contains_string,
            "contains_number": self.contains_number,
            "contains_integer": self.contains_integer,
            "contains_date": self.contains_date,
            "contains_non_date": self.contains_non_date,
            "min_value": self.min_value,
            "max_value": self.max_value,
            "min_date": self.min_date,
            "max_date": self.max_date,
            "long_text": self.long_text,
        }


# ---------------------------------------------------------------------------
# CacheField
# ---------------------------------------------------------------------------


@dataclass
class CacheField:
    """RFC-047 §10.3.

    `data_type` is one of ``"string" | "number" | "date" | "bool" |
    "mixed"`` matching RFC-047 §10.3. ``formula`` and ``hierarchy``
    are reserved for v2.1+ (RFC-047 §11) and pinned to ``None`` here.
    """

    name: str
    num_fmt_id: int = 0
    data_type: str = "string"
    shared_items: SharedItems = field(default_factory=SharedItems)
    formula: str | None = None
    hierarchy: int | None = None

    def to_rust_dict(self) -> dict:
        return {
            "name": self.name,
            "num_fmt_id": self.num_fmt_id,
            "data_type": self.data_type,
            "shared_items": self.shared_items.to_rust_dict(),
            "formula": self.formula,
            "hierarchy": self.hierarchy,
        }


# ---------------------------------------------------------------------------
# PivotCache (top-level)
# ---------------------------------------------------------------------------


class PivotCache:
    """The pivot cache — schema + records snapshot.

    Constructed by the user with a :class:`Reference` source range, then
    registered via :meth:`Workbook.add_pivot_cache(cache)` which:

    1. allocates ``cache_id`` (0-based, returned to the caller);
    2. walks the source range to materialize :attr:`fields` and
       :attr:`records` (RFC-047 §10.9 inference);
    3. enqueues the cache for emit at patcher Phase 2.5m.

    Until step 2 runs, :attr:`fields` and :attr:`records` are
    ``None`` and :meth:`to_rust_dict` raises.
    """

    def __init__(
        self,
        source: Reference,
        *,
        refresh_on_load: bool = False,
        refreshed_by: str = "wolfxl",
    ) -> None:
        if not isinstance(source, Reference):
            raise TypeError(
                f"PivotCache(source=...) must be a Reference, got {type(source).__name__}"
            )
        self.source = source
        self.refresh_on_load = refresh_on_load
        self.refreshed_by = refreshed_by
        # Set by Workbook.add_pivot_cache().
        self._cache_id: int | None = None
        # Set by _materialize() (called by add_pivot_cache).
        self._fields: list[CacheField] | None = None
        self._records: list[list[CacheValue]] | None = None

    @property
    def cache_id(self) -> int:
        if self._cache_id is None:
            raise RuntimeError(
                "PivotCache has not been registered yet — call "
                "Workbook.add_pivot_cache(cache) first"
            )
        return self._cache_id

    @property
    def fields(self) -> list[CacheField]:
        if self._fields is None:
            raise RuntimeError(
                "PivotCache has not been materialized yet — call "
                "Workbook.add_pivot_cache(cache) first"
            )
        return self._fields

    @property
    def field_names(self) -> list[str]:
        return [f.name for f in self.fields]

    def field_index(self, name: str) -> int:
        """Look up a field by name; raises ``KeyError`` if not found.

        Used by :class:`PivotTable` to resolve string field names
        (``rows=["region"]``) to indices.
        """
        for i, f in enumerate(self.fields):
            if f.name == name:
                return i
        raise KeyError(
            f"PivotCache has no field named {name!r}; "
            f"available: {self.field_names}"
        )

    # ------------------------------------------------------------------
    # Materialization — walk the source range, infer types, build
    # SharedItems and records. Called by Workbook.add_pivot_cache().
    # ------------------------------------------------------------------

    def _materialize(self, ws: Any) -> None:
        """Walk ``ws[self.source.range]``; build fields + records.

        Per-column type inference (RFC-047 §10.9):

        =======================  =============  =====================================
        Column observation       data_type      shared_items strategy
        =======================  =============  =====================================
        all numeric, ≤200 unique  number         enumerate items
        all numeric, >200 unique  number         min/max attrs only (no items)
        all string, ≤2000 unique  string         enumerate items
        all string, >2000 unique  string         attrs only
        all ISO date              date           enumerate dates as <d> items
        mixed types               mixed          contains_semi_mixed_types=True
        all None                  string         contains_blank=True, count=0
        =======================  =============  =====================================
        """
        rows = self._collect_source_rows(ws)
        if not rows:
            raise ValueError(
                f"PivotCache.source ({self.source}) is empty — "
                "needs ≥1 header row + ≥1 data row"
            )
        header = rows[0]
        data_rows = rows[1:]
        if not data_rows:
            raise ValueError(
                "PivotCache.source has only a header row; "
                "needs ≥1 data row"
            )

        n_cols = len(header)
        cache_fields: list[CacheField] = []
        records: list[list[CacheValue]] = [[] for _ in data_rows]

        for col_idx in range(n_cols):
            field_name = self._cell_to_field_name(header[col_idx], col_idx)
            col_values = [row[col_idx] if col_idx < len(row) else None
                          for row in data_rows]
            cf = self._infer_cache_field(field_name, col_values)
            cache_fields.append(cf)
            # Materialize per-row record cell using inferred field.
            for ri, raw in enumerate(col_values):
                cv = self._raw_to_cache_value(raw)
                records[ri].append(cv)

        self._fields = cache_fields
        self._records = records

    def _collect_source_rows(self, ws: Any) -> list[list[Any]]:
        """Iterate worksheet rows in :attr:`source` range; return values."""
        from wolfxl.chart.reference import _index_to_col

        if self.source.range_string:
            # Reference parsed from a string is portable across sheets.
            ws_target = self._resolve_worksheet(ws)
        else:
            ws_target = self.source.worksheet

        rows: list[list[Any]] = []
        for r in range(self.source.min_row, self.source.max_row + 1):
            row_vals: list[Any] = []
            for c in range(self.source.min_col, self.source.max_col + 1):
                col_letter = _index_to_col(c)
                addr = f"{col_letter}{r}"
                cell = ws_target[addr]
                row_vals.append(cell.value if hasattr(cell, "value") else cell)
            rows.append(row_vals)
        return rows

    def _resolve_worksheet(self, wb_or_ws: Any) -> Any:
        """If a workbook is passed, resolve via source's sheet name;
        if a worksheet is passed, return it directly."""
        title = self.source.worksheet.title if self.source.worksheet else None
        if hasattr(wb_or_ws, "sheetnames") and title is not None:
            return wb_or_ws[title]
        return wb_or_ws  # assume it's already a worksheet

    @staticmethod
    def _cell_to_field_name(cell: Any, col_idx: int) -> str:
        if cell is None or cell == "":
            return f"Field{col_idx + 1}"
        return str(cell)

    def _infer_cache_field(
        self, name: str, values: list[Any]
    ) -> CacheField:
        # Classify per-value type; pick dominant.
        kinds = [self._classify(v) for v in values]
        non_missing = [k for k in kinds if k != "missing"]
        unique_kinds = set(non_missing)

        if not unique_kinds:
            # All None.
            return CacheField(
                name=name,
                data_type="string",
                shared_items=SharedItems(
                    count=0,
                    items=[],
                    contains_blank=True,
                    contains_semi_mixed_types=False,
                    contains_string=False,
                    contains_non_date=False,
                ),
            )

        if unique_kinds == {"number"}:
            return self._infer_numeric_field(name, values)
        if unique_kinds == {"string"}:
            return self._infer_string_field(name, values)
        if unique_kinds == {"date"}:
            return self._infer_date_field(name, values)
        if unique_kinds == {"boolean"}:
            return self._infer_boolean_field(name, values)
        # Mixed.
        return self._infer_mixed_field(name, values)

    @staticmethod
    def _classify(v: Any) -> str:
        if v is None:
            return "missing"
        if isinstance(v, bool):
            return "boolean"
        if isinstance(v, (int, float)):
            return "number"
        if isinstance(v, (date, datetime)):
            return "date"
        return "string"

    def _infer_numeric_field(
        self, name: str, values: list[Any]
    ) -> CacheField:
        nums = [float(v) for v in values if v is not None]
        unique_nums = set(nums)
        threshold = _INFER_THRESHOLDS["max_number_unique_for_enumeration"]
        if len(unique_nums) <= threshold:
            # Enumerate.
            ordered = sorted(unique_nums)
            items = [CacheValue.number(n) for n in ordered]
            si = SharedItems(
                count=len(items),
                items=items,
                contains_semi_mixed_types=False,
                contains_string=False,
                contains_number=True,
                contains_integer=all(float(n).is_integer() for n in ordered),
                min_value=min(ordered) if ordered else None,
                max_value=max(ordered) if ordered else None,
            )
        else:
            # Attrs only.
            si = SharedItems(
                count=None,
                items=None,
                contains_semi_mixed_types=False,
                contains_string=False,
                contains_number=True,
                contains_integer=all(float(n).is_integer() for n in nums),
                min_value=min(nums) if nums else None,
                max_value=max(nums) if nums else None,
            )
        return CacheField(name=name, data_type="number", shared_items=si)

    def _infer_string_field(
        self, name: str, values: list[Any]
    ) -> CacheField:
        strs = [str(v) for v in values if v is not None]
        unique = list(dict.fromkeys(strs))  # preserve insertion order
        threshold = _INFER_THRESHOLDS["max_string_unique_for_enumeration"]
        if len(unique) <= threshold:
            items = [CacheValue.string(s) for s in unique]
            si = SharedItems(
                count=len(items),
                items=items,
                contains_blank=any(v is None for v in values),
                contains_semi_mixed_types=True,
                contains_string=True,
            )
        else:
            si = SharedItems(
                count=None,
                items=None,
                contains_blank=any(v is None for v in values),
                contains_semi_mixed_types=True,
                contains_string=True,
                long_text=any(len(s) > 256 for s in strs),
            )
        return CacheField(name=name, data_type="string", shared_items=si)

    def _infer_date_field(
        self, name: str, values: list[Any]
    ) -> CacheField:
        dates = [self._date_to_iso(v) for v in values if v is not None]
        unique = list(dict.fromkeys(dates))
        items = [CacheValue.date(d) for d in unique]
        si = SharedItems(
            count=len(items),
            items=items,
            contains_blank=any(v is None for v in values),
            contains_semi_mixed_types=False,
            contains_string=False,
            contains_date=True,
            contains_non_date=False,
            min_date=min(unique) if unique else None,
            max_date=max(unique) if unique else None,
        )
        return CacheField(name=name, data_type="date", shared_items=si)

    def _infer_boolean_field(
        self, name: str, values: list[Any]
    ) -> CacheField:
        bools = [bool(v) for v in values if v is not None]
        unique = list(dict.fromkeys(bools))
        items = [CacheValue.boolean(b) for b in unique]
        si = SharedItems(
            count=len(items),
            items=items,
            contains_blank=any(v is None for v in values),
            contains_semi_mixed_types=False,
            contains_string=False,
        )
        return CacheField(name=name, data_type="bool", shared_items=si)

    def _infer_mixed_field(
        self, name: str, values: list[Any]
    ) -> CacheField:
        items: list[CacheValue] = []
        seen: set = set()
        for v in values:
            if v is None:
                continue
            cv = self._raw_to_cache_value(v)
            key = (cv.kind, cv.value)
            if key not in seen:
                seen.add(key)
                items.append(cv)
        si = SharedItems(
            count=len(items),
            items=items,
            contains_blank=any(v is None for v in values),
            contains_mixed_types=True,
            contains_semi_mixed_types=True,
            contains_string=any(c.kind == "string" for c in items),
            contains_number=any(c.kind == "number" for c in items),
            contains_date=any(c.kind == "date" for c in items),
            contains_non_date=not all(c.kind == "date" for c in items),
        )
        return CacheField(name=name, data_type="mixed", shared_items=si)

    @staticmethod
    def _date_to_iso(v: Any) -> str:
        if isinstance(v, datetime):
            return v.isoformat()
        if isinstance(v, date):
            return f"{v.isoformat()}T00:00:00"
        return str(v)

    def _raw_to_cache_value(self, v: Any) -> CacheValue:
        if v is None:
            return CacheValue.missing()
        if isinstance(v, bool):
            return CacheValue.boolean(v)
        if isinstance(v, (int, float)):
            return CacheValue.number(float(v))
        if isinstance(v, (date, datetime)):
            return CacheValue.date(v)
        return CacheValue.string(str(v))

    # ------------------------------------------------------------------
    # to_rust_dict — RFC-047 §10.1 + §10.6 contracts.
    # ------------------------------------------------------------------

    def to_rust_dict(self) -> dict:
        """Cache-definition dict per RFC-047 §10.1."""
        if self._cache_id is None:
            raise RuntimeError(
                "PivotCache.cache_id not set — call "
                "Workbook.add_pivot_cache() first"
            )
        if self._fields is None:
            raise RuntimeError(
                "PivotCache._materialize not yet called"
            )
        return {
            "cache_id": self._cache_id,
            "source": self.source_to_rust_dict(),
            "fields": [f.to_rust_dict() for f in self._fields],
            "refresh_on_load": self.refresh_on_load,
            "refreshed_version": 6,
            "created_version": 6,
            "min_refreshable_version": 3,
            "refreshed_by": self.refreshed_by,
            "records_part_path": None,  # set by patcher
        }

    def to_rust_records_dict(self) -> dict:
        """Records dict per RFC-047 §10.6.

        Each record cell is converted to inline form unless the field
        has an enumerable ``shared_items.items`` — in which case the
        cell is rewritten as ``{"kind": "index", "value": N}`` where
        N is the item's index.
        """
        if self._records is None or self._fields is None:
            raise RuntimeError(
                "PivotCache._materialize not yet called"
            )
        index_lookups = []
        for f in self._fields:
            si = f.shared_items
            if si.items is not None:
                # Map (kind, value) → 0-based index.
                lookup = {(v.kind, v.value): i for i, v in enumerate(si.items)}
                index_lookups.append(lookup)
            else:
                index_lookups.append(None)

        out_records = []
        for row in self._records:
            cells = []
            for col_i, cv in enumerate(row):
                lookup = index_lookups[col_i]
                if lookup is not None and cv.kind != "missing":
                    key = (cv.kind, cv.value)
                    if key in lookup:
                        cells.append({"kind": "index", "value": lookup[key]})
                        continue
                # Inline form.
                cells.append(cv.to_rust_dict())
            out_records.append(cells)

        return {
            "field_count": len(self._fields),
            "record_count": len(out_records),
            "records": out_records,
        }

    def source_to_rust_dict(self) -> dict:
        sheet = (
            self.source.worksheet.title
            if self.source.worksheet is not None
            else ""
        )
        ref = self._reference_to_a1()
        return {"sheet": sheet, "ref": ref, "name": None}

    def _reference_to_a1(self) -> str:
        from wolfxl.chart.reference import _index_to_col

        if self.source.min_col == self.source.max_col and self.source.min_row == self.source.max_row:
            return f"{_index_to_col(self.source.min_col)}{self.source.min_row}"
        return (
            f"{_index_to_col(self.source.min_col)}{self.source.min_row}:"
            f"{_index_to_col(self.source.max_col)}{self.source.max_row}"
        )
