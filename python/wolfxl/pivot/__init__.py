"""``wolfxl.pivot`` — pivot-table construction (Sprint Ν, v2.0.0).

Mirrors :mod:`openpyxl.pivot`. The public surface is:

- :class:`PivotCache` — the schema + records snapshot for a source
  range. Constructed by the user, registered via
  :meth:`Workbook.add_pivot_cache`, and emitted as
  ``xl/pivotCache/pivotCacheDefinition{N}.xml`` plus a companion
  ``pivotCacheRecords{N}.xml``.
- :class:`PivotTable` — the layout (rows/cols/data/page assignments)
  pointing at a cache. Registered via
  :meth:`Worksheet.add_pivot_table`. Emitted as
  ``xl/pivotTables/pivotTable{N}.xml``.
- :class:`PivotField` / :class:`DataField` / :class:`RowField` /
  :class:`ColumnField` / :class:`PageField` — explicit builders for
  pivot-table axes when the bare-string convenience form
  (``rows=["region"]``) doesn't fit.
- :class:`Reference` — re-exported from :mod:`wolfxl.chart.reference`
  for source-range construction (the OOXML cache uses the exact same
  shape).

See the §10 contracts in
``Plans/rfcs/047-pivot-caches.md`` and
``Plans/rfcs/048-pivot-tables.md`` for the authoritative dict shape
emitted by ``to_rust_dict()``.

# Sprint Ν status

This module replaces the v0.5.0+ ``_make_stub`` with real
construction. The Rust emit functions live in ``crates/wolfxl-pivot``
(PyO3-free) and are reached via the ``wolfxl._rust`` bindings:

- ``serialize_pivot_cache_dict(d) -> bytes``
- ``serialize_pivot_records_dict(d) -> bytes``
- ``serialize_pivot_table_dict(d) -> bytes``

Pod-γ wires those bindings during patcher Phase 2.5m and into the
native writer.
"""

from __future__ import annotations

from wolfxl.chart.reference import Reference
from ._cache import (
    CacheField,
    CacheValue,
    PivotCache,
    SharedItems,
    WorksheetSource,
)
from ._table import (
    ColumnField,
    DataField,
    DataFunction,
    Location,
    PageField,
    PivotField,
    PivotItem,
    PivotSource,
    PivotTable,
    PivotTableStyleInfo,
    RowField,
)
# RFC-061 sub-features.
from ._slicer import Slicer, SlicerCache, SlicerItem
from ._calc import CalculatedField, CalculatedItem
from ._group import FieldGroup, FieldGroupRange, FieldGroupDate
from ._styling import (
    ChartFormat,
    Format,
    PivotArea,
    PivotConditionalFormat,
)

__all__ = [
    # Cache layer (RFC-047)
    "PivotCache",
    "CacheField",
    "SharedItems",
    "CacheValue",
    "WorksheetSource",
    # Table layer (RFC-048)
    "PivotTable",
    "PivotField",
    "DataField",
    "DataFunction",
    "RowField",
    "ColumnField",
    "PageField",
    "PivotItem",
    "Location",
    "PivotTableStyleInfo",
    # Chart linkage (RFC-049)
    "PivotSource",
    # Shared with charts
    "Reference",
    # RFC-061 — Slicers (§2.1)
    "Slicer",
    "SlicerCache",
    "SlicerItem",
    # RFC-061 — Calculated fields/items (§2.2 / §2.3)
    "CalculatedField",
    "CalculatedItem",
    # RFC-061 — Group items (§2.4)
    "FieldGroup",
    "FieldGroupRange",
    "FieldGroupDate",
    # RFC-061 — Pivot styling (§2.5)
    "Format",
    "PivotArea",
    "PivotConditionalFormat",
    "ChartFormat",
]
