"""``openpyxl.pivot.cache`` — re-export shim.

Wolfxl's pivot-cache classes live in :mod:`wolfxl.pivot._cache`; the
package init re-exports them.  This module surfaces the same surface
under the openpyxl-shaped explicit-module path so
``from openpyxl.pivot.cache import CacheDefinition`` ports mechanically.
"""

from __future__ import annotations

from wolfxl.pivot._cache import (
    CacheField,
    CacheValue,
    PivotCache,
    SharedItems,
    WorksheetSource,
)


# openpyxl exposes ``CacheDefinition`` as the underlying name; alias for
# import-compat parity.
CacheDefinition = PivotCache


__all__ = [
    "CacheDefinition",
    "CacheField",
    "CacheValue",
    "PivotCache",
    "SharedItems",
    "WorksheetSource",
]
