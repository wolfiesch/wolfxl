"""openpyxl-compatible utility surface.

Mirrors the symbols SynthGL imports from ``openpyxl.utils.*`` and
``openpyxl.styles.numbers.is_date_format``. The behavior is deliberately
*bug-for-bug* identical to openpyxl 3.1+ — pinned by
``tests/parity/test_utils_parity.py``. Any divergence is a parity bug.
"""

from __future__ import annotations

from wolfxl.utils.cell import (
    column_index_from_string,
    coordinate_to_tuple,
    get_column_letter,
    range_boundaries,
)
from wolfxl.utils.datetime import (
    CALENDAR_WINDOWS_1900,
    from_excel,
)
from wolfxl.utils.numbers import is_date_format

__all__ = [
    "CALENDAR_WINDOWS_1900",
    "column_index_from_string",
    "coordinate_to_tuple",
    "from_excel",
    "get_column_letter",
    "is_date_format",
    "range_boundaries",
]
