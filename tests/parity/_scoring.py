"""Scoring tiers for the parity harness.

Borrowed from ExcelBench's tiered fidelity model. Each dimension declares
the set of attributes that participate and the tier that governs them:

* ``HARD`` - any mismatch fails CI immediately.
* ``SOFT`` - count is tracked in ``ratchet.json``; the count is never
  allowed to *increase* over the committed baseline (regressions block CI).
* ``INFO`` - reported only, not gated.

See ``Plans`` section "Pass semantics" for the full table.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any


class Tier(str, Enum):
    """``str`` mixin keeps ``Tier.HARD == "hard"`` true on Python 3.9-3.10
    (``enum.StrEnum`` is 3.11+)."""

    HARD = "hard"
    SOFT = "soft"
    INFO = "info"


# Dimension-to-tier map. When adding a dimension, classify it here or the
# harness will fail loud — there is no implicit default.
DIMENSION_TIERS: dict[str, Tier] = {
    # ---- HARD (any mismatch fails CI) ----------------------------------
    "value": Tier.HARD,
    "number_format": Tier.HARD,
    "merged_cells": Tier.HARD,
    "max_row": Tier.HARD,
    "max_col": Tier.HARD,
    "freeze_panes": Tier.HARD,
    "defined_name.refers_to": Tier.HARD,
    "column_width": Tier.HARD,
    "utils.get_column_letter": Tier.HARD,
    "utils.column_index_from_string": Tier.HARD,
    "utils.range_boundaries": Tier.HARD,
    "utils.coordinate_to_tuple": Tier.HARD,
    "utils.is_date_format": Tier.HARD,
    "utils.from_excel": Tier.HARD,

    # ---- SOFT (tracked in ratchet, never allowed to drop) --------------
    "font.name": Tier.SOFT,
    "font.size": Tier.SOFT,
    "font.bold": Tier.SOFT,
    "font.italic": Tier.SOFT,
    "fill.fg_color": Tier.SOFT,
    "border.style": Tier.SOFT,
    "border.color": Tier.SOFT,
    "alignment.horizontal": Tier.SOFT,
    "alignment.vertical": Tier.SOFT,
    "alignment.wrap_text": Tier.SOFT,

    # ---- INFO (reported only) ------------------------------------------
    "rich_text_runs": Tier.INFO,
    "row_height": Tier.INFO,
    "tab_color": Tier.INFO,
    "font.underline": Tier.INFO,
    "font.strike": Tier.INFO,
}


@dataclass
class Mismatch:
    """A single point-wise difference between openpyxl and wolfxl."""

    dimension: str
    location: str
    """Human-readable locator, e.g. ``'Sheet1!A1'`` or ``'defined_name:Total'``."""
    openpyxl_value: Any
    wolfxl_value: Any

    @property
    def tier(self) -> Tier:
        return DIMENSION_TIERS.get(self.dimension, Tier.HARD)

    def __str__(self) -> str:  # pragma: no cover - diagnostic only
        return (
            f"[{self.tier.value.upper():4s}] {self.dimension} @ {self.location}: "
            f"openpyxl={self.openpyxl_value!r}  wolfxl={self.wolfxl_value!r}"
        )


@dataclass
class ParityReport:
    """Aggregate of all mismatches for a single fixture (or a test call)."""

    fixture_id: str
    mismatches: list[Mismatch] = field(default_factory=list)

    def record(
        self,
        dimension: str,
        location: str,
        openpyxl_value: Any,
        wolfxl_value: Any,
    ) -> None:
        if _values_equal(openpyxl_value, wolfxl_value):
            return
        self.mismatches.append(
            Mismatch(
                dimension=dimension,
                location=location,
                openpyxl_value=openpyxl_value,
                wolfxl_value=wolfxl_value,
            )
        )

    def hard_failures(self) -> list[Mismatch]:
        return [m for m in self.mismatches if m.tier is Tier.HARD]

    def soft_failures(self) -> list[Mismatch]:
        return [m for m in self.mismatches if m.tier is Tier.SOFT]

    def info_notes(self) -> list[Mismatch]:
        return [m for m in self.mismatches if m.tier is Tier.INFO]

    def summary(self) -> dict[str, int]:
        return {
            "hard": len(self.hard_failures()),
            "soft": len(self.soft_failures()),
            "info": len(self.info_notes()),
        }


def _values_equal(a: Any, b: Any) -> bool:
    """Fuzzy equality for floats, tolerant of None/empty-string ambiguity."""
    if a is None and b == "":
        return True
    if b is None and a == "":
        return True
    if isinstance(a, float) and isinstance(b, float):
        # NaN equality
        if a != a and b != b:
            return True
        return abs(a - b) <= max(1e-9, 1e-9 * max(abs(a), abs(b)))
    if isinstance(a, float) and isinstance(b, int):
        return _values_equal(a, float(b))
    if isinstance(b, float) and isinstance(a, int):
        return _values_equal(float(a), b)
    return a == b
