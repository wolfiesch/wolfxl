"""Array and data-table formula value classes for cell assignments.

RFC-057 — Dynamic-array formulas.

Provides two classes that mirror openpyxl's
``openpyxl.worksheet.formula`` surface:

* :class:`ArrayFormula` — Excel 365 spilled-range formulas
  (``=SEQUENCE(10)`` spilling to A1:A10) and pre-365 array formulas
  (``{=SUM(A1:A10*B1:B10)}``).
* :class:`DataTableFormula` — 1D and 2D Excel data tables created via
  Data > What-If Analysis > Data Table.

These shims intentionally match openpyxl's constructor / equality /
``__repr__`` semantics so user code that does
``cell.value = ArrayFormula("A1:A10", "B1:B10*2")`` Just Works
regardless of which library produced the value.

Pod 1C — Sprint Ο.
"""

from __future__ import annotations

from typing import Any, Optional


class ArrayFormula:
    """Pre-365 array formula (CSE) or Excel 365 spilled dynamic array.

    Constructor signature mirrors openpyxl's ``ArrayFormula(ref, text)``
    so existing user code that does
    ``cell.value = ArrayFormula(ref="A1:A10", text="B1:B10*2")``
    Just Works.

    Attributes:
        ref: Spill / array range, e.g. ``"A1:A10"`` for a single-column
            spill. The cell holding the formula is the *master* of this
            range; every other cell inside ``ref`` reads back as ``None``.
        text: Formula body without the leading ``"="`` and without
            surrounding ``{}`` braces.
    """

    __slots__ = ("ref", "text")

    def __init__(self, ref: str, text: str | None = None) -> None:
        """Create an array-formula value.

        Args:
            ref: Spill or array range, such as ``"A1:A10"``.
            text: Formula text, with or without a leading ``=`` or wrapping
                CSE braces. The stored value is normalized to the bare body.
        """
        self.ref = ref
        # Strip any leading "=" the caller may have passed for
        # convenience — matches openpyxl's coercion.  Also strip
        # surrounding braces so users can paste a CSE formula
        # verbatim from Excel's name box.
        if text is None:
            text = ""
        if text.startswith("{=") and text.endswith("}"):
            text = text[2:-1]
        elif text.startswith("="):
            text = text[1:]
        elif text.startswith("{") and text.endswith("}"):
            text = text[1:-1]
        self.text = text

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, ArrayFormula):
            return NotImplemented
        return self.ref == other.ref and self.text == other.text

    def __hash__(self) -> int:
        return hash((self.ref, self.text))

    def __repr__(self) -> str:
        return f"ArrayFormula(ref={self.ref!r}, text={self.text!r})"


class DataTableFormula:
    """1D or 2D Excel data table formula.

    Constructor signature mirrors openpyxl's ``DataTableFormula``.

    Attributes:
        ref: Range that the data table fills, e.g. ``"B2:F11"``.
        ca: Always-calculate flag (``calcArray``).
        dt2D: Two-variable data-table flag.
        dtr: Row-input flag.
        r1: First input cell.
        r2: Second input cell for 2D tables.
    """

    __slots__ = ("ref", "ca", "dt2D", "dtr", "r1", "r2", "del1", "del2")

    def __init__(
        self,
        ref: str,
        ca: bool = False,
        dt2D: bool = False,
        dtr: bool = False,
        r1: Optional[str] = None,
        r2: Optional[str] = None,
        del1: bool = False,
        del2: bool = False,
        t: Optional[str] = None,
        **kw: Any,
    ) -> None:
        """Create a data-table formula value.

        Args:
            ref: Range that the data table fills.
            ca: Always-calculate flag.
            dt2D: Whether this is a two-variable data table.
            dtr: Whether the data-table input is a row.
            r1: First input cell reference.
            r2: Second input cell reference for two-variable tables.
            del1: Deleted first input-cell flag.
            del2: Deleted second input-cell flag.
            t: Optional OOXML formula-type discriminator. Accepted for
                openpyxl API compatibility and ignored.
            **kw: Additional openpyxl-compatible keyword arguments, accepted
                and ignored for forwards compatibility.
        """
        if t is not None and t != "dataTable":
            raise ValueError(f"Unsupported data-table formula kind: {t!r}")
        self.ref = ref
        self.ca = bool(ca)
        self.dt2D = bool(dt2D)
        self.dtr = bool(dtr)
        self.r1 = r1
        self.r2 = r2
        self.del1 = bool(del1)
        self.del2 = bool(del2)

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, DataTableFormula):
            return NotImplemented
        return (
            self.ref == other.ref
            and self.ca == other.ca
            and self.dt2D == other.dt2D
            and self.dtr == other.dtr
            and self.r1 == other.r1
            and self.r2 == other.r2
            and self.del1 == other.del1
            and self.del2 == other.del2
        )

    def __hash__(self) -> int:
        return hash(
            (self.ref, self.ca, self.dt2D, self.dtr, self.r1, self.r2, self.del1, self.del2)
        )

    def __repr__(self) -> str:
        parts = [f"ref={self.ref!r}"]
        if self.ca:
            parts.append(f"ca={self.ca!r}")
        if self.dt2D:
            parts.append(f"dt2D={self.dt2D!r}")
        if self.dtr:
            parts.append(f"dtr={self.dtr!r}")
        if self.r1 is not None:
            parts.append(f"r1={self.r1!r}")
        if self.r2 is not None:
            parts.append(f"r2={self.r2!r}")
        if self.del1:
            parts.append(f"del1={self.del1!r}")
        if self.del2:
            parts.append(f"del2={self.del2!r}")
        return f"DataTableFormula({', '.join(parts)})"


# ---------------------------------------------------------------------------
# openpyxl-shaped re-exports (RFC-060 Pod 2).
#
# ``openpyxl.cell.cell`` is a kitchen-sink module — user code routinely does
# ``from openpyxl.cell.cell import Cell, MergedCell, WriteOnlyCell``.  Wolfxl
# keeps each class at its own canonical home (``wolfxl._cell.Cell``,
# ``wolfxl.cell._merged.MergedCell``, ``wolfxl.cell._write_only.WriteOnlyCell``)
# and surfaces them all here so a one-line import swap works.
# ---------------------------------------------------------------------------

from wolfxl._cell import Cell  # noqa: E402
from wolfxl.cell._merged import MergedCell  # noqa: E402
from wolfxl.cell._write_only import WriteOnlyCell  # noqa: E402
from wolfxl.cell.rich_text import CellRichText  # noqa: E402
from wolfxl.utils.exceptions import IllegalCharacterError  # noqa: E402
from wolfxl.worksheet.hyperlink import Hyperlink  # noqa: E402


__all__ = [
    "ArrayFormula",
    "Cell",
    "CellRichText",
    "DataTableFormula",
    "Hyperlink",
    "IllegalCharacterError",
    "MergedCell",
    "WriteOnlyCell",
]
