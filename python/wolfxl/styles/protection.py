"""``openpyxl.styles.protection`` — cell-protection value type."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass
class Protection:
    """Cell protection flags.

    Matches openpyxl's lightweight ``Protection(locked=True, hidden=False)``
    construction surface. Sheet-level enforcement still lives in
    ``wolfxl.worksheet.protection.SheetProtection``.
    """

    locked: bool = True
    hidden: bool = False

    def to_rust_dict(self) -> dict[str, bool]:
        return {"locked": self.locked, "hidden": self.hidden}

__all__ = ["Protection"]
