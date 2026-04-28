"""`<c:errBars>` — series error bars.

Mirrors :class:`openpyxl.chart.error_bar.ErrorBars`.

Sprint Μ Pod-β (RFC-046).
"""

from __future__ import annotations

from typing import Any

from .shapes import GraphicalProperties

_VALID_DIR = (None, "x", "y")
_VALID_BAR_TYPE = ("both", "minus", "plus")
_VALID_VAL_TYPE = ("cust", "fixedVal", "percentage", "stdDev", "stdErr")


class ErrorBars:
    """`<c:errBars>` — direction, magnitude, and source data."""

    __slots__ = (
        "errDir",
        "errBarType",
        "errValType",
        "noEndCap",
        "plus",
        "minus",
        "val",
        "spPr",
    )

    def __init__(
        self,
        errDir: str | None = None,
        errBarType: str = "both",
        errValType: str = "fixedVal",
        noEndCap: bool | None = None,
        plus: Any | None = None,
        minus: Any | None = None,
        val: float | None = None,
        spPr: GraphicalProperties | None = None,
    ) -> None:
        if errDir not in _VALID_DIR:
            raise ValueError(f"errDir={errDir!r} not in {_VALID_DIR}")
        if errBarType not in _VALID_BAR_TYPE:
            raise ValueError(f"errBarType={errBarType!r} not in {_VALID_BAR_TYPE}")
        if errValType not in _VALID_VAL_TYPE:
            raise ValueError(f"errValType={errValType!r} not in {_VALID_VAL_TYPE}")
        self.errDir = errDir
        self.errBarType = errBarType
        self.errValType = errValType
        self.noEndCap = noEndCap
        self.plus = plus
        self.minus = minus
        self.val = val
        self.spPr = spPr

    # openpyxl aliases
    @property
    def direction(self) -> str | None:
        return self.errDir

    @direction.setter
    def direction(self, value: str | None) -> None:
        self.errDir = value

    @property
    def style(self) -> str:
        return self.errBarType

    @style.setter
    def style(self, value: str) -> None:
        self.errBarType = value

    @property
    def size(self) -> str:
        return self.errValType

    @size.setter
    def size(self, value: str) -> None:
        self.errValType = value

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {
            "errBarType": self.errBarType,
            "errValType": self.errValType,
        }
        if self.errDir is not None:
            d["errDir"] = self.errDir
        if self.noEndCap is not None:
            d["noEndCap"] = self.noEndCap
        if self.val is not None:
            d["val"] = self.val
        if self.plus is not None:
            d["plus"] = self.plus.to_dict() if hasattr(self.plus, "to_dict") else self.plus
        if self.minus is not None:
            d["minus"] = self.minus.to_dict() if hasattr(self.minus, "to_dict") else self.minus
        if self.spPr is not None:
            d["spPr"] = self.spPr.to_dict()
        return d


__all__ = ["ErrorBars"]
