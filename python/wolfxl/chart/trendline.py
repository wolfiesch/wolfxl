"""`<c:trendline>` — series trendlines.

Mirrors :class:`openpyxl.chart.trendline.Trendline`.

Sprint Μ Pod-β (RFC-046).
"""

from __future__ import annotations

from typing import Any

from .layout import Layout
from .shapes import GraphicalProperties
from .text import RichText, Text


_VALID_TRENDLINE_TYPES = ("exp", "linear", "log", "movingAvg", "poly", "power")


class TrendlineLabel:
    """`<c:trendlineLbl>` — display label attached to a trendline."""

    __slots__ = ("layout", "tx", "numFmt", "spPr", "txPr")

    def __init__(
        self,
        layout: Layout | None = None,
        tx: Text | None = None,
        numFmt: Any | None = None,
        spPr: GraphicalProperties | None = None,
        txPr: RichText | None = None,
    ) -> None:
        self.layout = layout
        self.tx = tx
        self.numFmt = numFmt
        self.spPr = spPr
        self.txPr = txPr

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {}
        if self.layout is not None:
            d["layout"] = self.layout.to_dict()
        if self.tx is not None:
            d["tx"] = self.tx.to_dict()
        if self.numFmt is not None:
            if hasattr(self.numFmt, "to_dict"):
                d["numFmt"] = self.numFmt.to_dict()
            else:
                d["numFmt"] = self.numFmt
        if self.spPr is not None:
            d["spPr"] = self.spPr.to_dict()
        if self.txPr is not None:
            d["txPr"] = self.txPr.to_dict()
        return d


class Trendline:
    """`<c:trendline>` — trendline kind, parameters, and display options."""

    __slots__ = (
        "name",
        "spPr",
        "trendlineType",
        "order",
        "period",
        "forward",
        "backward",
        "intercept",
        "dispRSqr",
        "dispEq",
        "trendlineLbl",
    )

    def __init__(
        self,
        name: str | None = None,
        spPr: GraphicalProperties | None = None,
        trendlineType: str = "linear",
        order: int | None = None,
        period: int | None = None,
        forward: float | None = None,
        backward: float | None = None,
        intercept: float | None = None,
        dispRSqr: bool | None = None,
        dispEq: bool | None = None,
        trendlineLbl: TrendlineLabel | None = None,
    ) -> None:
        if trendlineType not in _VALID_TRENDLINE_TYPES:
            raise ValueError(
                f"trendlineType={trendlineType!r} not in {_VALID_TRENDLINE_TYPES}"
            )
        self.name = name
        self.spPr = spPr
        self.trendlineType = trendlineType
        self.order = order
        self.period = period
        self.forward = forward
        self.backward = backward
        self.intercept = intercept
        self.dispRSqr = dispRSqr
        self.dispEq = dispEq
        self.trendlineLbl = trendlineLbl

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {"trendlineType": self.trendlineType}
        if self.name is not None:
            d["name"] = self.name
        if self.spPr is not None:
            d["spPr"] = self.spPr.to_dict()
        if self.order is not None:
            d["order"] = self.order
        if self.period is not None:
            d["period"] = self.period
        if self.forward is not None:
            d["forward"] = self.forward
        if self.backward is not None:
            d["backward"] = self.backward
        if self.intercept is not None:
            d["intercept"] = self.intercept
        if self.dispRSqr is not None:
            d["dispRSqr"] = self.dispRSqr
        if self.dispEq is not None:
            d["dispEq"] = self.dispEq
        if self.trendlineLbl is not None:
            d["trendlineLbl"] = self.trendlineLbl.to_dict()
        return d


__all__ = ["Trendline", "TrendlineLabel"]
