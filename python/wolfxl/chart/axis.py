"""Chart axes — `<c:catAx>`, `<c:valAx>`, `<c:dateAx>`, `<c:serAx>`.

Mirrors :mod:`openpyxl.chart.axis`. Each axis subclass shares the
:class:`_BaseAxis` slot set and adds type-specific extras.

Chart-side axis IDs default to the same constants openpyxl picks
(``catAx`` 10, ``valAx`` 100, ``dateAx`` 500, ``serAx`` 1000) so the
emitted XML matches openpyxl's by default.

Sprint Μ Pod-β (RFC-046).
"""

from __future__ import annotations

from typing import Any

from .data_source import NumFmt
from .layout import Layout
from .shapes import GraphicalProperties
from .text import RichText, Text
from .title import Title, TitleDescriptor


_VALID_AX_POS = ("b", "l", "r", "t")
_VALID_TICK_MARK = (None, "cross", "in", "out", "none")
_VALID_TICK_LBL_POS = (None, "high", "low", "nextTo", "none")
_VALID_CROSSES = (None, "autoZero", "max", "min")
_VALID_TIME_UNIT = (None, "days", "months", "years")


class ChartLines:
    """`<c:majorGridlines>` / `<c:minorGridlines>` — optional spPr-only block."""

    __slots__ = ("spPr",)

    def __init__(self, spPr: GraphicalProperties | None = None) -> None:
        self.spPr = spPr

    @property
    def graphicalProperties(self) -> GraphicalProperties | None:
        return self.spPr

    @graphicalProperties.setter
    def graphicalProperties(self, v: GraphicalProperties | None) -> None:
        self.spPr = v

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {}
        if self.spPr is not None:
            d["spPr"] = self.spPr.to_dict()
        return d


class Scaling:
    """`<c:scaling>` — log base, orientation, manual min/max."""

    __slots__ = ("logBase", "orientation", "max", "min")

    def __init__(
        self,
        logBase: float | None = None,
        orientation: str = "minMax",
        max: float | None = None,
        min: float | None = None,
    ) -> None:
        if orientation not in ("minMax", "maxMin"):
            raise ValueError(f"orientation={orientation!r} must be 'minMax' or 'maxMin'")
        self.logBase = logBase
        self.orientation = orientation
        self.max = max
        self.min = min

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {"orientation": self.orientation}
        if self.logBase is not None:
            d["logBase"] = self.logBase
        if self.max is not None:
            d["max"] = self.max
        if self.min is not None:
            d["min"] = self.min
        return d


class DisplayUnitsLabel:
    """`<c:dispUnitsLbl>` — label for axis display units."""

    __slots__ = ("layout", "tx", "spPr", "txPr")

    def __init__(
        self,
        layout: Layout | None = None,
        tx: Text | None = None,
        spPr: GraphicalProperties | None = None,
        txPr: RichText | None = None,
    ) -> None:
        self.layout = layout
        self.tx = tx
        self.spPr = spPr
        self.txPr = txPr

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {}
        if self.layout is not None:
            d["layout"] = self.layout.to_dict()
        if self.tx is not None:
            d["tx"] = self.tx.to_dict()
        if self.spPr is not None:
            d["spPr"] = self.spPr.to_dict()
        if self.txPr is not None:
            d["txPr"] = self.txPr.to_dict()
        return d


class DisplayUnitsLabelList:
    """`<c:dispUnits>` — display unit selector + label."""

    __slots__ = ("custUnit", "builtInUnit", "dispUnitsLbl")

    _VALID_BUILTIN = (
        None,
        "hundreds",
        "thousands",
        "tenThousands",
        "hundredThousands",
        "millions",
        "tenMillions",
        "hundredMillions",
        "billions",
        "trillions",
    )

    def __init__(
        self,
        custUnit: float | None = None,
        builtInUnit: str | None = None,
        dispUnitsLbl: DisplayUnitsLabel | None = None,
    ) -> None:
        if builtInUnit not in self._VALID_BUILTIN:
            raise ValueError(f"builtInUnit={builtInUnit!r} not in {self._VALID_BUILTIN}")
        self.custUnit = custUnit
        self.builtInUnit = builtInUnit
        self.dispUnitsLbl = dispUnitsLbl

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {}
        if self.custUnit is not None:
            d["custUnit"] = self.custUnit
        if self.builtInUnit is not None:
            d["builtInUnit"] = self.builtInUnit
        if self.dispUnitsLbl is not None:
            d["dispUnitsLbl"] = self.dispUnitsLbl.to_dict()
        return d


class _BaseAxis:
    """Common axis state shared by every axis kind.

    Attributes mirror openpyxl's :class:`_BaseAxis` exactly. ``title``
    accepts either a string (auto-inflated via :class:`TitleDescriptor`)
    or a constructed :class:`Title`.
    """

    title = TitleDescriptor()

    # Per-instance slot list — declared via __init_subclass__ on subclasses
    # via plain attributes. We keep ``__slots__`` empty here so the
    # descriptor's ``_title`` storage on the instance works.

    def __init__(
        self,
        axId: int | None = None,
        scaling: Scaling | None = None,
        delete: bool | None = None,
        axPos: str = "l",
        majorGridlines: ChartLines | None = None,
        minorGridlines: ChartLines | None = None,
        title: Any | None = None,
        numFmt: Any | None = None,
        majorTickMark: str | None = None,
        minorTickMark: str | None = None,
        tickLblPos: str | None = None,
        spPr: GraphicalProperties | None = None,
        txPr: RichText | None = None,
        crossAx: int | None = None,
        crosses: str | None = None,
        crossesAt: float | None = None,
    ) -> None:
        if axPos not in _VALID_AX_POS:
            raise ValueError(f"axPos={axPos!r} not in {_VALID_AX_POS}")
        if majorTickMark not in _VALID_TICK_MARK:
            raise ValueError(f"majorTickMark={majorTickMark!r} not in {_VALID_TICK_MARK}")
        if minorTickMark not in _VALID_TICK_MARK:
            raise ValueError(f"minorTickMark={minorTickMark!r} not in {_VALID_TICK_MARK}")
        if tickLblPos not in _VALID_TICK_LBL_POS:
            raise ValueError(f"tickLblPos={tickLblPos!r} not in {_VALID_TICK_LBL_POS}")
        if crosses not in _VALID_CROSSES:
            raise ValueError(f"crosses={crosses!r} not in {_VALID_CROSSES}")

        self.axId = axId
        self.scaling = scaling if scaling is not None else Scaling()
        self.delete = delete
        self.axPos = axPos
        self.majorGridlines = majorGridlines
        self.minorGridlines = minorGridlines
        self.title = title  # via TitleDescriptor
        self._numFmt: Any | None = None
        self.numFmt = numFmt
        self.majorTickMark = majorTickMark
        self.minorTickMark = minorTickMark
        self.tickLblPos = tickLblPos
        self.spPr = spPr
        self.txPr = txPr
        self.crossAx = crossAx
        self.crosses = crosses
        self.crossesAt = crossesAt

    # numFmt accepts either a NumFmt or a bare format string (openpyxl alias)
    @property
    def numFmt(self) -> NumFmt | None:
        return self._numFmt

    @numFmt.setter
    def numFmt(self, value: Any) -> None:
        if value is None:
            self._numFmt = None
        elif isinstance(value, str):
            self._numFmt = NumFmt(formatCode=value)
        else:
            self._numFmt = value

    @property
    def number_format(self) -> NumFmt | None:
        return self._numFmt

    @number_format.setter
    def number_format(self, value: Any) -> None:
        self.numFmt = value

    @property
    def graphicalProperties(self) -> GraphicalProperties | None:
        return self.spPr

    @graphicalProperties.setter
    def graphicalProperties(self, v: GraphicalProperties | None) -> None:
        self.spPr = v

    @property
    def textProperties(self) -> RichText | None:
        return self.txPr

    @textProperties.setter
    def textProperties(self, v: RichText | None) -> None:
        self.txPr = v

    def _base_to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {
            "axId": self.axId,
            "scaling": self.scaling.to_dict(),
            "axPos": self.axPos,
            "crossAx": self.crossAx,
        }
        if self.delete is not None:
            d["delete"] = self.delete
        if self.majorGridlines is not None:
            d["majorGridlines"] = self.majorGridlines.to_dict()
        if self.minorGridlines is not None:
            d["minorGridlines"] = self.minorGridlines.to_dict()
        if self.title is not None:
            d["title"] = self.title.to_dict()
        if self._numFmt is not None:
            d["numFmt"] = self._numFmt.to_dict()
        if self.majorTickMark is not None:
            d["majorTickMark"] = self.majorTickMark
        if self.minorTickMark is not None:
            d["minorTickMark"] = self.minorTickMark
        if self.tickLblPos is not None:
            d["tickLblPos"] = self.tickLblPos
        if self.spPr is not None:
            d["spPr"] = self.spPr.to_dict()
        if self.txPr is not None:
            d["txPr"] = self.txPr.to_dict()
        if self.crosses is not None:
            d["crosses"] = self.crosses
        if self.crossesAt is not None:
            d["crossesAt"] = self.crossesAt
        return d


class NumericAxis(_BaseAxis):
    """`<c:valAx>` — numeric (value) axis."""

    tagname = "valAx"

    def __init__(
        self,
        crossBetween: str | None = None,
        majorUnit: float | None = None,
        minorUnit: float | None = None,
        dispUnits: DisplayUnitsLabelList | None = None,
        **kw: Any,
    ) -> None:
        if crossBetween is not None and crossBetween not in ("between", "midCat"):
            raise ValueError(f"crossBetween={crossBetween!r} not in (between, midCat)")
        kw.setdefault("majorGridlines", ChartLines())
        kw.setdefault("axId", 100)
        kw.setdefault("crossAx", 10)
        super().__init__(**kw)
        self.crossBetween = crossBetween
        self.majorUnit = majorUnit
        self.minorUnit = minorUnit
        self.dispUnits = dispUnits

    def to_dict(self) -> dict[str, Any]:
        d = self._base_to_dict()
        d["_kind"] = "valAx"
        if self.crossBetween is not None:
            d["crossBetween"] = self.crossBetween
        if self.majorUnit is not None:
            d["majorUnit"] = self.majorUnit
        if self.minorUnit is not None:
            d["minorUnit"] = self.minorUnit
        if self.dispUnits is not None:
            d["dispUnits"] = self.dispUnits.to_dict()
        return d


# openpyxl alias
ValueAxis = NumericAxis
ValAx = NumericAxis


class TextAxis(_BaseAxis):
    """`<c:catAx>` — categorical (text) axis."""

    tagname = "catAx"

    def __init__(
        self,
        auto: bool | None = None,
        lblAlgn: str | None = None,
        lblOffset: int = 100,
        tickLblSkip: int | None = None,
        tickMarkSkip: int | None = None,
        noMultiLvlLbl: bool | None = None,
        **kw: Any,
    ) -> None:
        if lblAlgn is not None and lblAlgn not in ("ctr", "l", "r"):
            raise ValueError(f"lblAlgn={lblAlgn!r} not in (ctr, l, r)")
        if not (0 <= lblOffset <= 1000):
            raise ValueError(f"lblOffset={lblOffset} must be in [0, 1000]")
        kw.setdefault("axId", 10)
        kw.setdefault("crossAx", 100)
        super().__init__(**kw)
        self.auto = auto
        self.lblAlgn = lblAlgn
        self.lblOffset = lblOffset
        self.tickLblSkip = tickLblSkip
        self.tickMarkSkip = tickMarkSkip
        self.noMultiLvlLbl = noMultiLvlLbl

    def to_dict(self) -> dict[str, Any]:
        d = self._base_to_dict()
        d["_kind"] = "catAx"
        d["lblOffset"] = self.lblOffset
        for slot in ("auto", "lblAlgn", "tickLblSkip", "tickMarkSkip", "noMultiLvlLbl"):
            v = getattr(self, slot)
            if v is not None:
                d[slot] = v
        return d


CategoryAxis = TextAxis
CatAx = TextAxis


class DateAxis(TextAxis):
    """`<c:dateAx>` — date axis (subclass of catAx in the spec)."""

    tagname = "dateAx"

    def __init__(
        self,
        auto: bool | None = None,
        lblOffset: int | None = None,
        baseTimeUnit: str | None = None,
        majorUnit: float | None = None,
        majorTimeUnit: str | None = None,
        minorUnit: float | None = None,
        minorTimeUnit: str | None = None,
        **kw: Any,
    ) -> None:
        if baseTimeUnit not in _VALID_TIME_UNIT:
            raise ValueError(f"baseTimeUnit={baseTimeUnit!r} not in {_VALID_TIME_UNIT}")
        if majorTimeUnit not in _VALID_TIME_UNIT:
            raise ValueError(f"majorTimeUnit={majorTimeUnit!r} not in {_VALID_TIME_UNIT}")
        if minorTimeUnit not in _VALID_TIME_UNIT:
            raise ValueError(f"minorTimeUnit={minorTimeUnit!r} not in {_VALID_TIME_UNIT}")
        kw.setdefault("axId", 500)
        # Avoid TextAxis lblOffset bounds check by providing a default.
        if lblOffset is None:
            kw["lblOffset"] = 100
        else:
            kw["lblOffset"] = lblOffset
        super().__init__(**kw)
        # Re-assign post-init since super() set lblOffset to a possibly-default
        self.auto = auto if auto is not None else self.auto
        self.baseTimeUnit = baseTimeUnit
        self.majorUnit = majorUnit
        self.majorTimeUnit = majorTimeUnit
        self.minorUnit = minorUnit
        self.minorTimeUnit = minorTimeUnit

    def to_dict(self) -> dict[str, Any]:
        d = self._base_to_dict()
        d["_kind"] = "dateAx"
        if self.auto is not None:
            d["auto"] = self.auto
        if self.lblOffset is not None:
            d["lblOffset"] = self.lblOffset
        for slot in ("baseTimeUnit", "majorUnit", "majorTimeUnit", "minorUnit", "minorTimeUnit"):
            v = getattr(self, slot)
            if v is not None:
                d[slot] = v
        return d


DateAx = DateAxis


class SeriesAxis(_BaseAxis):
    """`<c:serAx>` — series axis (only used by 3-D charts; we keep it for compat)."""

    tagname = "serAx"

    def __init__(
        self,
        tickLblSkip: int | None = None,
        tickMarkSkip: int | None = None,
        **kw: Any,
    ) -> None:
        kw.setdefault("axId", 1000)
        kw.setdefault("crossAx", 10)
        super().__init__(**kw)
        self.tickLblSkip = tickLblSkip
        self.tickMarkSkip = tickMarkSkip

    def to_dict(self) -> dict[str, Any]:
        d = self._base_to_dict()
        d["_kind"] = "serAx"
        if self.tickLblSkip is not None:
            d["tickLblSkip"] = self.tickLblSkip
        if self.tickMarkSkip is not None:
            d["tickMarkSkip"] = self.tickMarkSkip
        return d


SerAx = SeriesAxis


__all__ = [
    "Axis",
    "CategoryAxis",
    "CatAx",
    "ChartLines",
    "DateAxis",
    "DateAx",
    "DisplayUnitsLabel",
    "DisplayUnitsLabelList",
    "NumericAxis",
    "Scaling",
    "SeriesAxis",
    "SerAx",
    "TextAxis",
    "ValAx",
    "ValueAxis",
    "_BaseAxis",
]


# Public alias matching openpyxl's surface
Axis = _BaseAxis
