"""`<c:ser>` — chart data series + the SeriesFactory helper.

Mirrors :mod:`openpyxl.chart.series` and :mod:`openpyxl.chart.series_factory`.
``XYSeries`` is identical to ``Series`` for our purposes (xVal/yVal/bubbleSize
are already on the parent class); we keep it as a separate name purely so
``isinstance(...)`` checks in client code continue to work.

The ``attribute_mapping`` dict mirrors openpyxl's: it tells each chart type
which slot subset of ``Series`` actually round-trips for its kind.

Sprint Μ Pod-β (RFC-046).
"""

from __future__ import annotations

from typing import Any

from .data_source import (
    AxDataSource,
    NumData,
    NumDataSource,
    NumRef,
    NumVal,
    StrData,
    StrRef,
    StrVal,
)
from .error_bar import ErrorBars
from .label import DataLabelList
from .marker import DataPoint, Marker
from .reference import Reference
from .shapes import GraphicalProperties
from .trendline import Trendline


# Per-chart-type filter for ``Series.__elements__`` — drives the order in
# which the Rust emitter writes child elements for that chart's series.
attribute_mapping = {
    "area": ("idx", "order", "tx", "spPr", "pictureOptions", "dPt", "dLbls",
             "errBars", "trendline", "cat", "val"),
    "bar": ("idx", "order", "tx", "spPr", "invertIfNegative", "pictureOptions",
            "dPt", "dLbls", "trendline", "errBars", "cat", "val", "shape"),
    "bubble": ("idx", "order", "tx", "spPr", "invertIfNegative", "dPt", "dLbls",
               "trendline", "errBars", "xVal", "yVal", "bubbleSize", "bubble3D"),
    "line": ("idx", "order", "tx", "spPr", "marker", "dPt", "dLbls",
             "trendline", "errBars", "cat", "val", "smooth"),
    "pie": ("idx", "order", "tx", "spPr", "explosion", "dPt", "dLbls", "cat", "val"),
    "radar": ("idx", "order", "tx", "spPr", "marker", "dPt", "dLbls", "cat", "val"),
    "scatter": ("idx", "order", "tx", "spPr", "marker", "dPt", "dLbls",
                "trendline", "errBars", "xVal", "yVal", "smooth"),
    "surface": ("idx", "order", "tx", "spPr", "cat", "val"),
}


class SeriesLabel:
    """`<c:tx>` for a series — either a strRef or a literal value."""

    __slots__ = ("strRef", "v")

    def __init__(self, strRef: StrRef | None = None, v: str | None = None) -> None:
        self.strRef = strRef
        self.v = v

    @property
    def value(self) -> str | None:
        return self.v

    @value.setter
    def value(self, val: str | None) -> None:
        self.v = val

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {}
        if self.strRef is not None:
            d["strRef"] = self.strRef.to_dict()
        if self.v is not None:
            d["v"] = self.v
        return d


class Series:
    """A chart data series.

    Carries every slot any chart type might need — :func:`series_factory`
    populates the relevant subset based on whether xvalues / zvalues
    (bubble sizes) are provided. The Rust emitter consults
    :data:`attribute_mapping` to pick which slots become XML.
    """

    __slots__ = (
        "idx",
        "order",
        "tx",
        "spPr",
        "pictureOptions",
        "dPt",
        "dLbls",
        "trendline",
        "errBars",
        "cat",
        "val",
        "invertIfNegative",
        "shape",
        "xVal",
        "yVal",
        "bubbleSize",
        "bubble3D",
        "marker",
        "smooth",
        "explosion",
    )

    def __init__(
        self,
        idx: int = 0,
        order: int = 0,
        tx: SeriesLabel | None = None,
        spPr: GraphicalProperties | None = None,
        pictureOptions: Any | None = None,
        dPt: list[DataPoint] | tuple[DataPoint, ...] = (),
        dLbls: DataLabelList | None = None,
        trendline: Trendline | None = None,
        errBars: ErrorBars | None = None,
        cat: AxDataSource | None = None,
        val: NumDataSource | None = None,
        invertIfNegative: bool | None = None,
        shape: str | None = None,
        xVal: AxDataSource | None = None,
        yVal: NumDataSource | None = None,
        bubbleSize: NumDataSource | None = None,
        bubble3D: bool | None = None,
        marker: Marker | None = None,
        smooth: bool | None = None,
        explosion: int | None = None,
    ) -> None:
        self.idx = idx
        self.order = order
        self.tx = tx
        self.spPr = spPr if spPr is not None else GraphicalProperties()
        self.pictureOptions = pictureOptions
        self.dPt = list(dPt)
        self.dLbls = dLbls
        self.trendline = trendline
        self.errBars = errBars
        self.cat = cat
        self.val = val
        self.invertIfNegative = invertIfNegative
        self.shape = shape
        self.xVal = xVal
        self.yVal = yVal
        self.bubbleSize = bubbleSize
        self.bubble3D = bubble3D
        self.marker = marker if marker is not None else Marker()
        self.smooth = smooth
        self.explosion = explosion

    # openpyxl aliases
    @property
    def title(self) -> SeriesLabel | None:
        return self.tx

    @title.setter
    def title(self, value: SeriesLabel | None) -> None:
        self.tx = value

    @property
    def graphicalProperties(self) -> GraphicalProperties | None:
        return self.spPr

    @graphicalProperties.setter
    def graphicalProperties(self, value: GraphicalProperties | None) -> None:
        self.spPr = value

    @property
    def data_points(self) -> list[DataPoint]:
        return self.dPt

    @data_points.setter
    def data_points(self, value: list[DataPoint]) -> None:
        self.dPt = list(value)

    @property
    def labels(self) -> DataLabelList | None:
        return self.dLbls

    @labels.setter
    def labels(self, value: DataLabelList | None) -> None:
        self.dLbls = value

    @property
    def identifiers(self) -> AxDataSource | None:
        return self.cat

    @identifiers.setter
    def identifiers(self, value: AxDataSource | None) -> None:
        self.cat = value

    @property
    def zVal(self) -> NumDataSource | None:
        return self.bubbleSize

    @zVal.setter
    def zVal(self, value: NumDataSource | None) -> None:
        self.bubbleSize = value

    def to_rust_dict(self, series_type: str = "bar") -> dict[str, Any]:
        """Serialise this series for a chart of the given ``series_type``.

        Only the slots in ``attribute_mapping[series_type]`` are emitted —
        matches the per-type ``__elements__`` filtering openpyxl applies
        on ``to_tree``.
        """
        keys = attribute_mapping.get(series_type, attribute_mapping["bar"])
        d: dict[str, Any] = {}
        for key in keys:
            v = getattr(self, key, None)
            if v is None:
                continue
            if isinstance(v, (list, tuple)):
                if not v:
                    continue
                d[key] = [
                    item.to_dict() if hasattr(item, "to_dict") else item for item in v
                ]
            elif hasattr(v, "to_dict"):
                d[key] = v.to_dict()
            else:
                d[key] = v
        # idx + order are always emitted even if zero — they're identifiers,
        # not optional metadata.
        d.setdefault("idx", self.idx)
        d.setdefault("order", self.order)
        return d


class XYSeries(Series):
    """Series for chart types with explicit X/Y (Scatter, Bubble).

    Identical storage to :class:`Series`; the dedicated subclass exists so
    callers can ``isinstance(s, XYSeries)`` to gate xVal/yVal-only logic.
    """


def series_factory(
    values: Any,
    xvalues: Any | None = None,
    zvalues: Any | None = None,
    title: str | SeriesLabel | None = None,
    title_from_data: bool = False,
) -> Series:
    """Factory that wraps cell ranges into a :class:`Series` of the right shape.

    Mirrors :func:`openpyxl.chart.series_factory.SeriesFactory`. The
    ``title_from_data`` flag pops the first cell of ``values`` and uses
    its sheet-qualified address as a strRef title (so a column header
    becomes the legend label).
    """
    if not isinstance(values, Reference):
        values = Reference(range_string=values)

    series_title: SeriesLabel | None = None
    if title_from_data:
        cell = values.pop()
        title_str = f"{values.sheetname}!{cell}"
        series_title = SeriesLabel(strRef=StrRef(title_str))
    elif title is not None:
        if isinstance(title, SeriesLabel):
            series_title = title
        else:
            series_title = SeriesLabel(v=str(title))

    val = NumDataSource(numRef=NumRef(f=values))

    if xvalues is not None:
        if not isinstance(xvalues, Reference):
            xvalues = Reference(range_string=xvalues)
        series: Series = XYSeries()
        series.yVal = val
        series.xVal = AxDataSource(numRef=NumRef(f=xvalues))
        if zvalues is not None:
            if not isinstance(zvalues, Reference):
                zvalues = Reference(range_string=zvalues)
            series.bubbleSize = NumDataSource(numRef=NumRef(f=zvalues))
    else:
        series = Series()
        series.val = val

    if series_title is not None:
        series.tx = series_title

    return series


# openpyxl-style camelCase alias
SeriesFactory = series_factory


__all__ = [
    "Series",
    "SeriesFactory",
    "SeriesLabel",
    "XYSeries",
    "attribute_mapping",
    "series_factory",
]
