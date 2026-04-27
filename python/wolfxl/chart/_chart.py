"""`ChartBase` — common ancestor for every chart type.

Mirrors :class:`openpyxl.chart._chart.ChartBase`. Every chart subclass
inherits ``title``, ``legend``, ``layout``, ``style``, ``display_blanks``,
``visible_cells_only``, ``roundedCorners``, ``graphical_properties``, plus
the ``add_data``/``set_categories``/``append`` helpers that drive the
fluent construction pattern.

The :meth:`to_rust_dict` method is the contract surface Pod-α's PyO3
binding consumes. It returns a typed-dict-shaped payload describing the
chart's kind, series, axes, and decorations. Each subclass overrides
:meth:`_chart_dict_extras` to layer in chart-type-specific fields.

Sprint Μ Pod-β (RFC-046).
"""

from __future__ import annotations

from collections import OrderedDict
from operator import attrgetter
from typing import Any

from .data_source import AxDataSource, NumRef, StrRef
from .layout import Layout
from .legend import Legend
from .reference import Reference
from .series import Series, SeriesFactory, SeriesLabel
from .shapes import GraphicalProperties
from .title import Title, TitleDescriptor


_VALID_DISPLAY_BLANKS = ("span", "gap", "zero")


class ChartBase:
    """Base class for all chart kinds.

    Subclasses provide:
    - ``tagname`` — XML root for this chart's plot block (e.g. ``"barChart"``).
    - ``_series_type`` — key into :data:`series.attribute_mapping`.
    - Per-type properties (``barDir``, ``grouping``, ``smooth``, …).

    The class is a plain attribute carrier — no descriptor heavy lifting
    beyond the ``TitleDescriptor`` for ``title``.
    """

    title = TitleDescriptor()

    # Default tagname / series-type — subclasses override.
    tagname: str = "chart"
    _series_type: str = ""

    # Anchor + dimension defaults match openpyxl
    anchor: str | Any = "E15"
    width: float = 15.0  # cm
    height: float = 7.5  # cm
    mime_type: str = (
        "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
    )

    def __init__(self, axId: tuple[int, ...] = (), **kw: Any) -> None:
        # Per-instance state (kept off __slots__ so subclasses can add).
        self._charts: list[ChartBase] = [self]
        self.title = None  # via TitleDescriptor; user can later assign str/Title
        self.layout: Layout | None = None
        self.roundedCorners: bool | None = None
        self.legend: Legend | None = Legend()
        self.graphical_properties: GraphicalProperties | None = None
        self._style: int | None = None
        self.axId = tuple(axId) if axId else ()
        self.display_blanks: str = "gap"
        self.pivotSource: Any | None = None
        self.pivotFormats: tuple[Any, ...] = ()
        self.visible_cells_only: bool = True
        self.idx_base: int = 0

        # ser is the storage; ``series`` is the openpyxl alias.
        self.ser: list[Series] = []

        # Per-type defaults can be set in subclass __init__ before super().__init__.
        # We swallow remaining kwargs that subclasses already bound on themselves.
        for key, val in kw.items():
            if not hasattr(self, key):
                setattr(self, key, val)

        # Anchor for placement on the worksheet — set by ``ws.add_chart``.
        self._anchor: Any = None

    # ------------------------------------------------------------------
    # ``style`` is bounded 1..48 in the spec.
    # ------------------------------------------------------------------
    @property
    def style(self) -> int | None:
        return self._style

    @style.setter
    def style(self, value: int | None) -> None:
        if value is None:
            self._style = None
            return
        v = int(value)
        if not (1 <= v <= 48):
            raise ValueError(f"style={v} must be in [1, 48]")
        self._style = v

    # ------------------------------------------------------------------
    # ``display_blanks`` is bounded to ('span', 'gap', 'zero').
    # ------------------------------------------------------------------
    @property
    def display_blanks(self) -> str:
        return self._display_blanks

    @display_blanks.setter
    def display_blanks(self, value: str) -> None:
        if value not in _VALID_DISPLAY_BLANKS:
            raise ValueError(f"display_blanks={value!r} not in {_VALID_DISPLAY_BLANKS}")
        self._display_blanks = value

    # ------------------------------------------------------------------
    # openpyxl alias: ``series`` is read/write on top of ``ser``.
    # ------------------------------------------------------------------
    @property
    def series(self) -> list[Series]:
        return self.ser

    @series.setter
    def series(self, value: list[Series]) -> None:
        self.ser = list(value)

    # openpyxl alias used by some legacy code paths
    @property
    def dataLabels(self) -> Any:
        return getattr(self, "dLbls", None)

    @dataLabels.setter
    def dataLabels(self, value: Any) -> None:
        self.dLbls = value

    # ------------------------------------------------------------------
    # Composition helpers
    # ------------------------------------------------------------------
    def __hash__(self) -> int:  # noqa: D401
        return id(self)

    def __iadd__(self, other: "ChartBase") -> "ChartBase":
        if not isinstance(other, ChartBase):
            raise TypeError("Only other charts can be added")
        self._charts.append(other)
        return self

    def append(self, value: Series) -> None:
        """Append a single :class:`Series` to ``self.series``."""
        self.ser = list(self.ser) + [value]

    def add_data(
        self,
        data: Any,
        from_rows: bool = False,
        titles_from_data: bool = False,
    ) -> None:
        """Add a range of data as one or more series.

        If ``from_rows`` is True, each row of *data* becomes a series;
        otherwise (default) each column becomes a series. ``titles_from_data``
        consumes the first cell of each row/column as the legend label
        (matching openpyxl exactly).
        """
        if not isinstance(data, Reference):
            data = Reference(range_string=data)
        values = data.rows if from_rows else data.cols
        for ref in values:
            series = SeriesFactory(ref, title_from_data=titles_from_data)
            self.ser.append(series)

    def set_categories(self, labels: Any) -> None:
        """Set the categories (x-axis labels) for every series."""
        if not isinstance(labels, Reference):
            labels = Reference(range_string=labels)
        cat = AxDataSource(numRef=NumRef(f=labels))
        for s in self.ser:
            s.cat = cat

    def _reindex(self) -> None:
        """Sort series by ``order`` and rebase indexes (matches openpyxl)."""
        ds = sorted(self.ser, key=attrgetter("order"))
        for idx, s in enumerate(ds):
            s.order = idx
        self.ser = ds

    @property
    def _axes(self) -> "OrderedDict[int, Any]":
        x = getattr(self, "x_axis", None)
        y = getattr(self, "y_axis", None)
        z = getattr(self, "z_axis", None)
        return OrderedDict(
            [(axis.axId, axis) for axis in (x, y, z) if axis is not None]
        )

    # ------------------------------------------------------------------
    # Rust-side serialisation
    # ------------------------------------------------------------------
    def _chart_dict_extras(self) -> dict[str, Any]:
        """Hook for subclasses — extra type-specific keys for ``to_rust_dict``."""
        return {}

    def to_rust_dict(self) -> dict[str, Any]:
        """Produce a typed dict describing this chart for the Rust emitter.

        Schema (top-level keys):

        * ``kind`` — short string matching ``self.tagname`` (``"barChart"``,
          ``"lineChart"`` …).
        * ``series_type`` — :data:`series.attribute_mapping` key.
        * ``style`` — int 1..48 or None.
        * ``display_blanks`` — ``"span"`` | ``"gap"`` | ``"zero"``.
        * ``visible_cells_only`` — bool.
        * ``rounded_corners`` — bool | None.
        * ``title`` — see :meth:`Title.to_dict` or None.
        * ``legend`` — see :meth:`Legend.to_dict` or None.
        * ``layout`` — see :meth:`Layout.to_dict` or None.
        * ``graphical_properties`` — chart-wide spPr, or None.
        * ``axes`` — list of ``{kind, axId, crossAx, ...}`` dicts.
        * ``series`` — list of dicts (slots filtered by ``series_type``).
        * ``extras`` — chart-type-specific dict (gapWidth, smooth, ...).
        * ``anchor`` — A1 string the ws.add_chart call resolved (or None).
        """
        axes_list: list[dict[str, Any]] = []
        for axis in (
            getattr(self, "x_axis", None),
            getattr(self, "y_axis", None),
            getattr(self, "z_axis", None),
        ):
            if axis is not None:
                axes_list.append(axis.to_dict())

        series_list = [s.to_rust_dict(self._series_type) for s in self.ser]

        # Map openpyxl tagname (e.g. "barChart") to Pod-α's short kind
        # ("bar"). 3D variants like "bar3DChart" should not reach
        # to_rust_dict() in v1.6.0 — those are deferred to v1.6.1 and
        # raise NotImplementedError at the class constructor — but if
        # they slip through, we let Pod-α surface "unknown chart kind".
        _TAGNAME_TO_KIND = {
            "barChart": "bar",
            "lineChart": "line",
            "pieChart": "pie",
            "doughnutChart": "doughnut",
            "areaChart": "area",
            "scatterChart": "scatter",
            "bubbleChart": "bubble",
            "radarChart": "radar",
        }

        d: dict[str, Any] = {
            "kind": _TAGNAME_TO_KIND.get(self.tagname, self.tagname),
            "series_type": self._series_type,
            "style": self._style,
            "display_blanks": self._display_blanks,
            "visible_cells_only": self.visible_cells_only,
            "rounded_corners": self.roundedCorners,
            "title": self.title.to_dict() if self.title is not None else None,
            "legend": self.legend.to_dict() if self.legend is not None else None,
            "layout": self.layout.to_dict() if self.layout is not None else None,
            "graphical_properties": (
                self.graphical_properties.to_dict()
                if self.graphical_properties is not None
                else None
            ),
            "axes": axes_list,
            "series": series_list,
            "extras": self._chart_dict_extras(),
            "anchor": self._anchor,
            "width": self.width,
            "height": self.height,
        }
        return d


__all__ = ["ChartBase"]
