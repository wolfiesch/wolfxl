"""`ChartBase` — common ancestor for every chart type.

Mirrors :class:`openpyxl.chart._chart.ChartBase`. Every chart subclass
inherits ``title``, ``legend``, ``layout``, ``style``, ``display_blanks``,
``visible_cells_only``, ``roundedCorners``, ``graphical_properties``, plus
the ``add_data``/``set_categories``/``append`` helpers that drive the
fluent construction pattern.

The :meth:`to_rust_dict` method is the contract surface Pod-α's PyO3
binding consumes. It returns a flat-shape dict matching RFC-046 §10.

Sprint Μ-prime Pod-β′ (RFC-046 §10) — v1.6.1 contract.
"""

from __future__ import annotations

import re
from collections import OrderedDict
from operator import attrgetter
from typing import Any

from .data_source import AxDataSource, NumRef
from .layout import Layout
from .legend import Legend
from .reference import Reference
from .series import Series, SeriesFactory
from .shapes import GraphicalProperties
from .title import TitleDescriptor


_VALID_DISPLAY_BLANKS = ("span", "gap", "zero")

# RFC-049 §10.2 — `pivot_source.name` regex. Optional sheet prefix +
# table name. Sheet prefix only allows the conservative identifier set;
# table-name segment additionally allows spaces (Excel pivot names like
# "PivotTable 1" are commonplace in the wild).
_PIVOT_SOURCE_NAME_RE = re.compile(
    r"^([A-Za-z_][A-Za-z0-9_]*!)?[A-Za-z_][A-Za-z0-9_ ]*$"
)
_PIVOT_FMT_ID_MAX = 65535


# Map openpyxl tagname (`barChart`, `bar3DChart`, …) to the §10.2 short
# kind string. Both 2D and 3D variants — the 8 new families ship in
# Sprint Μ-prime (v1.6.1).
_TAGNAME_TO_KIND = {
    "barChart": "bar",
    "bar3DChart": "bar3d",
    "lineChart": "line",
    "line3DChart": "line3d",
    "pieChart": "pie",
    "pie3DChart": "pie3d",
    "ofPieChart": "of_pie",
    "doughnutChart": "doughnut",
    "areaChart": "area",
    "area3DChart": "area3d",
    "scatterChart": "scatter",
    "bubbleChart": "bubble",
    "radarChart": "radar",
    "surfaceChart": "surface",
    "surface3DChart": "surface3d",
    "stockChart": "stock",
}


class ChartBase:
    """Base class for all chart kinds.

    Subclasses provide:
    - ``tagname`` — XML root for this chart's plot block (e.g. ``"barChart"``).
    - ``_series_type`` — key into :data:`series.attribute_mapping`.
    - ``_chart_type_specific_keys()`` — flat dict of per-type keys merged
      into the §10.1 top-level shape (replaces v1.6.0 ``_chart_dict_extras``).
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
        # Sprint Ν Pod-δ (RFC-049). Internal storage for the snake-case
        # ``pivot_source`` attribute is a dict shaped like the §10.1
        # contract (or ``None``). The legacy ``pivotSource`` openpyxl
        # alias (typed as ``Any``) is preserved for back-compat with
        # callers that imported the openpyxl PivotSource class
        # directly; it does not flow through ``to_rust_dict``.
        self._pivot_source: dict[str, Any] | None = None
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
    # ``pivot_source`` — Sprint Ν Pod-δ (RFC-049 §10).
    #
    # Linking a chart to a pivot table is what makes Excel render it as
    # a "pivot chart" (right-click → Refresh, pivot-aware toolbar, etc.).
    # The OOXML serialization is a top-of-`<c:chart>` ``<c:pivotSource>``
    # block + an extra ``<c:fmtId val="0"/>`` on every series. Pod-δ's
    # Rust emitter handles both; this attribute is the Python surface.
    # ------------------------------------------------------------------
    @property
    def pivot_source(self) -> dict[str, Any] | None:
        """The chart's pivot-source linkage as a §10.1-shaped dict, or
        ``None`` if unlinked.
        """
        return self._pivot_source

    @pivot_source.setter
    def pivot_source(self, value: Any) -> None:
        if value is None:
            self._pivot_source = None
            return

        # Tuple form: (name, fmt_id).
        if isinstance(value, tuple):
            if len(value) != 2:
                raise ValueError(
                    "Chart.pivot_source tuple must be (name, fmt_id); "
                    f"got tuple of length {len(value)}"
                )
            name, fmt_id = value
            self._pivot_source = self._validate_pivot_source(name, fmt_id)
            return

        # Dict form (round-tripped from to_rust_dict).
        if isinstance(value, dict):
            if "name" not in value:
                raise ValueError(
                    "Chart.pivot_source dict must include 'name'"
                )
            name = value["name"]
            fmt_id = value.get("fmt_id", 0)
            self._pivot_source = self._validate_pivot_source(name, fmt_id)
            return

        # Duck-typed PivotTable. We avoid a hard import to dodge the
        # circular wolfxl.pivot._table → wolfxl.chart import cycle. Any
        # object that quacks with a ``.name`` string attribute and isn't
        # a primitive is treated as a pivot table.
        name_attr = getattr(value, "name", None)
        if isinstance(name_attr, str):
            self._pivot_source = self._validate_pivot_source(name_attr, 0)
            return

        raise TypeError(
            "Chart.pivot_source accepts a PivotTable, (name, fmt_id) "
            "tuple, dict {'name': str, 'fmt_id': int}, or None; "
            f"got {type(value).__name__}"
        )

    @staticmethod
    def _validate_pivot_source(name: Any, fmt_id: Any) -> dict[str, Any]:
        """RFC-049 §10.2 validation. Returns a normalised dict."""
        if not isinstance(name, str) or not name:
            raise ValueError(
                "pivot_source.name must be a non-empty string"
            )
        if not _PIVOT_SOURCE_NAME_RE.match(name):
            raise ValueError(
                f"pivot_source.name={name!r} does not match the OOXML "
                f"pivot-source name regex"
            )
        try:
            fmt_id_int = int(fmt_id)
        except (TypeError, ValueError) as exc:
            raise ValueError(
                f"pivot_source.fmt_id must be an int, got {fmt_id!r}"
            ) from exc
        if not (0 <= fmt_id_int <= _PIVOT_FMT_ID_MAX):
            raise ValueError(
                f"pivot_source.fmt_id={fmt_id_int} must be in "
                f"[0, {_PIVOT_FMT_ID_MAX}]"
            )
        return {"name": name, "fmt_id": fmt_id_int}

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
            s.idx = idx
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
    # Validation
    # ------------------------------------------------------------------
    def _validate_at_emit(self) -> None:
        """Per RFC-046 §10.11: raise at ``to_rust_dict`` time on bad state."""
        if not self.ser:
            raise ValueError(
                f"{type(self).__name__} requires at least one series "
                f"(call chart.add_data(...) before saving)."
            )

    # ------------------------------------------------------------------
    # Rust-side serialisation
    # ------------------------------------------------------------------
    def _chart_type_specific_keys(self) -> dict[str, Any]:
        """Return per-type flat keys to merge into the §10.1 top-level dict.

        Replaces the v1.6.0 ``_chart_dict_extras`` envelope. Keys are
        snake_case, flat at top-level (no nesting). Subclasses override.
        """
        return {}

    # v1.6.0 hook kept as a no-op default for any subclass that hasn't
    # been migrated yet — never called by ``to_rust_dict`` itself.
    def _chart_dict_extras(self) -> dict[str, Any]:  # pragma: no cover
        return {}

    def to_rust_dict(self) -> dict[str, Any]:
        """Produce a typed dict describing this chart for the Rust emitter.

        Shape: see RFC-046 §10.1 (flat top-level keys, snake_case
        throughout; no ``axes`` list, no ``extras`` envelope).
        """
        self._validate_at_emit()
        self._reindex()

        kind = _TAGNAME_TO_KIND.get(self.tagname)
        if kind is None:
            raise ValueError(
                f"Unknown chart tagname={self.tagname!r}; expected one of "
                f"{sorted(_TAGNAME_TO_KIND)}"
            )

        x_axis = getattr(self, "x_axis", None)
        y_axis = getattr(self, "y_axis", None)
        z_axis = getattr(self, "z_axis", None)

        series_list = [s.to_rust_dict(self._series_type) for s in self.ser]

        d: dict[str, Any] = {
            # Required
            "kind": kind,
            "series_type": self._series_type,
            "series": series_list,

            # Optional shared
            "style": self._style,
            "display_blanks_as": self._display_blanks,
            "plot_visible_only": self.visible_cells_only,
            "rounded_corners": self.roundedCorners,

            # Decorations
            "title": self.title.to_dict() if self.title is not None else None,
            "legend": self.legend.to_dict() if self.legend is not None else None,
            "layout": self.layout.to_dict() if self.layout is not None else None,
            "graphical_properties": (
                self.graphical_properties.to_dict()
                if self.graphical_properties is not None
                else None
            ),

            # Axes — flat keys (NOT a list)
            "x_axis": x_axis.to_dict() if x_axis is not None else None,
            "y_axis": y_axis.to_dict() if y_axis is not None else None,
            "z_axis": z_axis.to_dict() if z_axis is not None else None,

            # Anchor + dimensions
            "anchor": self._anchor,
            "width_emu": int(self.width * 360_000) if self.width is not None else None,
            "height_emu": int(self.height * 360_000) if self.height is not None else None,

            # Sprint Ν Pod-δ — RFC-049 §10.1. ``None`` → no
            # ``<c:pivotSource>`` block emitted; chart is a standard
            # chart. Dict shape `{"name": str, "fmt_id": int}`.
            "pivot_source": self._pivot_source,
        }

        # Default vary_colors (None → omit; subclasses may set)
        vc = getattr(self, "vary_colors", None)
        if vc is not None:
            d["vary_colors"] = vc

        # Merge in per-type flat keys (snake_case at top level).
        d.update(self._chart_type_specific_keys())
        return d


__all__ = ["ChartBase"]
