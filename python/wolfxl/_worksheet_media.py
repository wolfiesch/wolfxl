"""Worksheet chart, pivot, slicer, and image queue helpers."""

from __future__ import annotations

import re
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from wolfxl._worksheet import Worksheet


def add_chart(ws: Worksheet, chart: Any, anchor: Any = None) -> None:
    """Queue a chart for write-mode or modify-mode save processing."""
    from wolfxl.chart._chart import ChartBase as _ChartBase

    if not isinstance(chart, _ChartBase):
        raise TypeError(
            f"add_chart expected wolfxl.chart.ChartBase, got "
            f"{type(chart).__name__}"
        )

    if anchor is None:
        anchor = chart.anchor if chart.anchor is not None else "E15"

    if isinstance(anchor, str):
        validate_a1_anchor(anchor)

    chart._anchor = anchor  # noqa: SLF001
    ws._pending_charts.append(chart)  # noqa: SLF001
    if ws._charts_cache is not None:  # noqa: SLF001
        ws._charts_cache.append(chart)  # noqa: SLF001


def add_pivot_table(ws: Worksheet, pivot_table: Any) -> None:
    """Queue a pivot table for modify-mode save processing."""
    from wolfxl.pivot import PivotTable as _PivotTable

    if not isinstance(pivot_table, _PivotTable):
        raise TypeError(
            f"add_pivot_table expected wolfxl.pivot.PivotTable, "
            f"got {type(pivot_table).__name__}"
        )
    if ws._workbook._rust_patcher is None:  # noqa: SLF001
        raise RuntimeError(
            "add_pivot_table requires modify mode — open the "
            "workbook with load_workbook(..., modify=True). "
            "Write-mode pivot table emission is not yet supported."
        )
    if pivot_table.cache._cache_id is None:  # noqa: SLF001
        raise ValueError(
            "PivotTable.cache has not been registered with the "
            "workbook yet. Call Workbook.add_pivot_cache(cache) "
            "before Worksheet.add_pivot_table(pt)."
        )
    if hasattr(pivot_table, "_compute_layout"):
        pivot_table._compute_layout()
    ws._pending_pivot_tables.append(pivot_table)  # noqa: SLF001


def add_slicer(ws: Worksheet, slicer: Any, anchor: str) -> None:
    """Queue a slicer presentation for modify-mode save processing."""
    from wolfxl.pivot import Slicer as _Slicer

    if not isinstance(slicer, _Slicer):
        raise TypeError(
            f"add_slicer expected wolfxl.pivot.Slicer, got "
            f"{type(slicer).__name__}"
        )
    if ws._workbook._rust_patcher is None:  # noqa: SLF001
        raise RuntimeError(
            "add_slicer requires modify mode — open the workbook "
            "with load_workbook(..., modify=True)."
        )
    if slicer.cache._slicer_cache_id is None:  # noqa: SLF001
        raise ValueError(
            "Slicer.cache has not been registered with the "
            "workbook yet. Call Workbook.add_slicer_cache(cache) "
            "before Worksheet.add_slicer(slicer, anchor)."
        )
    if not isinstance(anchor, str) or not anchor:
        raise ValueError("Worksheet.add_slicer: anchor must be a non-empty A1 string")
    validate_a1_anchor(anchor)
    slicer.anchor = anchor
    ws._pending_slicers.append(slicer)  # noqa: SLF001


def validate_a1_anchor(anchor: str) -> None:
    """Raise ValueError when *anchor* is not a valid single A1 cell ref."""
    if not anchor:
        raise ValueError("anchor must not be empty")
    match = re.match(r"^([A-Z]+)([0-9]+)$", anchor)
    if not match:
        raise ValueError(
            f"anchor={anchor!r} must be a single A1 cell ref like 'E15' "
            f"(regex ^[A-Z]+[0-9]+$); for ranged or absolute placement "
            f"pass an OneCellAnchor / TwoCellAnchor / AbsoluteAnchor"
        )
    col_letters, row_str = match.group(1), match.group(2)
    col_idx = 0
    for char in col_letters:
        col_idx = col_idx * 26 + (ord(char) - ord("A") + 1)
    if col_idx > 16_384:
        raise ValueError(
            f"anchor={anchor!r}: column {col_letters!r} exceeds Excel max XFD (16384)"
        )
    row_idx = int(row_str)
    if row_idx < 1 or row_idx > 1_048_576:
        raise ValueError(
            f"anchor={anchor!r}: row {row_idx} out of Excel range [1, 1048576]"
        )


def remove_chart(ws: Worksheet, chart: Any) -> None:
    """Remove a not-yet-flushed chart from this worksheet."""
    try:
        ws._pending_charts.remove(chart)  # noqa: SLF001
    except ValueError:
        raise ValueError(
            "chart was not added to this worksheet via add_chart() "
            "(or has already been removed). Removal of charts that "
            "survive from the source workbook is a v1.8 follow-up; "
            "see RFC-050 §6."
        ) from None


def replace_chart(ws: Worksheet, old: Any, new: Any) -> None:
    """Replace one not-yet-flushed chart with another."""
    from wolfxl.chart._chart import ChartBase as _ChartBase

    if not isinstance(new, _ChartBase):
        raise TypeError(
            f"replace_chart expected wolfxl.chart.ChartBase for new, got "
            f"{type(new).__name__}"
        )
    try:
        index = ws._pending_charts.index(old)  # noqa: SLF001
    except ValueError:
        raise ValueError("old chart was not added to this worksheet via add_chart()") from None
    anchor = new._anchor if new._anchor is not None else old._anchor  # noqa: SLF001
    if anchor is None:
        anchor = "E15"
    if isinstance(anchor, str):
        validate_a1_anchor(anchor)
    new._anchor = anchor  # noqa: SLF001
    ws._pending_charts[index] = new  # noqa: SLF001


def get_charts(ws: Worksheet) -> list[Any]:
    """Return charts attached to this worksheet, hydrating read-mode drawings."""
    workbook = ws._workbook  # noqa: SLF001
    reader = getattr(workbook, "_rust_reader", None)
    if reader is None or not hasattr(reader, "read_charts"):
        return ws._pending_charts  # noqa: SLF001

    if ws._charts_cache is None:  # noqa: SLF001
        charts = []
        for payload in reader.read_charts(ws._title):  # noqa: SLF001
            if isinstance(payload, dict):
                chart = _chart_from_payload(payload)
                if chart is not None:
                    charts.append(chart)
        charts.extend(ws._pending_charts)  # noqa: SLF001
        ws._charts_cache = charts  # noqa: SLF001
    return ws._charts_cache  # noqa: SLF001


def _chart_from_payload(payload: dict[str, Any]) -> Any:
    from wolfxl.chart import (
        AreaChart,
        AreaChart3D,
        BarChart,
        BarChart3D,
        BubbleChart,
        DoughnutChart,
        LineChart,
        LineChart3D,
        PieChart,
        PieChart3D,
        ProjectedPieChart,
        RadarChart,
        ScatterChart,
        StockChart,
        SurfaceChart,
        SurfaceChart3D,
    )

    classes = {
        "area": AreaChart,
        "area3d": AreaChart3D,
        "bar": BarChart,
        "bar3d": BarChart3D,
        "bubble": BubbleChart,
        "doughnut": DoughnutChart,
        "line": LineChart,
        "line3d": LineChart3D,
        "pie": PieChart,
        "pie3d": PieChart3D,
        "of_pie": ProjectedPieChart,
        "radar": RadarChart,
        "scatter": ScatterChart,
        "stock": StockChart,
        "surface": SurfaceChart,
        "surface3d": SurfaceChart3D,
    }
    chart_cls = classes.get(str(payload.get("kind") or ""))
    if chart_cls is None:
        return None

    chart = chart_cls()
    if payload.get("title") is not None:
        chart.title = str(payload["title"])
    if payload.get("x_axis_title") is not None and hasattr(chart, "x_axis"):
        chart.x_axis.title = str(payload["x_axis_title"])
    if payload.get("y_axis_title") is not None and hasattr(chart, "y_axis"):
        chart.y_axis.title = str(payload["y_axis_title"])
    if isinstance(payload.get("x_axis"), dict) and hasattr(chart, "x_axis"):
        _hydrate_axis_from_payload(chart.x_axis, payload["x_axis"])
    if isinstance(payload.get("y_axis"), dict) and hasattr(chart, "y_axis"):
        _hydrate_axis_from_payload(chart.y_axis, payload["y_axis"])
    if payload.get("legend_position") is not None and chart.legend is not None:
        chart.legend.position = str(payload["legend_position"])
    if payload.get("bar_dir") is not None and hasattr(chart, "barDir"):
        chart.barDir = str(payload["bar_dir"])
    if payload.get("grouping") is not None and hasattr(chart, "grouping"):
        chart.grouping = str(payload["grouping"])
    if payload.get("scatter_style") is not None and hasattr(chart, "scatterStyle"):
        chart.scatterStyle = str(payload["scatter_style"])
    if payload.get("vary_colors") is not None and hasattr(chart, "varyColors"):
        chart.varyColors = bool(payload["vary_colors"])
    if payload.get("style") is not None:
        chart.style = int(payload["style"])
    chart._anchor = _anchor_from_payload(payload.get("anchor"))  # noqa: SLF001
    chart.ser = [
        _series_from_payload(str(payload.get("kind") or ""), series_payload)
        for series_payload in payload.get("series", [])
        if isinstance(series_payload, dict)
    ]
    return chart


def _hydrate_axis_from_payload(axis: Any, payload: dict[str, Any]) -> None:
    from wolfxl.chart.axis import DisplayUnitsLabelList
    from wolfxl.chart.data_source import NumFmt

    scalar_attrs = {
        "ax_id": "axId",
        "cross_ax": "crossAx",
        "axis_position": "axPos",
        "major_unit": "majorUnit",
        "minor_unit": "minorUnit",
        "tick_lbl_pos": "tickLblPos",
        "major_tick_mark": "majorTickMark",
        "minor_tick_mark": "minorTickMark",
        "crosses": "crosses",
        "crosses_at": "crossesAt",
        "cross_between": "crossBetween",
    }
    for key, attr in scalar_attrs.items():
        if payload.get(key) is not None and hasattr(axis, attr):
            setattr(axis, attr, payload[key])

    if payload.get("num_format_code") is not None:
        axis.numFmt = NumFmt(
            formatCode=str(payload["num_format_code"]),
            sourceLinked=bool(payload.get("num_format_source_linked") or False),
        )

    scaling = getattr(axis, "scaling", None)
    if scaling is not None:
        scaling_attrs = {
            "scaling_min": "min",
            "scaling_max": "max",
            "scaling_orientation": "orientation",
            "scaling_log_base": "logBase",
        }
        for key, attr in scaling_attrs.items():
            if payload.get(key) is not None:
                setattr(scaling, attr, payload[key])

    if payload.get("display_unit") is not None and hasattr(axis, "dispUnits"):
        axis.dispUnits = DisplayUnitsLabelList(builtInUnit=str(payload["display_unit"]))


def _series_from_payload(kind: str, payload: dict[str, Any]) -> Any:
    from wolfxl.chart.data_source import AxDataSource, NumDataSource, NumRef, StrRef
    from wolfxl.chart.series import Series, SeriesLabel, XYSeries

    is_xy = kind in {"scatter", "bubble"}
    series = XYSeries() if is_xy else Series()
    if payload.get("idx") is not None:
        series.idx = int(payload["idx"])
    if payload.get("order") is not None:
        series.order = int(payload["order"])
    if payload.get("title_ref"):
        series.tx = SeriesLabel(strRef=StrRef(str(payload["title_ref"])))
    elif payload.get("title_value"):
        series.tx = SeriesLabel(v=str(payload["title_value"]))

    if is_xy:
        if payload.get("x_ref"):
            series.xVal = AxDataSource(numRef=NumRef(f=str(payload["x_ref"])))
        if payload.get("y_ref"):
            series.yVal = NumDataSource(numRef=NumRef(f=str(payload["y_ref"])))
        if payload.get("bubble_size_ref"):
            series.bubbleSize = NumDataSource(
                numRef=NumRef(f=str(payload["bubble_size_ref"]))
            )
    else:
        if payload.get("cat_ref"):
            series.cat = AxDataSource(strRef=StrRef(str(payload["cat_ref"])))
        if payload.get("val_ref"):
            series.val = NumDataSource(numRef=NumRef(f=str(payload["val_ref"])))
    return series


def add_image(ws: Worksheet, image: Any, anchor: Any = None) -> None:
    """Queue an image for write-mode or modify-mode save processing."""
    from wolfxl.drawing.image import Image as _Image

    if not isinstance(image, _Image):
        raise TypeError(
            f"add_image expected wolfxl.drawing.image.Image, got {type(image).__name__}"
        )

    if anchor is None:
        anchor = "A1"

    image.anchor = anchor
    ws._pending_images.append(image)  # noqa: SLF001
    if ws._images_cache is not None:  # noqa: SLF001
        ws._images_cache.append(image)  # noqa: SLF001


def get_images(ws: Worksheet) -> list[Any]:
    """Return images attached to this worksheet, hydrating read-mode drawings."""
    workbook = ws._workbook  # noqa: SLF001
    reader = getattr(workbook, "_rust_reader", None)
    if reader is None or not hasattr(reader, "read_images"):
        return ws._pending_images  # noqa: SLF001

    if ws._images_cache is None:  # noqa: SLF001
        images = [
            _image_from_payload(payload)
            for payload in reader.read_images(ws._title)  # noqa: SLF001
            if isinstance(payload, dict)
        ]
        images.extend(ws._pending_images)  # noqa: SLF001
        ws._images_cache = images  # noqa: SLF001
    return ws._images_cache  # noqa: SLF001


def _image_from_payload(payload: dict[str, Any]) -> Any:
    from wolfxl.drawing.image import Image

    image = Image(payload["data"])
    image.anchor = _anchor_from_payload(payload.get("anchor"))
    return image


def _anchor_from_payload(payload: Any) -> Any:
    if not isinstance(payload, dict):
        return None

    from wolfxl.drawing.spreadsheet_drawing import (
        AbsoluteAnchor,
        OneCellAnchor,
        TwoCellAnchor,
        XDRPoint2D,
        XDRPositiveSize2D,
    )

    kind = payload.get("type")
    if kind == "one_cell":
        ext = _extent_from_payload(payload)
        return OneCellAnchor(
            _from=_marker_from_payload(payload, "from"),
            ext=ext,
        )
    if kind == "two_cell":
        return TwoCellAnchor(
            _from=_marker_from_payload(payload, "from"),
            to=_marker_from_payload(payload, "to"),
            editAs=str(payload.get("edit_as") or "oneCell"),
        )
    if kind == "absolute":
        return AbsoluteAnchor(
            pos=XDRPoint2D(
                x=int(payload.get("x_emu") or 0),
                y=int(payload.get("y_emu") or 0),
            ),
            ext=_extent_from_payload(payload) or XDRPositiveSize2D(),
        )
    return None


def _marker_from_payload(payload: dict[str, Any], prefix: str) -> Any:
    from wolfxl.drawing.spreadsheet_drawing import AnchorMarker

    return AnchorMarker(
        col=int(payload.get(f"{prefix}_col") or 0),
        row=int(payload.get(f"{prefix}_row") or 0),
        colOff=int(payload.get(f"{prefix}_col_off") or 0),
        rowOff=int(payload.get(f"{prefix}_row_off") or 0),
    )


def _extent_from_payload(payload: dict[str, Any]) -> Any:
    from wolfxl.drawing.spreadsheet_drawing import XDRPositiveSize2D

    cx = payload.get("cx_emu")
    cy = payload.get("cy_emu")
    if cx is None or cy is None:
        return None
    return XDRPositiveSize2D(cx=int(cx), cy=int(cy))
