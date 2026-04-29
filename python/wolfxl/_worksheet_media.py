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
