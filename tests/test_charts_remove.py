"""Sprint Ξ (RFC-050) — ``Worksheet.remove_chart`` / ``replace_chart``.

Coverage matrix:

* ``remove_chart`` removes a not-yet-flushed chart from the pending list
  and produces a workbook with no chart parts.
* ``remove_chart`` raises ``ValueError`` on an unknown chart.
* ``replace_chart`` swaps one chart for another in the pending list,
  preserving the anchor and the list position.
* ``replace_chart`` raises ``TypeError`` if the new value isn't a
  :class:`ChartBase` subclass.
* ``replace_chart`` raises ``ValueError`` if the old chart is unknown.
* Anchor inheritance: replacement chart with no anchor of its own
  inherits the anchor from the chart it replaces.
"""

from __future__ import annotations

import zipfile
from pathlib import Path
from typing import Any

import pytest


pytestmark = pytest.mark.skipif(
    not pytest.importorskip("wolfxl", reason="wolfxl not installed"),
    reason="wolfxl required",
)


def _build_workbook_with_pending_chart() -> tuple[Any, Any, Any]:
    """Helper — build a write-mode workbook with one pending chart."""
    import wolfxl
    from wolfxl.chart import BarChart, LineChart, Reference

    wb = wolfxl.Workbook()
    ws = wb.active
    ws.append(["Region", "Q1", "Q2"])
    ws.append(["NA", 100, 110])
    ws.append(["EU", 80, 95])

    bar = BarChart()
    bar.title = "Bar"
    data = Reference(ws, min_col=2, min_row=1, max_col=3, max_row=3)
    bar.add_data(data, titles_from_data=True)
    ws.add_chart(bar, "E2")

    return wb, ws, bar


def test_remove_chart_drops_pending(tmp_path: Path) -> None:
    """Removing a pending chart yields a workbook with no chart parts."""
    wb, ws, bar = _build_workbook_with_pending_chart()
    assert len(ws._pending_charts) == 1

    ws.remove_chart(bar)
    assert len(ws._pending_charts) == 0

    out = tmp_path / "no_charts.xlsx"
    wb.save(out)

    with zipfile.ZipFile(out) as zf:
        names = zf.namelist()
    assert not any(n.startswith("xl/charts/") for n in names), (
        f"expected no chart parts after remove_chart, got {names!r}"
    )


def test_remove_chart_unknown_raises() -> None:
    """Removing a chart that was never added raises ``ValueError``."""
    import wolfxl
    from wolfxl.chart import BarChart

    wb = wolfxl.Workbook()
    ws = wb.active
    ghost = BarChart()
    with pytest.raises(ValueError, match="not added"):
        ws.remove_chart(ghost)


def test_replace_chart_swaps_in_place(tmp_path: Path) -> None:
    """Replacing keeps the chart count, position, and inherited anchor."""
    from wolfxl.chart import LineChart, Reference

    wb, ws, bar = _build_workbook_with_pending_chart()
    assert ws._pending_charts == [bar]
    assert bar._anchor == "E2"

    line = LineChart()
    line.title = "Line"
    data = Reference(ws, min_col=2, min_row=1, max_col=3, max_row=3)
    line.add_data(data, titles_from_data=True)

    ws.replace_chart(bar, line)
    assert ws._pending_charts == [line]
    # Anchor inherited from the chart it replaces.
    assert line._anchor == "E2"

    out = tmp_path / "replaced.xlsx"
    wb.save(out)
    with zipfile.ZipFile(out) as zf:
        chart_names = [n for n in zf.namelist() if n.startswith("xl/charts/")]
        # Exactly one chart part (the replacement, not both).
        chart_xmls = [n for n in chart_names if n.endswith(".xml")]
        assert len(chart_xmls) == 1, chart_names
        # The line chart emits <c:lineChart>, not <c:barChart>.
        xml = zf.read(chart_xmls[0]).decode()
        assert "<c:lineChart>" in xml
        assert "<c:barChart>" not in xml


def test_replace_chart_keeps_explicit_anchor(tmp_path: Path) -> None:
    """If new._anchor is set explicitly, it wins over inheritance."""
    from wolfxl.chart import LineChart, Reference

    wb, ws, bar = _build_workbook_with_pending_chart()

    line = LineChart()
    line._anchor = "G10"
    data = Reference(ws, min_col=2, min_row=1, max_col=3, max_row=3)
    line.add_data(data, titles_from_data=True)

    ws.replace_chart(bar, line)
    assert line._anchor == "G10"


def test_replace_chart_unknown_raises() -> None:
    """Replacing a chart that was never added raises ``ValueError``."""
    import wolfxl
    from wolfxl.chart import BarChart, LineChart

    wb = wolfxl.Workbook()
    ws = wb.active
    ghost_bar = BarChart()
    new_line = LineChart()
    with pytest.raises(ValueError, match="not added"):
        ws.replace_chart(ghost_bar, new_line)


def test_replace_chart_wrong_type_raises() -> None:
    """``replace_chart`` rejects non-ChartBase replacements."""
    wb, ws, bar = _build_workbook_with_pending_chart()
    with pytest.raises(TypeError, match="ChartBase"):
        ws.replace_chart(bar, "not a chart")


def test_remove_chart_then_add_chart_works(tmp_path: Path) -> None:
    """remove_chart followed by a fresh add_chart yields a 1-chart workbook."""
    from wolfxl.chart import LineChart, Reference

    wb, ws, bar = _build_workbook_with_pending_chart()
    ws.remove_chart(bar)

    line = LineChart()
    data = Reference(ws, min_col=2, min_row=1, max_col=3, max_row=3)
    line.add_data(data, titles_from_data=True)
    ws.add_chart(line, "G5")

    out = tmp_path / "remove_then_add.xlsx"
    wb.save(out)
    with zipfile.ZipFile(out) as zf:
        chart_xmls = [
            n for n in zf.namelist()
            if n.startswith("xl/charts/") and n.endswith(".xml")
        ]
        assert len(chart_xmls) == 1
        xml = zf.read(chart_xmls[0]).decode()
        assert "<c:lineChart>" in xml
