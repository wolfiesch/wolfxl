"""Sprint Μ-prime — 3D / Stock / Surface / ProjectedPie chart families.

Pod-β′ ships the Python chart classes (``BarChart3D``, ``LineChart3D``,
``PieChart3D``, ``AreaChart3D``, ``SurfaceChart``, ``SurfaceChart3D``,
``StockChart``, ``ProjectedPieChart``) and Pod-α′ ships the Rust XML
emitter that handles their flat-shape ``to_rust_dict``. This file owns
the **modify-mode round-trip** tests added by Pod-γ′ (RFC-046 §10.12).

If Pod-β′ extends this file with write-mode tests, the two test sets
should coexist — the modify-mode tests below are clearly suffixed
``_modify_mode_round_trip`` so duplication is unlikely.

Each test:
1. Builds a small fixture xlsx with seed data on sheet ``"Data"``.
2. Loads it via ``wolfxl.load_workbook(path, modify=True)``.
3. Calls ``ws.add_chart(<NewFamilyChart>(...), "D2")``.
4. Saves and verifies the chart part is present in the resulting
   workbook (and openpyxl can read the file back).

When neither Pod-α′ nor Pod-β′ have landed, individual tests skip
gracefully (NotImplementedError on construction → pytest.skip).
"""
from __future__ import annotations

import re
import zipfile
from pathlib import Path

import openpyxl
import pytest


pytestmark = (
    pytest.mark.rfc046 if hasattr(pytest.mark, "rfc046") else pytest.mark.usefixtures()
)


@pytest.fixture(autouse=True)
def _force_test_epoch(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("WOLFXL_TEST_EPOCH", "0")


# ---------------------------------------------------------------------------
# Fixture helpers
# ---------------------------------------------------------------------------


def _make_data_fixture(path: Path, sheet_title: str = "Data") -> None:
    """A1:B5 mini table — enough to seed any 2D-style chart."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_title
    for r in range(1, 6):
        ws.cell(r, 1, f"l{r}")
        ws.cell(r, 2, r * 10)
    wb.save(path)


def _make_ohlc_fixture(path: Path, sheet_title: str = "Data") -> None:
    """5×5 OHLC matrix for a StockChart (4 OHLC series).

    Layout::

        | label | open | high | low | close |
        | l1    | 10   | 12   | 9   | 11    |
        | ...   | ...  | ...  | ... | ...   |
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_title
    ws.cell(1, 1, "label")
    ws.cell(1, 2, "open")
    ws.cell(1, 3, "high")
    ws.cell(1, 4, "low")
    ws.cell(1, 5, "close")
    for r in range(2, 6):
        ws.cell(r, 1, f"l{r-1}")
        ws.cell(r, 2, 10 + r)
        ws.cell(r, 3, 12 + r)
        ws.cell(r, 4, 9 + r)
        ws.cell(r, 5, 11 + r)
    wb.save(path)


def _make_surface_grid_fixture(path: Path, sheet_title: str = "Data") -> None:
    """4x4 numeric grid for a SurfaceChart (each column is a series)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_title
    for r in range(1, 5):
        for c in range(1, 5):
            ws.cell(r, c, (r * c) + 1)
    wb.save(path)


def _zip_listing(path: Path) -> list[str]:
    with zipfile.ZipFile(path, "r") as z:
        return sorted(z.namelist())


def _zip_read(path: Path, member: str) -> bytes:
    with zipfile.ZipFile(path, "r") as z:
        return z.read(member)


def _try_construct(chart_cls, *args, **kwargs):
    """Construct ``chart_cls`` or skip if Pod-β′ hasn't landed."""
    try:
        return chart_cls(*args, **kwargs)
    except NotImplementedError as exc:
        pytest.skip(f"{chart_cls.__name__} not yet implemented: {exc}")


# ---------------------------------------------------------------------------
# Modify-mode round-trip — 3D bar / line / pie / area
# ---------------------------------------------------------------------------


def test_bar_chart_3d_modify_mode_round_trip(tmp_path: Path) -> None:
    from wolfxl import load_workbook
    from wolfxl.chart import BarChart3D, Reference

    src = tmp_path / "src.xlsx"
    dst = tmp_path / "dst.xlsx"
    _make_data_fixture(src)
    wb = load_workbook(src, modify=True)
    ws = wb["Data"]
    chart = _try_construct(BarChart3D)
    chart.add_data(
        Reference(ws, min_col=2, min_row=1, max_row=5),
        titles_from_data=False,
    )
    ws.add_chart(chart, "D2")
    wb.save(dst)

    entries = _zip_listing(dst)
    chart_files = [e for e in entries if e.startswith("xl/charts/chart")]
    assert chart_files, f"BarChart3D not emitted; entries={entries}"
    op = openpyxl.load_workbook(dst, data_only=False)
    assert len(op["Data"]._charts) >= 1


def test_line_chart_3d_modify_mode_round_trip(tmp_path: Path) -> None:
    from wolfxl import load_workbook
    from wolfxl.chart import LineChart3D, Reference

    src = tmp_path / "src.xlsx"
    dst = tmp_path / "dst.xlsx"
    _make_data_fixture(src)
    wb = load_workbook(src, modify=True)
    ws = wb["Data"]
    chart = _try_construct(LineChart3D)
    chart.add_data(
        Reference(ws, min_col=2, min_row=1, max_row=5),
        titles_from_data=False,
    )
    ws.add_chart(chart, "D2")
    wb.save(dst)

    chart_files = [
        e for e in _zip_listing(dst) if e.startswith("xl/charts/chart")
    ]
    assert chart_files, "LineChart3D not emitted"


def test_pie_chart_3d_modify_mode_round_trip(tmp_path: Path) -> None:
    from wolfxl import load_workbook
    from wolfxl.chart import PieChart3D, Reference

    src = tmp_path / "src.xlsx"
    dst = tmp_path / "dst.xlsx"
    _make_data_fixture(src)
    wb = load_workbook(src, modify=True)
    ws = wb["Data"]
    chart = _try_construct(PieChart3D)
    chart.add_data(
        Reference(ws, min_col=2, min_row=1, max_row=5),
        titles_from_data=False,
    )
    ws.add_chart(chart, "D2")
    wb.save(dst)

    chart_files = [
        e for e in _zip_listing(dst) if e.startswith("xl/charts/chart")
    ]
    assert chart_files, "PieChart3D not emitted"


def test_area_chart_3d_modify_mode_round_trip(tmp_path: Path) -> None:
    from wolfxl import load_workbook
    from wolfxl.chart import AreaChart3D, Reference

    src = tmp_path / "src.xlsx"
    dst = tmp_path / "dst.xlsx"
    _make_data_fixture(src)
    wb = load_workbook(src, modify=True)
    ws = wb["Data"]
    chart = _try_construct(AreaChart3D)
    chart.add_data(
        Reference(ws, min_col=2, min_row=1, max_row=5),
        titles_from_data=False,
    )
    ws.add_chart(chart, "D2")
    wb.save(dst)

    chart_files = [
        e for e in _zip_listing(dst) if e.startswith("xl/charts/chart")
    ]
    assert chart_files, "AreaChart3D not emitted"


# ---------------------------------------------------------------------------
# Modify-mode round-trip — Surface (2D) and Surface3D
# ---------------------------------------------------------------------------


def test_surface_chart_modify_mode_round_trip(tmp_path: Path) -> None:
    from wolfxl import load_workbook
    from wolfxl.chart import Reference, SurfaceChart

    src = tmp_path / "src.xlsx"
    dst = tmp_path / "dst.xlsx"
    _make_surface_grid_fixture(src)
    wb = load_workbook(src, modify=True)
    ws = wb["Data"]
    chart = _try_construct(SurfaceChart)
    chart.add_data(
        Reference(ws, min_col=1, max_col=4, min_row=1, max_row=4),
        titles_from_data=False,
    )
    ws.add_chart(chart, "F2")
    wb.save(dst)

    chart_files = [
        e for e in _zip_listing(dst) if e.startswith("xl/charts/chart")
    ]
    assert chart_files, "SurfaceChart not emitted"


def test_surface_chart_3d_modify_mode_round_trip(tmp_path: Path) -> None:
    from wolfxl import load_workbook
    from wolfxl.chart import Reference, SurfaceChart3D

    src = tmp_path / "src.xlsx"
    dst = tmp_path / "dst.xlsx"
    _make_surface_grid_fixture(src)
    wb = load_workbook(src, modify=True)
    ws = wb["Data"]
    chart = _try_construct(SurfaceChart3D)
    chart.add_data(
        Reference(ws, min_col=1, max_col=4, min_row=1, max_row=4),
        titles_from_data=False,
    )
    ws.add_chart(chart, "F2")
    wb.save(dst)

    chart_files = [
        e for e in _zip_listing(dst) if e.startswith("xl/charts/chart")
    ]
    assert chart_files, "SurfaceChart3D not emitted"


# ---------------------------------------------------------------------------
# Modify-mode round-trip — Stock chart (4 OHLC series)
# ---------------------------------------------------------------------------


def test_stock_chart_modify_mode_round_trip(tmp_path: Path) -> None:
    from wolfxl import load_workbook
    from wolfxl.chart import Reference, StockChart

    src = tmp_path / "src.xlsx"
    dst = tmp_path / "dst.xlsx"
    _make_ohlc_fixture(src)
    wb = load_workbook(src, modify=True)
    ws = wb["Data"]
    chart = _try_construct(StockChart)
    # Seed all 4 OHLC series (open / high / low / close).
    chart.add_data(
        Reference(ws, min_col=2, max_col=5, min_row=1, max_row=5),
        titles_from_data=True,
    )
    chart.set_categories(
        Reference(ws, min_col=1, max_col=1, min_row=2, max_row=5),
    )
    ws.add_chart(chart, "G2")
    wb.save(dst)

    chart_files = [
        e for e in _zip_listing(dst) if e.startswith("xl/charts/chart")
    ]
    assert chart_files, "StockChart not emitted"
    op = openpyxl.load_workbook(dst, data_only=False)
    assert len(op["Data"]._charts) >= 1


# ---------------------------------------------------------------------------
# Modify-mode round-trip — ProjectedPie (Pie of Pie / Bar of Pie)
# ---------------------------------------------------------------------------


def test_projected_pie_chart_modify_mode_round_trip(tmp_path: Path) -> None:
    from wolfxl import load_workbook
    from wolfxl.chart import ProjectedPieChart, Reference

    src = tmp_path / "src.xlsx"
    dst = tmp_path / "dst.xlsx"
    _make_data_fixture(src)
    wb = load_workbook(src, modify=True)
    ws = wb["Data"]
    chart = _try_construct(ProjectedPieChart)
    chart.add_data(
        Reference(ws, min_col=2, min_row=1, max_row=5),
        titles_from_data=False,
    )
    ws.add_chart(chart, "D2")
    wb.save(dst)

    chart_files = [
        e for e in _zip_listing(dst) if e.startswith("xl/charts/chart")
    ]
    assert chart_files, "ProjectedPieChart not emitted"
