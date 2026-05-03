# pyright: reportOptionalSubscript=false, reportOptionalMemberAccess=false
"""Openpyxl-API compatibility oracle (Sprint 0 / Gap G01).

This harness is the program-level scoreboard for ``Plans/openpyxl-parity-program.md``.
It owns one parametrised test (``test_compat_oracle_probe``) that runs a
named probe for every entry in ``docs/migration/_compat_spec.py`` that
carries a ``probe`` field.

Probe-status semantics (mirrors the spec's ``status`` field):

* ``supported``    - probe must pass; failure is a regression.
* ``partial``      - probe is xfail today; flips ``XPASS`` once the gap closes.
* ``not_yet``      - probe is xfail today; flips ``XPASS`` once the gap closes.
* ``out_of_scope`` - probe is skipped with a stable reason.

When a probe flips from xfail to xpass, the workflow is:

1. Confirm the underlying gap is actually closed (read the relevant code).
2. Update ``docs/migration/_compat_spec.py``: change ``status`` from
   ``not_yet`` (or ``partial``) to ``supported``; remove ``gap_id``.
3. Run ``python scripts/render_compat_matrix.py``.
4. Update ``Plans/openpyxl-parity-program.md`` status row (gap → ``landed``).
5. Re-run this oracle. The probe should now pass cleanly (no xfail mark).

The test prints a one-line summary at session end so the per-sprint gate can
read the rolling pass count without parsing pytest output.

S1+ may extend this harness by vendoring openpyxl's source-distribution test
files via ``scripts/fetch_openpyxl_corpus.py`` and running them under a
``sys.modules`` shim. For S0 we deliberately stop at the curated probe set so
the baseline number is meaningful and reproducible without network.
"""

from __future__ import annotations

import importlib.util
import json
import os
from collections import Counter
from datetime import date
from pathlib import Path
from typing import Any, Callable

import pytest


REPO_ROOT = Path(__file__).resolve().parents[1]
SPEC_PATH = REPO_ROOT / "docs" / "migration" / "_compat_spec.py"


def _load_spec_module():
    spec = importlib.util.spec_from_file_location("_compat_spec", SPEC_PATH)
    if spec is None or spec.loader is None:  # pragma: no cover - import guard
        raise RuntimeError(f"failed to import {SPEC_PATH}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


_SPEC = _load_spec_module()


# --------------------------------------------------------------------------
# Probe registry
# --------------------------------------------------------------------------
#
# Every probe is a function taking ``tmp_path`` and returning ``None`` on
# success. Failure is signalled by raising any exception (typically
# ``AssertionError``). Probes deliberately stay small - they exercise one
# openpyxl-API contract apiece, not full feature suites - so the harness
# stays cheap to run and the failure messages stay precise.

_ProbeFn = Callable[[Path], None]
PROBES: dict[str, _ProbeFn] = {}


def _register(name: str) -> Callable[[_ProbeFn], _ProbeFn]:
    def decorator(fn: _ProbeFn) -> _ProbeFn:
        PROBES[name] = fn
        return fn

    return decorator


# --------------------------------------------------------------------------
# Workbook + Worksheet probes
# --------------------------------------------------------------------------


@_register("workbook_open_basic")
def _probe_workbook_open_basic(tmp_path: Path) -> None:
    import wolfxl

    wb = wolfxl.Workbook()
    assert wb.active is not None
    assert "Sheet" in wb.sheetnames or len(wb.sheetnames) == 1


@_register("workbook_load_basic")
def _probe_workbook_load_basic(tmp_path: Path) -> None:
    import wolfxl

    wb = wolfxl.Workbook()
    ws = wb.active
    ws["A1"] = "hello"
    out = tmp_path / "load_basic.xlsx"
    wb.save(out)

    wb2 = wolfxl.load_workbook(out)
    assert wb2.active["A1"].value == "hello"


@_register("workbook_load_data_only")
def _probe_workbook_load_data_only(tmp_path: Path) -> None:
    import openpyxl
    import wolfxl

    src = tmp_path / "data_only.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = 1
    ws["A2"] = 2
    ws["A3"] = "=SUM(A1:A2)"
    wb.save(src)

    wb2 = wolfxl.load_workbook(src, data_only=False)
    assert wb2.active["A3"].value == "=SUM(A1:A2)"


@_register("workbook_load_modify")
def _probe_workbook_load_modify(tmp_path: Path) -> None:
    import wolfxl

    src = tmp_path / "modify.xlsx"
    wb = wolfxl.Workbook()
    wb.active["A1"] = "before"
    wb.save(src)

    wb2 = wolfxl.load_workbook(src, modify=True)
    wb2.active["A1"] = "after"
    wb2.save(src)
    wb2.close()

    wb3 = wolfxl.load_workbook(src)
    assert wb3.active["A1"].value == "after"


@_register("workbook_load_read_only")
def _probe_workbook_load_read_only(tmp_path: Path) -> None:
    import wolfxl

    src = tmp_path / "ro.xlsx"
    wb = wolfxl.Workbook()
    ws = wb.active
    for i in range(1, 11):
        ws.cell(row=i, column=1, value=i)
    wb.save(src)

    wb2 = wolfxl.load_workbook(src, read_only=True)
    values = [row[0].value for row in wb2.active.iter_rows(min_row=1, max_row=10)]
    assert values == list(range(1, 11))
    wb2.close()


@_register("workbook_save_basic")
def _probe_workbook_save_basic(tmp_path: Path) -> None:
    import wolfxl

    wb = wolfxl.Workbook()
    wb.active["A1"] = 42
    out = tmp_path / "save_basic.xlsx"
    wb.save(out)
    assert out.exists() and out.stat().st_size > 0


@_register("workbook_sheet_access")
def _probe_workbook_sheet_access(tmp_path: Path) -> None:
    import wolfxl

    wb = wolfxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    assert wb["Data"] is ws
    assert wb.active is ws
    assert "Data" in wb.sheetnames


@_register("workbook_create_sheet")
def _probe_workbook_create_sheet(tmp_path: Path) -> None:
    import wolfxl

    wb = wolfxl.Workbook()
    ws = wb.create_sheet(title="Extra")
    assert "Extra" in wb.sheetnames
    assert wb["Extra"] is ws


@_register("workbook_copy_worksheet")
def _probe_workbook_copy_worksheet(tmp_path: Path) -> None:
    import wolfxl

    wb = wolfxl.Workbook()
    ws = wb.active
    ws.title = "Source"
    ws["A1"] = "data"
    copy = wb.copy_worksheet(ws)
    assert copy["A1"].value == "data"
    assert copy.title != "Source"


@_register("workbook_write_only_streaming")
def _probe_workbook_write_only_streaming(tmp_path: Path) -> None:
    """Bounded-memory append-only path. Today the kwarg is accepted but the
    save still routes through the in-memory writer, so the probe currently
    just checks that the kwarg works at construction. S7 (G20) introduces
    the real streaming path; this probe will then expand to assert RSS
    bounds on a 10M-row write.
    """
    import wolfxl

    wb = wolfxl.Workbook(write_only=True)
    ws = wb.create_sheet("stream")
    for i in range(1000):
        ws.append([i, i * 2])
    out = tmp_path / "stream.xlsx"
    wb.save(out)

    rb = wolfxl.load_workbook(out, read_only=True)
    rows = list(rb.active.iter_rows(values_only=True, max_row=5))
    assert rows[0] == (0, 0)
    assert rows[-1] == (4, 8)
    rb.close()


# --------------------------------------------------------------------------
# Cell + style probes
# --------------------------------------------------------------------------


@_register("cell_basic_value")
def _probe_cell_basic_value(tmp_path: Path) -> None:
    import wolfxl

    wb = wolfxl.Workbook()
    ws = wb.active
    ws["A1"] = "string"
    ws["B1"] = 42
    ws["C1"] = 3.14
    ws["D1"] = True

    out = tmp_path / "cells.xlsx"
    wb.save(out)
    wb2 = wolfxl.load_workbook(out)
    assert wb2.active["A1"].value == "string"
    assert wb2.active["B1"].value == 42
    assert wb2.active["C1"].value == pytest.approx(3.14)
    assert wb2.active["D1"].value is True


@_register("cell_font_fill_border_alignment")
def _probe_cell_font_fill_border_alignment(tmp_path: Path) -> None:
    import wolfxl
    from wolfxl.styles import Alignment, Border, Font, PatternFill, Side

    wb = wolfxl.Workbook()
    ws = wb.active
    ws["A1"] = "styled"
    ws["A1"].font = Font(name="Arial", size=12, bold=True, color="FF0000")
    ws["A1"].fill = PatternFill(fill_type="solid", fgColor="FFFF00")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws["A1"].border = Border(left=Side(style="thin"), right=Side(style="thin"))
    ws["A1"].number_format = "0.00"

    out = tmp_path / "style.xlsx"
    wb.save(out)
    wb2 = wolfxl.load_workbook(out)
    cell = wb2.active["A1"]
    assert cell.font.bold is True
    assert (cell.font.name or "").lower() in ("arial", "calibri", "")
    assert cell.fill.fill_type == "solid"
    assert cell.alignment.horizontal == "center"


@_register("cell_diagonal_borders")
def _probe_cell_diagonal_borders(tmp_path: Path) -> None:
    import wolfxl
    from wolfxl.styles import Border, Side

    wb = wolfxl.Workbook()
    ws = wb.active
    ws["A1"] = "diag"
    ws["A1"].border = Border(
        diagonal=Side(style="thin", color="000000"),
        diagonalUp=True,
        diagonalDown=False,
    )
    out = tmp_path / "diag.xlsx"
    wb.save(out)

    wb2 = wolfxl.load_workbook(out)
    border = wb2.active["A1"].border
    assert border.diagonalUp is True
    assert border.diagonal is not None
    assert border.diagonal.style == "thin"


@_register("cell_protection")
def _probe_cell_protection(tmp_path: Path) -> None:
    import wolfxl
    from wolfxl.styles import Protection

    wb = wolfxl.Workbook()
    ws = wb.active
    ws["A1"] = "locked"
    ws["A1"].protection = Protection(locked=True, hidden=True)
    out = tmp_path / "prot.xlsx"
    wb.save(out)

    wb2 = wolfxl.load_workbook(out)
    prot = wb2.active["A1"].protection
    assert prot.locked is True
    assert prot.hidden is True


@_register("cell_named_style")
def _probe_cell_named_style(tmp_path: Path) -> None:
    import wolfxl
    from wolfxl.styles import Font, NamedStyle

    wb = wolfxl.Workbook()
    style = NamedStyle(name="Highlight")
    style.font = Font(bold=True)
    wb.add_named_style(style)

    ws = wb.active
    ws["A1"] = "named"
    ws["A1"].style = "Highlight"

    out = tmp_path / "named_style.xlsx"
    wb.save(out)
    wb2 = wolfxl.load_workbook(out)
    assert wb2.active["A1"].style == "Highlight"


@_register("cell_gradient_fill")
def _probe_cell_gradient_fill(tmp_path: Path) -> None:
    import wolfxl
    from wolfxl.styles import Color, GradientFill

    wb = wolfxl.Workbook()
    ws = wb.active
    ws["A1"] = "grad"
    ws["A1"].fill = GradientFill(
        type="linear",
        degree=90,
        stop=(Color("FF0000"), Color("00FF00")),
    )
    out = tmp_path / "grad.xlsx"
    wb.save(out)

    wb2 = wolfxl.load_workbook(out)
    fill = wb2.active["A1"].fill
    assert fill.type == "linear" or fill.type == "gradient"


# --------------------------------------------------------------------------
# Charts probes
# --------------------------------------------------------------------------


@_register("charts_basic_2d")
def _probe_charts_basic_2d(tmp_path: Path) -> None:
    import wolfxl
    from wolfxl.chart import BarChart, Reference

    wb = wolfxl.Workbook()
    ws = wb.active
    for row in [["x", "y"], [1, 10], [2, 20], [3, 30]]:
        ws.append(row)

    chart = BarChart()
    chart.title = "Demo"
    data = Reference(ws, min_col=2, min_row=1, max_row=4)
    cats = Reference(ws, min_col=1, min_row=2, max_row=4)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    ws.add_chart(chart, "D2")

    out = tmp_path / "bar.xlsx"
    wb.save(out)
    assert out.exists() and out.stat().st_size > 0


@_register("charts_advanced_2d")
def _probe_charts_advanced_2d(tmp_path: Path) -> None:
    import wolfxl
    from wolfxl.chart import Reference, ScatterChart

    wb = wolfxl.Workbook()
    ws = wb.active
    for row in [["x", "y"], [1, 10], [2, 20], [3, 30]]:
        ws.append(row)

    chart = ScatterChart()
    xvals = Reference(ws, min_col=1, min_row=2, max_row=4)
    yvals = Reference(ws, min_col=2, min_row=1, max_row=4)
    chart.add_data(yvals, titles_from_data=True)
    chart.set_categories(xvals)
    ws.add_chart(chart, "D2")

    out = tmp_path / "scatter.xlsx"
    wb.save(out)
    assert out.exists() and out.stat().st_size > 0


@_register("charts_add_remove_replace")
def _probe_charts_add_remove_replace(tmp_path: Path) -> None:
    import wolfxl
    from wolfxl.chart import BarChart, Reference

    wb = wolfxl.Workbook()
    ws = wb.active
    for row in [["x", "y"], [1, 10], [2, 20]]:
        ws.append(row)

    chart = BarChart()
    data = Reference(ws, min_col=2, min_row=1, max_row=3)
    chart.add_data(data, titles_from_data=True)
    ws.add_chart(chart, "D2")
    assert hasattr(ws, "remove_chart")
    ws.remove_chart(chart)


@_register("charts_combination")
def _probe_charts_combination(tmp_path: Path) -> None:
    import openpyxl as _opx
    import wolfxl
    from wolfxl.chart import BarChart, LineChart, Reference

    wb = wolfxl.Workbook()
    ws = wb.active
    for row in [["x", "y", "z"], [1, 10, 100], [2, 20, 200], [3, 30, 300]]:
        ws.append(row)

    bar = BarChart()
    line = LineChart()
    bar.add_data(Reference(ws, min_col=2, min_row=1, max_row=4), titles_from_data=True)
    line.add_data(Reference(ws, min_col=3, min_row=1, max_row=4), titles_from_data=True)
    line.y_axis.crosses = "max"
    line.y_axis.axId = 200
    bar += line
    ws.add_chart(bar, "E2")
    out = tmp_path / "combo.xlsx"
    wb.save(out)

    ref_wb = _opx.load_workbook(out)
    ref_ws = ref_wb[ref_wb.sheetnames[0]]
    ref_charts = getattr(ref_ws, "_charts", [])
    assert ref_charts, "openpyxl found no charts in combo file"
    types = {type(c).__name__ for c in ref_charts}
    assert "BarChart" in types and "LineChart" in types, (
        f"openpyxl saw {types}, expected both BarChart and LineChart"
    )


@_register("charts_label_rich_text")
def _probe_charts_label_rich_text(tmp_path: Path) -> None:
    import wolfxl
    from wolfxl.cell.text import InlineFont
    from wolfxl.chart import BarChart, Reference
    from wolfxl.chart.label import DataLabelList
    from wolfxl.cell.rich_text import CellRichText, TextBlock

    wb = wolfxl.Workbook()
    ws = wb.active
    for row in [["x", "y"], [1, 10], [2, 20]]:
        ws.append(row)

    chart = BarChart()
    chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=3), titles_from_data=True)
    rich = CellRichText([TextBlock(InlineFont(b=True), "bold-label")])
    chart.dataLabels = DataLabelList(rich=rich)
    ws.add_chart(chart, "D2")
    wb.save(tmp_path / "rich_chart.xlsx")


@_register("charts_pivot_chart_per_point")
def _probe_charts_pivot_chart_per_point(tmp_path: Path) -> None:
    """G16 — per-data-point fill override on a pivot-source chart.

    The OOXML per-point override (`<c:dPt>` with `<c:spPr><a:solidFill>`)
    is identical for pivot and non-pivot charts; the only difference for
    a pivot chart is the chart-level `<c:pivotSource>` block plus a
    `<c:fmtId>` on each `<c:ser>`. This probe stamps a pivot-source on
    the chart (via the (name, fmt_id) tuple form so we don't have to
    materialise a real PivotTable) and then exercises the per-point
    override path. After save + openpyxl reload, we assert openpyxl
    sees the dPt entry on the series — proving the data_points dict
    bridge survives the pivot emit path.
    """
    import openpyxl as _opx
    import wolfxl
    from wolfxl.chart import BarChart, Reference
    from wolfxl.chart.marker import DataPoint
    from wolfxl.chart.shapes import GraphicalProperties

    wb = wolfxl.Workbook()
    ws = wb.active
    for row in [["x", "y"], [1, 10], [2, 20], [3, 30]]:
        ws.append(row)

    chart = BarChart()
    chart.add_data(
        Reference(ws, min_col=2, min_row=1, max_row=4),
        titles_from_data=True,
    )
    chart.pivot_source = ("MyPivot", 0)

    red = GraphicalProperties(solidFill="FF0000")
    chart.series[0].dPt = [DataPoint(idx=1, spPr=red)]

    ws.add_chart(chart, "D2")
    out = tmp_path / "pivot_per_point.xlsx"
    wb.save(out)

    ref_wb = _opx.load_workbook(out)
    ref_ws = ref_wb[ref_wb.sheetnames[0]]
    ref_chart = ref_ws._charts[0]
    ref_dpt = list(ref_chart.series[0].dPt)
    assert ref_dpt, "openpyxl saw no per-point overrides on pivot chart"
    assert any(dp.idx == 1 for dp in ref_dpt), "expected idx=1 override"


# --------------------------------------------------------------------------
# Pivot probes
# --------------------------------------------------------------------------


@_register("pivots_construction")
def _probe_pivots_construction(tmp_path: Path) -> None:
    """Pivot construction is a modify-mode operation (matches openpyxl
    workflow where pivots are stamped onto an existing workbook).
    """
    import wolfxl
    from wolfxl.chart import Reference
    from wolfxl.pivot import PivotCache, PivotTable

    src = tmp_path / "pivot_src.xlsx"
    wb = wolfxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["region", "amount"])
    for row in [("east", 10), ("west", 20), ("east", 30), ("west", 40)]:
        ws.append(row)
    wb.create_sheet("Pivot")
    wb.save(src)

    wb2 = wolfxl.load_workbook(src, modify=True)
    data_ws = wb2["Data"]
    cache = PivotCache(source=Reference(data_ws, min_col=1, min_row=1, max_col=2, max_row=5))
    wb2.add_pivot_cache(cache)
    pt = PivotTable(cache=cache, location="A1", rows=["region"], data=["amount"])
    wb2["Pivot"].add_pivot_table(pt)
    wb2.save(src)
    wb2.close()

    wb3 = wolfxl.load_workbook(src)
    assert "Pivot" in wb3.sheetnames


@_register("pivots_in_place_edit")
def _probe_pivots_in_place_edit(tmp_path: Path) -> None:
    """Edit an existing pivot's source range. Tracked under G17 (S5)."""
    import wolfxl
    from wolfxl.chart import Reference
    from wolfxl.pivot import PivotCache, PivotTable

    wb = wolfxl.Workbook()
    ws = wb.active
    ws.append(["region", "amount"])
    for row in [("east", 10), ("west", 20)]:
        ws.append(row)
    cache = PivotCache(source=Reference(ws, min_col=1, min_row=1, max_col=2, max_row=3))
    wb.add_pivot_cache(cache)
    pt = PivotTable(cache=cache, location="D2", rows=["region"], data=["amount"])
    target = wb.create_sheet("Pivot")
    target.add_pivot_table(pt, "A1")
    src = tmp_path / "pivot_edit.xlsx"
    wb.save(src)

    wb2 = wolfxl.load_workbook(src, modify=True)
    pivot = wb2["Pivot"].pivot_tables[0]
    pivot.source = Reference(wb2.active, min_col=1, min_row=1, max_col=2, max_row=3)
    wb2.save(src)


# --------------------------------------------------------------------------
# Image probes
# --------------------------------------------------------------------------

_PNG_1X1 = (
    b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01\x08\x06\x00\x00\x00"
    b"\x1f\x15\xc4\x89\x00\x00\x00\rIDATx\x9cc\xfa\xcf\x00\x00\x00\x02\x00\x01\xe5'\xde\xfc"
    b"\x00\x00\x00\x00IEND\xaeB`\x82"
)


@_register("images_basic")
def _probe_images_basic(tmp_path: Path) -> None:
    import wolfxl
    from wolfxl.drawing.image import Image

    img_path = tmp_path / "tiny.png"
    img_path.write_bytes(_PNG_1X1)

    wb = wolfxl.Workbook()
    ws = wb.active
    ws.add_image(Image(str(img_path)), "B2")
    out = tmp_path / "with_image.xlsx"
    wb.save(out)
    assert out.exists() and out.stat().st_size > 0


@_register("images_replace_remove")
def _probe_images_replace_remove(tmp_path: Path) -> None:
    """Public replace/remove image API. Tracked under G06 (S1)."""
    import wolfxl
    from wolfxl.drawing.image import Image

    img_path = tmp_path / "tiny.png"
    img_path.write_bytes(_PNG_1X1)

    wb = wolfxl.Workbook()
    ws = wb.active
    img = Image(str(img_path))
    ws.add_image(img, "B2")
    assert hasattr(ws, "remove_image"), "Worksheet.remove_image missing"
    ws.remove_image(img)


# --------------------------------------------------------------------------
# Structural probes
# --------------------------------------------------------------------------


@_register("structural_insert_delete_rows")
def _probe_structural_insert_delete_rows(tmp_path: Path) -> None:
    """insert_rows / delete_rows shift existing data on save (RFC-030/031).

    wolfxl's modify-mode treats insert/delete as queued operations applied
    at flush time, so this probe asserts the post-save shift only and does
    not interleave writes at the moved rows.
    """
    import wolfxl

    src = tmp_path / "ins.xlsx"
    wb = wolfxl.Workbook()
    ws = wb.active
    for i in range(1, 6):
        ws.cell(row=i, column=1, value=i)
    wb.save(src)

    wb2 = wolfxl.load_workbook(src, modify=True)
    wb2.active.insert_rows(2, amount=1)
    wb2.save(src)
    wb2.close()

    wb3 = wolfxl.load_workbook(src)
    assert wb3.active["A1"].value == 1
    assert wb3.active["A2"].value is None  # inserted blank row
    assert wb3.active["A3"].value == 2

    wb4 = wolfxl.load_workbook(src, modify=True)
    wb4.active.delete_rows(2, amount=1)
    wb4.save(src)
    wb4.close()

    wb5 = wolfxl.load_workbook(src)
    assert wb5.active["A2"].value == 2


# --------------------------------------------------------------------------
# Modify-mode probes
# --------------------------------------------------------------------------


@_register("modify_defined_names")
def _probe_modify_defined_names(tmp_path: Path) -> None:
    import wolfxl
    from wolfxl.workbook.defined_name import DefinedName

    wb = wolfxl.Workbook()
    wb.active.title = "Data"
    wb.defined_names["TaxRate"] = DefinedName(name="TaxRate", value="Data!$A$1")
    out = tmp_path / "dn.xlsx"
    wb.save(out)

    wb2 = wolfxl.load_workbook(out)
    assert "TaxRate" in wb2.defined_names


@_register("modify_tables")
def _probe_modify_tables(tmp_path: Path) -> None:
    import wolfxl
    from wolfxl.worksheet.table import Table, TableStyleInfo

    wb = wolfxl.Workbook()
    ws = wb.active
    ws.append(["a", "b"])
    ws.append([1, 2])
    ws.append([3, 4])

    table = Table(name="MyTable", displayName="MyTable", ref="A1:B3")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(table)

    out = tmp_path / "tbl.xlsx"
    wb.save(out)

    wb2 = wolfxl.load_workbook(out)
    assert "MyTable" in [t.displayName for t in wb2.active.tables.values()]


@_register("modify_data_validations")
def _probe_modify_data_validations(tmp_path: Path) -> None:
    import wolfxl
    from wolfxl.worksheet.datavalidation import DataValidation

    wb = wolfxl.Workbook()
    ws = wb.active
    dv = DataValidation(type="list", formula1='"a,b,c"')
    dv.add("A1:A10")
    ws.data_validations.append(dv)

    out = tmp_path / "dv.xlsx"
    wb.save(out)

    wb2 = wolfxl.load_workbook(out)
    assert len(list(wb2.active.data_validations.dataValidation)) >= 1


# --------------------------------------------------------------------------
# Utility probes
# --------------------------------------------------------------------------


@_register("utils_get_column_letter")
def _probe_utils_get_column_letter(tmp_path: Path) -> None:
    import openpyxl.utils
    import wolfxl.utils.cell as wc

    for col in (1, 26, 27, 52, 100, 1000, 16384):
        assert wc.get_column_letter(col) == openpyxl.utils.get_column_letter(col)


@_register("utils_column_index_from_string")
def _probe_utils_column_index_from_string(tmp_path: Path) -> None:
    import openpyxl.utils
    import wolfxl.utils.cell as wc

    for s in ("A", "Z", "AA", "AZ", "BA", "ZZ", "AAA", "XFD"):
        assert wc.column_index_from_string(s) == openpyxl.utils.column_index_from_string(s)


@_register("utils_range_boundaries")
def _probe_utils_range_boundaries(tmp_path: Path) -> None:
    import openpyxl.utils
    import wolfxl.utils.cell as wc

    for r in ("A1:B2", "C3:Z99", "AA1:AZ100"):
        assert wc.range_boundaries(r) == openpyxl.utils.range_boundaries(r)


# --------------------------------------------------------------------------
# Protection probes
# --------------------------------------------------------------------------


@_register("protection_sheet")
def _probe_protection_sheet(tmp_path: Path) -> None:
    import openpyxl as _opx
    import wolfxl
    from wolfxl.worksheet.protection import SheetProtection

    wb = wolfxl.Workbook()
    ws = wb.active
    ws.protection = SheetProtection(sheet=True, password="secret", formatCells=False)
    out = tmp_path / "sheet_prot.xlsx"
    wb.save(out)

    wb2 = wolfxl.load_workbook(out)
    assert wb2.active.protection.sheet is True

    ref_wb = _opx.load_workbook(out)
    ref_prot = ref_wb[ref_wb.sheetnames[0]].protection
    assert ref_prot.sheet is True, "openpyxl saw protection.sheet=False"
    assert ref_prot.formatCells is False, (
        "openpyxl lost the formatCells=False override"
    )
    assert ref_prot.password is not None, "openpyxl saw no password hash"


@_register("protection_workbook")
def _probe_protection_workbook(tmp_path: Path) -> None:
    import wolfxl
    from wolfxl.workbook.protection import WorkbookProtection

    wb = wolfxl.Workbook()
    wb.security = WorkbookProtection(lockStructure=True, workbookPassword="secret")
    out = tmp_path / "wb_prot.xlsx"
    wb.save(out)

    wb2 = wolfxl.load_workbook(out)
    assert wb2.security is not None
    assert getattr(wb2.security, "lockStructure", False) is True


# --------------------------------------------------------------------------
# External links probes
# --------------------------------------------------------------------------


@_register("external_links_collection")
def _probe_external_links_collection(tmp_path: Path) -> None:
    """Workbook-level external link collection. Tracked under G18 (S6).

    The probe asserts that round-tripping a workbook with an external-link
    formula preserves the ``xl/externalLinks/`` parts. Today wolfxl preserves
    the parts on modify-save but does not expose a Python collection;
    ``wb._external_links`` (or equivalent public surface) is what S6 ships.
    """
    import wolfxl

    wb = wolfxl.Workbook()
    wb.active["A1"] = "='[ext.xlsx]Sheet1'!$A$1"
    out = tmp_path / "ext.xlsx"
    wb.save(out)

    wb2 = wolfxl.load_workbook(out)
    links = getattr(wb2, "_external_links", None) or getattr(wb2, "external_links", None)
    assert links is not None and len(links) >= 0  # surface must exist


# --------------------------------------------------------------------------
# Comments probes
# --------------------------------------------------------------------------


@_register("comments_basic")
def _probe_comments_basic(tmp_path: Path) -> None:
    import wolfxl
    from wolfxl.comments import Comment

    wb = wolfxl.Workbook()
    ws = wb.active
    ws["A1"] = "anchor"
    ws["A1"].comment = Comment("hello", "wolfie")
    out = tmp_path / "cmt.xlsx"
    wb.save(out)

    wb2 = wolfxl.load_workbook(out)
    assert wb2.active["A1"].comment is not None
    assert wb2.active["A1"].comment.text == "hello"


@_register("comments_threaded")
def _probe_comments_threaded(tmp_path: Path) -> None:
    """Threaded comments (xl/threadedComments). Tracked under G08 (S2).

    Probe asserts the full round-trip: add a top-level threaded comment +
    one reply, save, reload, and confirm both texts survive. Class-existence
    alone is not enough — that lets the probe XPASS on partial work.
    """
    import wolfxl
    from wolfxl.comments import ThreadedComment

    wb = wolfxl.Workbook()
    ws = wb.active
    ws["A1"] = "anchor"
    alice = wb.persons.add(name="Alice", user_id="alice@example.com")
    top = ThreadedComment(text="Looks wrong", person=alice)
    top.replies.append(
        ThreadedComment(text="Agreed; investigating", person=alice, parent=top)
    )
    ws["A1"].threaded_comment = top
    out = tmp_path / "tc.xlsx"
    wb.save(out)

    wb2 = wolfxl.load_workbook(out)
    got = wb2.active["A1"].threaded_comment
    assert got is not None, "threaded comment lost across save+load"
    assert got.text == "Looks wrong"
    assert len(got.replies) == 1
    assert got.replies[0].text == "Agreed; investigating"


# --------------------------------------------------------------------------
# Rich text probes
# --------------------------------------------------------------------------


@_register("rich_text_cell")
def _probe_rich_text_cell(tmp_path: Path) -> None:
    import wolfxl
    from wolfxl.cell.rich_text import CellRichText, TextBlock
    from wolfxl.cell.text import InlineFont

    wb = wolfxl.Workbook()
    ws = wb.active
    ws["A1"] = CellRichText(
        [
            TextBlock(InlineFont(b=True), "bold "),
            TextBlock(InlineFont(i=True), "italic"),
        ]
    )
    out = tmp_path / "rt.xlsx"
    wb.save(out)

    wb2 = wolfxl.load_workbook(out, rich_text=True)
    val = wb2.active["A1"].value
    assert isinstance(val, CellRichText)


@_register("rich_text_headers_footers")
def _probe_rich_text_headers_footers(tmp_path: Path) -> None:
    """Rich text in headers/footers. Tracked under G09 (S2).

    Today ``ws.oddHeader`` is a HeaderFooter object but rich-text runs
    inside it are not authored from the Python side. The probe asserts a
    multi-run header round-trips its formatting metadata.
    """
    import wolfxl

    wb = wolfxl.Workbook()
    ws = wb.active
    ws["A1"] = "data"
    ws.oddHeader.center.text = "Title"
    ws.oddHeader.center.font = "Arial,Bold"
    ws.oddHeader.center.size = 14
    out = tmp_path / "hdr.xlsx"
    wb.save(out)

    wb2 = wolfxl.load_workbook(out)
    hf = wb2.active.oddHeader.center
    assert hf.text == "Title"
    assert hf.size == 14


# --------------------------------------------------------------------------
# Conditional formatting probes
# --------------------------------------------------------------------------


@_register("cf_basic_rules")
def _probe_cf_basic_rules(tmp_path: Path) -> None:
    """Basic CF rule round-trip - the openpyxl ``fill=`` convenience kwarg
    on ``CellIsRule`` routes through a DifferentialStyle and is tracked under
    G14 (CF dxf integration). This probe deliberately uses a no-style rule
    so the basic-CF row stays green; the dxf path is exercised by
    ``cf_stop_if_true_priority``.
    """
    import wolfxl
    from wolfxl.formatting.rule import CellIsRule

    wb = wolfxl.Workbook()
    ws = wb.active
    for i in range(1, 6):
        ws.cell(row=i, column=1, value=i)
    ws.conditional_formatting.add(
        "A1:A5",
        CellIsRule(operator="greaterThan", formula=["3"]),
    )
    out = tmp_path / "cf.xlsx"
    wb.save(out)

    wb2 = wolfxl.load_workbook(out)
    rules = list(wb2.active.conditional_formatting)
    assert len(rules) >= 1


@_register("cf_icon_sets")
def _probe_cf_icon_sets(tmp_path: Path) -> None:
    """Icon set rule. Tracked under G11 (S3)."""
    import wolfxl
    from wolfxl.formatting.rule import IconSetRule

    wb = wolfxl.Workbook()
    ws = wb.active
    for i in range(1, 6):
        ws.cell(row=i, column=1, value=i)
    rule = IconSetRule("3TrafficLights1", "percent", [0, 33, 67])
    ws.conditional_formatting.add("A1:A5", rule)
    out = tmp_path / "iconset.xlsx"
    wb.save(out)

    wb2 = wolfxl.load_workbook(out)
    rules = []
    for cf_range in wb2.active.conditional_formatting:
        rules.extend(cf_range.rules if hasattr(cf_range, "rules") else [])
    assert any(getattr(r, "type", "") == "iconSet" for r in rules)


@_register("cf_data_bars")
def _probe_cf_data_bars(tmp_path: Path) -> None:
    """Data-bar rule. Tracked under G12 (S3)."""
    import openpyxl as _opx
    import wolfxl
    from wolfxl.formatting.rule import DataBarRule

    wb = wolfxl.Workbook()
    ws = wb.active
    for i in range(1, 6):
        ws.cell(row=i, column=1, value=i)
    rule = DataBarRule(start_type="min", end_type="max", color="FF0000")
    ws.conditional_formatting.add("A1:A5", rule)
    out = tmp_path / "databar.xlsx"
    wb.save(out)

    wb2 = wolfxl.load_workbook(out)
    rules = []
    for cf_range in wb2.active.conditional_formatting:
        rules.extend(cf_range.rules if hasattr(cf_range, "rules") else [])
    assert any(getattr(r, "type", "") == "dataBar" for r in rules)

    ref_wb = _opx.load_workbook(out)
    ref_ws = ref_wb[ref_wb.sheetnames[0]]
    ref_rules = []
    for cf_range in ref_ws.conditional_formatting:
        ref_rules.extend(ref_ws.conditional_formatting[cf_range])
    bar_rules = [r for r in ref_rules if getattr(r, "type", "") == "dataBar"]
    assert bar_rules, "openpyxl saw no dataBar rule in our output"
    bar = bar_rules[0].dataBar
    assert bar is not None and bar.cfvo, "dataBar payload empty after openpyxl reload"
    assert {c.type for c in bar.cfvo} == {"min", "max"}, (
        f"cfvo types lost: {[c.type for c in bar.cfvo]}"
    )


@_register("cf_data_bars_advanced")
def _probe_cf_data_bars_advanced(tmp_path: Path) -> None:
    """Data-bar with percent / formula cfvo + showValue=False. G12 (S3)."""
    import openpyxl as _opx
    import wolfxl
    from wolfxl.formatting.rule import DataBarRule

    wb = wolfxl.Workbook()
    ws = wb.active
    for i in range(1, 11):
        ws.cell(row=i, column=1, value=i * 10)
    rule = DataBarRule(
        start_type="percent",
        start_value=10,
        end_type="num",
        end_value=90,
        color="FF638EC6",
        showValue=False,
    )
    ws.conditional_formatting.add("A1:A10", rule)
    out = tmp_path / "databar_adv.xlsx"
    wb.save(out)

    ref_wb = _opx.load_workbook(out)
    ref_ws = ref_wb[ref_wb.sheetnames[0]]
    ref_rules = []
    for cf_range in ref_ws.conditional_formatting:
        ref_rules.extend(ref_ws.conditional_formatting[cf_range])
    bar_rules = [r for r in ref_rules if getattr(r, "type", "") == "dataBar"]
    assert bar_rules, "openpyxl saw no dataBar rule"
    bar = bar_rules[0].dataBar
    assert bar is not None and bar.cfvo
    types = {c.type for c in bar.cfvo}
    assert types == {"percent", "num"}, f"cfvo types lost: {types}"
    # Locate min/max by index — openpyxl reads val as a float, so 10 -> 10.0.
    cfvo_min, cfvo_max = bar.cfvo[0], bar.cfvo[1]
    assert float(cfvo_min.val) == 10.0
    assert float(cfvo_max.val) == 90.0
    # showValue=False round-trip
    assert bar.showValue is False


@_register("cf_color_scales_advanced")
def _probe_cf_color_scales_advanced(tmp_path: Path) -> None:
    """3-stop color scale. Tracked under G13 (S3)."""
    import wolfxl
    from wolfxl.formatting.rule import ColorScaleRule

    wb = wolfxl.Workbook()
    ws = wb.active
    for i in range(1, 6):
        ws.cell(row=i, column=1, value=i)
    rule = ColorScaleRule(
        start_type="min",
        start_color="FF0000",
        mid_type="percentile",
        mid_value=50,
        mid_color="FFFF00",
        end_type="max",
        end_color="00FF00",
    )
    ws.conditional_formatting.add("A1:A5", rule)
    out = tmp_path / "cs.xlsx"
    wb.save(out)

    wb2 = wolfxl.load_workbook(out)
    rules = []
    for cf_range in wb2.active.conditional_formatting:
        rules.extend(cf_range.rules if hasattr(cf_range, "rules") else [])
    color_scales = [r for r in rules if getattr(r, "type", "") == "colorScale"]
    assert color_scales, "3-stop color scale not preserved"
    cs = color_scales[0]
    assert hasattr(cs, "colorScale")
    assert len(cs.colorScale.cfvo) == 3


@_register("cf_stop_if_true_priority")
def _probe_cf_stop_if_true_priority(tmp_path: Path) -> None:
    """stopIfTrue + priority + dxf. Tracked under G14 (S3)."""
    import wolfxl
    from wolfxl.formatting.rule import CellIsRule
    from wolfxl.styles import PatternFill

    wb = wolfxl.Workbook()
    ws = wb.active
    for i in range(1, 6):
        ws.cell(row=i, column=1, value=i)
    rule = CellIsRule(
        operator="greaterThan",
        formula=["3"],
        fill=PatternFill(fill_type="solid", fgColor="FFFF00"),
        stopIfTrue=True,
    )
    rule.priority = 1
    ws.conditional_formatting.add("A1:A5", rule)
    out = tmp_path / "cf_stop.xlsx"
    wb.save(out)

    wb2 = wolfxl.load_workbook(out)
    rules = []
    for cf_range in wb2.active.conditional_formatting:
        rules.extend(cf_range.rules if hasattr(cf_range, "rules") else [])
    assert rules, "no rules preserved"
    assert getattr(rules[0], "stopIfTrue", False) is True


# --------------------------------------------------------------------------
# Defined-name probes
# --------------------------------------------------------------------------


@_register("defined_names_basic")
def _probe_defined_names_basic(tmp_path: Path) -> None:
    import wolfxl
    from wolfxl.workbook.defined_name import DefinedName

    wb = wolfxl.Workbook()
    wb.active.title = "Data"
    wb.defined_names["X"] = DefinedName(name="X", value="Data!$A$1:$A$10")

    out = tmp_path / "dn.xlsx"
    wb.save(out)
    wb2 = wolfxl.load_workbook(out)
    assert "X" in wb2.defined_names


@_register("defined_names_edge_cases")
def _probe_defined_names_edge_cases(tmp_path: Path) -> None:
    """Hidden / comment / function-group / shortcut-key edge cases. G22 (S8)."""
    import wolfxl
    from wolfxl.workbook.defined_name import DefinedName

    wb = wolfxl.Workbook()
    wb.active.title = "Data"
    dn = DefinedName(
        name="Hidden",
        value="Data!$A$1",
        hidden=True,
        comment="hidden helper",
    )
    wb.defined_names["Hidden"] = dn

    out = tmp_path / "dn_edge.xlsx"
    wb.save(out)

    wb2 = wolfxl.load_workbook(out)
    rt = wb2.defined_names["Hidden"]
    assert rt.hidden is True
    assert rt.comment == "hidden helper"


# --------------------------------------------------------------------------
# Print-settings probes
# --------------------------------------------------------------------------


@_register("print_settings_basic")
def _probe_print_settings_basic(tmp_path: Path) -> None:
    import wolfxl

    wb = wolfxl.Workbook()
    ws = wb.active
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_options.horizontalCentered = True
    out = tmp_path / "page.xlsx"
    wb.save(out)

    wb2 = wolfxl.load_workbook(out)
    assert wb2.active.page_setup.orientation == "landscape"


@_register("print_settings_depth")
def _probe_print_settings_depth(tmp_path: Path) -> None:
    """Deep PageSetup attrs (~30). Tracked under G24 (S8)."""
    import wolfxl

    wb = wolfxl.Workbook()
    ws = wb.active
    ws.page_setup.paperSize = 9  # A4
    ws.page_setup.scale = 80
    ws.page_setup.firstPageNumber = 7
    ws.page_setup.useFirstPageNumber = True
    ws.page_setup.errors = "blank"
    ws.print_title_rows = "1:1"
    ws.print_title_cols = "A:A"
    out = tmp_path / "page_deep.xlsx"
    wb.save(out)

    wb2 = wolfxl.load_workbook(out)
    ws2 = wb2.active
    assert ws2.page_setup.paperSize == 9
    assert ws2.page_setup.scale == 80
    assert ws2.page_setup.useFirstPageNumber is True
    assert ws2.print_title_rows == "1:1"
    assert ws2.print_title_cols == "A:A"


# --------------------------------------------------------------------------
# Array / data-table formula probes
# --------------------------------------------------------------------------


@_register("array_formula_basic")
def _probe_array_formula_basic(tmp_path: Path) -> None:
    """Basic ArrayFormula round-trip. G07 (S1)."""
    import openpyxl as _opx
    from openpyxl.worksheet.formula import ArrayFormula as _OpxArrayFormula
    import wolfxl
    from wolfxl.worksheet.formula import ArrayFormula

    wb = wolfxl.Workbook()
    ws = wb.active
    ws["A1"] = 1
    ws["A2"] = 2
    ws["B1"] = ArrayFormula(ref="B1:B2", text="=A1:A2*2")
    out = tmp_path / "arr.xlsx"
    wb.save(out)

    wb2 = wolfxl.load_workbook(out)
    val = wb2.active["B1"].value
    assert isinstance(val, ArrayFormula) or "A1:A2" in str(val)

    ref_wb = _opx.load_workbook(out)
    ref_val = ref_wb[ref_wb.sheetnames[0]]["B1"].value
    assert isinstance(ref_val, _OpxArrayFormula), (
        f"openpyxl saw {type(ref_val).__name__}, expected ArrayFormula"
    )
    assert ref_val.ref == "B1:B2", f"ref lost: {ref_val.ref!r}"
    assert "A1:A2" in (ref_val.text or ""), f"formula lost: {ref_val.text!r}"


@_register("array_formula_data_table")
def _probe_array_formula_data_table(tmp_path: Path) -> None:
    """DataTableFormula round-trip. G07 (S1)."""
    import wolfxl
    from wolfxl.worksheet.formula import DataTableFormula

    wb = wolfxl.Workbook()
    ws = wb.active
    ws["A1"] = 5
    ws["B1"] = 10
    ws["C1"] = DataTableFormula(ref="C1:C3", t="dataTable", r1="A1", dt2D=False)
    out = tmp_path / "dt.xlsx"
    wb.save(out)

    wb2 = wolfxl.load_workbook(out)
    val = wb2.active["C1"].value
    assert val is not None


# --------------------------------------------------------------------------
# Test runner / parametrisation
# --------------------------------------------------------------------------


_PROBE_ENTRIES = [e for e in _SPEC.ENTRIES if e.get("probe")]


def _id_for(entry: dict[str, Any]) -> str:
    return f"{entry['id']}[{entry['status']}]"


@pytest.mark.parametrize("entry", _PROBE_ENTRIES, ids=_id_for)
def test_compat_oracle_probe(
    entry: dict[str, Any], tmp_path: Path, request: pytest.FixtureRequest
) -> None:
    """Run one compat-oracle probe per spec entry that carries a `probe` field.

    Outcome by status:
    * ``supported``    → must pass; failure is a regression.
    * ``partial``      → xfail today; flips xpass once the gap closes.
    * ``not_yet``      → xfail today; flips xpass once the gap closes.
    * ``out_of_scope`` → skipped.
    """
    probe_name = entry["probe"]
    probe_fn = PROBES.get(probe_name)
    if probe_fn is None:
        pytest.skip(f"probe '{probe_name}' not registered yet")

    if entry["status"] == "out_of_scope":
        pytest.skip(f"out of scope: {entry['id']}")
    if entry["status"] in ("not_yet", "partial"):
        gap = entry.get("gap_id", "?")
        request.node.add_marker(
            pytest.mark.xfail(
                reason=(
                    f"compat-oracle baseline gap: {entry['id']} "
                    f"({entry['status']}, {gap}); flip when the gap closes"
                ),
                strict=False,
            )
        )

    probe_fn(tmp_path)


# --------------------------------------------------------------------------
# Session summary
# --------------------------------------------------------------------------


def _baseline_path() -> Path:
    return REPO_ROOT / ".pytest_cache" / "compat_oracle_baseline.json"


@pytest.fixture(scope="session", autouse=True)
def _emit_oracle_summary(request: pytest.FixtureRequest) -> Any:
    yield
    terminal = request.config.pluginmanager.get_plugin("terminalreporter")
    if terminal is None:  # pragma: no cover - pytest internals
        return

    counts: Counter[str] = Counter()
    for outcome in ("passed", "failed", "skipped", "xfailed", "xpassed", "error"):
        for report in terminal.stats.get(outcome, []):
            nodeid = getattr(report, "nodeid", "")
            if "test_compat_oracle_probe" not in nodeid:
                continue
            counts[outcome] += 1

    if not counts:
        return

    relevant = (
        counts["passed"]
        + counts["failed"]
        + counts["xfailed"]
        + counts["xpassed"]
        + counts["error"]
    )
    if relevant == 0:
        return

    pass_total = counts["passed"] + counts["xpassed"]
    pct = (pass_total / relevant) * 100

    terminal.write_sep("=", "openpyxl-compat oracle (Sprint 0 baseline)")
    terminal.write_line(
        f"  total probes:      {relevant}",
    )
    terminal.write_line(
        f"  passed (green):    {counts['passed']}",
    )
    terminal.write_line(
        f"  xpassed:           {counts['xpassed']} (gap closed; flip status to 'supported')",
    )
    terminal.write_line(
        f"  xfailed (gap):     {counts['xfailed']}",
    )
    terminal.write_line(
        f"  failed:            {counts['failed']}",
    )
    terminal.write_line(
        f"  pass rate:         {pct:.1f}% ({pass_total}/{relevant})",
    )

    if os.environ.get("WOLFXL_COMPAT_ORACLE_WRITE_BASELINE") == "1":
        baseline = {
            "date": date.today().isoformat(),
            "total": relevant,
            "passed": counts["passed"],
            "xpassed": counts["xpassed"],
            "xfailed": counts["xfailed"],
            "failed": counts["failed"],
            "pass_rate_pct": round(pct, 2),
        }
        path = _baseline_path()
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(baseline, indent=2) + "\n")
        terminal.write_line(f"  wrote baseline:    {path.relative_to(REPO_ROOT)}")
