"""RFC-055 — Print/page-setup write/read parity with openpyxl.

These tests exercise the full pipeline:

  Python PageSetup / PageMargins / HeaderFooter / SheetView
  → set_sheet_setup_native PyO3 binding
  → wolfxl-writer's emit_page_setup / emit_page_margins / etc.
  → openpyxl reads it back

If openpyxl can round-trip our XML, Excel/LibreOffice will too.
"""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

wolfxl = pytest.importorskip("wolfxl")
openpyxl = pytest.importorskip("openpyxl")

from wolfxl.worksheet.header_footer import HeaderFooter, HeaderFooterItem
from wolfxl.worksheet.page_setup import PageMargins, PageSetup
from wolfxl.worksheet.views import Pane, SheetView


@pytest.fixture(autouse=True)
def _force_test_epoch(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("WOLFXL_TEST_EPOCH", "0")


@pytest.fixture
def tmp_xlsx(tmp_path: Path) -> Path:
    return tmp_path / "print_settings.xlsx"


def _read_sheet_xml(p: Path) -> str:
    with zipfile.ZipFile(p) as zf:
        return zf.read("xl/worksheets/sheet1.xml").decode("utf-8")


def test_page_setup_orientation_round_trips(tmp_xlsx: Path) -> None:
    wb = wolfxl.Workbook()
    ws = wb.active
    ws["A1"] = "x"
    ws.page_setup.orientation = "landscape"
    wb.save(tmp_xlsx)

    text = _read_sheet_xml(tmp_xlsx)
    assert 'orientation="landscape"' in text

    op = openpyxl.load_workbook(tmp_xlsx)
    try:
        assert op.active.page_setup.orientation == "landscape"
    finally:
        op.close()


def test_page_setup_paper_and_scale(tmp_xlsx: Path) -> None:
    wb = wolfxl.Workbook()
    ws = wb.active
    ws["A1"] = "x"
    ws.page_setup.paperSize = 9
    ws.page_setup.scale = 75
    wb.save(tmp_xlsx)

    op = openpyxl.load_workbook(tmp_xlsx)
    try:
        assert int(op.active.page_setup.paperSize) == 9
        assert int(op.active.page_setup.scale) == 75
    finally:
        op.close()


def test_page_setup_fit_to_width_height(tmp_xlsx: Path) -> None:
    wb = wolfxl.Workbook()
    ws = wb.active
    ws["A1"] = "x"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    wb.save(tmp_xlsx)

    op = openpyxl.load_workbook(tmp_xlsx)
    try:
        assert int(op.active.page_setup.fitToWidth) == 1
        assert int(op.active.page_setup.fitToHeight) == 0
    finally:
        op.close()


def test_page_margins_round_trip(tmp_xlsx: Path) -> None:
    wb = wolfxl.Workbook()
    ws = wb.active
    ws["A1"] = "x"
    ws.page_margins = PageMargins(left=1.0, right=1.5, top=2.0, bottom=2.0,
                                  header=0.5, footer=0.5)
    wb.save(tmp_xlsx)

    op = openpyxl.load_workbook(tmp_xlsx)
    try:
        m = op.active.page_margins
        assert float(m.left) == 1.0
        assert float(m.right) == 1.5
        assert float(m.top) == 2.0
        assert float(m.bottom) == 2.0
        assert float(m.header) == 0.5
        assert float(m.footer) == 0.5
    finally:
        op.close()


def test_header_footer_odd_header_round_trip(tmp_xlsx: Path) -> None:
    wb = wolfxl.Workbook()
    ws = wb.active
    ws["A1"] = "x"
    ws.header_footer = HeaderFooter(
        odd_header=HeaderFooterItem(left="L", center="C", right="R"),
    )
    wb.save(tmp_xlsx)

    text = _read_sheet_xml(tmp_xlsx)
    assert "<oddHeader>" in text
    # openpyxl recognises the embedded format codes.
    op = openpyxl.load_workbook(tmp_xlsx)
    try:
        oh = op.active.oddHeader
        # openpyxl parses the &L/&C/&R back into separate segments
        assert oh.left.text == "L"
        assert oh.center.text == "C"
        assert oh.right.text == "R"
    finally:
        op.close()


def test_header_footer_different_first(tmp_xlsx: Path) -> None:
    wb = wolfxl.Workbook()
    ws = wb.active
    ws["A1"] = "x"
    ws.header_footer.different_first = True
    wb.save(tmp_xlsx)

    op = openpyxl.load_workbook(tmp_xlsx)
    try:
        # openpyxl exposes this as differentFirst=True
        assert op.active.HeaderFooter.differentFirst is True
    finally:
        op.close()


def test_sheet_view_zoom_round_trip(tmp_xlsx: Path) -> None:
    wb = wolfxl.Workbook()
    ws = wb.active
    ws["A1"] = "x"
    ws.sheet_view.zoom_scale = 150
    wb.save(tmp_xlsx)

    op = openpyxl.load_workbook(tmp_xlsx)
    try:
        sv = op.active.sheet_view
        assert int(sv.zoomScale) == 150
    finally:
        op.close()


def test_sheet_view_grid_lines_off(tmp_xlsx: Path) -> None:
    wb = wolfxl.Workbook()
    ws = wb.active
    ws["A1"] = "x"
    ws.sheet_view.show_grid_lines = False
    wb.save(tmp_xlsx)

    op = openpyxl.load_workbook(tmp_xlsx)
    try:
        assert op.active.sheet_view.showGridLines is False
    finally:
        op.close()


def test_sheet_view_freeze_pane_via_typed_view(tmp_xlsx: Path) -> None:
    wb = wolfxl.Workbook()
    ws = wb.active
    ws["A1"] = "x"
    ws.sheet_view = SheetView(
        pane=Pane(ySplit=1.0, topLeftCell="A2", activePane="bottomLeft", state="frozen"),
    )
    wb.save(tmp_xlsx)

    text = _read_sheet_xml(tmp_xlsx)
    assert "<pane " in text
    assert 'ySplit="1"' in text
    assert 'topLeftCell="A2"' in text
    # openpyxl reads the pane back.
    op = openpyxl.load_workbook(tmp_xlsx)
    try:
        pane = op.active.sheet_view.pane
        assert pane is not None
        assert pane.topLeftCell == "A2"
    finally:
        op.close()


def test_combined_print_settings(tmp_xlsx: Path) -> None:
    """Set page_setup + page_margins + header_footer all together;
    verify all three round-trip via openpyxl."""
    wb = wolfxl.Workbook()
    ws = wb.active
    ws["A1"] = "x"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = 9
    ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.75, bottom=0.75,
                                  header=0.3, footer=0.3)
    ws.header_footer.odd_header.center = "Annual Report"
    wb.save(tmp_xlsx)

    op = openpyxl.load_workbook(tmp_xlsx)
    try:
        ws_op = op.active
        assert ws_op.page_setup.orientation == "landscape"
        assert int(ws_op.page_setup.paperSize) == 9
        assert float(ws_op.page_margins.left) == 0.5
        assert ws_op.oddHeader.center.text == "Annual Report"
    finally:
        op.close()
