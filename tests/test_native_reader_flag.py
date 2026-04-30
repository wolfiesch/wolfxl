"""Opt-in smoke tests for the native XLSX reader.

The native reader is intentionally hidden behind ``WOLFXL_NATIVE_READER`` while
it grows to parity. These tests pin the first public seam without changing the
default calamine-backed path.
"""

from __future__ import annotations

import datetime as dt
import re
import zipfile
from pathlib import Path

import pytest

openpyxl = pytest.importorskip("openpyxl")
openpyxl_datavalidation = pytest.importorskip("openpyxl.worksheet.datavalidation")
openpyxl_hyperlink = pytest.importorskip("openpyxl.worksheet.hyperlink")
openpyxl_image = pytest.importorskip("openpyxl.drawing.image")
wolfxl = pytest.importorskip("wolfxl")

DataValidation = openpyxl_datavalidation.DataValidation
Hyperlink = openpyxl_hyperlink.Hyperlink
OpenpyxlImage = openpyxl_image.Image
Border = openpyxl.styles.Border
Font = openpyxl.styles.Font
PatternFill = openpyxl.styles.PatternFill
Side = openpyxl.styles.Side
Alignment = openpyxl.styles.Alignment

FIXTURES = Path(__file__).parent / "fixtures"
PNG_PATH = FIXTURES / "images" / "tiny_red_dot.png"


def _make_basic_xlsx(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"] = "Label"
    ws["B1"] = 42
    ws["B1"].number_format = "#,##0.00"
    ws["B1"].font = Font(
        name="Arial",
        size=14,
        bold=True,
        italic=True,
        underline="single",
        strike=True,
        color="FF123456",
    )
    ws["B1"].fill = PatternFill(patternType="solid", fgColor="FFABCDEF")
    ws["B1"].alignment = Alignment(
        horizontal="center",
        vertical="top",
        wrap_text=True,
        text_rotation=45,
        indent=2,
    )
    ws["B1"].border = Border(
        left=Side(style="thin", color="FFFF0000"),
        right=Side(style="medium"),
        top=Side(style="double", color="FF00FF00"),
        bottom=Side(style="dashed", color="FF0000FF"),
    )
    ws["C1"] = True
    ws["A2"] = "Formula"
    ws["B2"] = "=B1*2"
    ws["A3"] = dt.datetime(2024, 1, 15, 12, 30)
    ws["B3"] = dt.date(2024, 6, 1)
    ws["D1"] = "Anchor"
    ws["D1"].number_format = "#,##0"
    ws.merge_cells("D1:E1")
    ws["A5"] = "External"
    ws["A5"].hyperlink = Hyperlink(
        ref="A5",
        target="https://example.com/report",
        tooltip="Example report",
    )
    ws["B5"] = "Internal"
    ws["B5"].hyperlink = Hyperlink(ref="B5", location="Data!A1", display="Jump")
    ws["A6"] = "Commented"
    ws["A6"].comment = openpyxl.comments.Comment("Native reader note", "Wolf")
    ws.row_dimensions[6].height = 24
    ws.column_dimensions["C"].width = 18
    dv = DataValidation(type="list", formula1='"Red,Blue"', allow_blank=True)
    dv.add("C2:C6")
    ws.add_data_validation(dv)
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = "A1:D6"
    ws.protection.sheet = True
    ws.protection.objects = True
    ws.protection.formatCells = False
    ws.protection.sort = False
    ws.protection.set_password("hunter2")
    wb.save(path)
    wb.close()
    _inject_workbook_security(path)
    _inject_auto_filter_details(path)
    _inject_merged_subordinate_style(path)


def _inject_workbook_security(path: Path) -> None:
    with zipfile.ZipFile(path, "r") as zin:
        entries = {name: zin.read(name) for name in zin.namelist()}
    workbook_name = "xl/workbook.xml"
    workbook_xml = entries[workbook_name].decode()
    workbook_xml = re.sub(r"<fileSharing\b[^>]*/>", "", workbook_xml)
    workbook_xml = re.sub(r"<workbookProtection\b[^>]*/>", "", workbook_xml)
    security_xml = (
        '<fileSharing readOnlyRecommended="1" userName="Wolf" algorithmName="SHA-512" '
        'hashValue="FILEHASH" saltValue="FILESALT" spinCount="100000"/>'
        '<workbookProtection lockStructure="1" workbookAlgorithmName="SHA-512" '
        'workbookHashValue="HASH" workbookSaltValue="SALT" workbookSpinCount="100000"/>'
    )
    workbook_xml, count = re.subn(r"(<workbookPr\b)", security_xml + r"\1", workbook_xml)
    assert count == 1
    entries[workbook_name] = workbook_xml.encode()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in entries.items():
            zout.writestr(name, data)


def _inject_merged_subordinate_style(path: Path) -> None:
    """Add a styled blank subordinate cell to the merged D1:E1 range."""
    with zipfile.ZipFile(path, "r") as zin:
        entries = {name: zin.read(name) for name in zin.namelist()}
    sheet_name = "xl/worksheets/sheet1.xml"
    sheet_xml = entries[sheet_name].decode()
    if '<c r="E1"' in sheet_xml:
        return
    sheet_xml = sheet_xml.replace("</row>", '<c r="E1" s="1"/></row>', 1)
    entries[sheet_name] = sheet_xml.encode()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in entries.items():
            zout.writestr(name, data)


def _inject_auto_filter_details(path: Path) -> None:
    with zipfile.ZipFile(path, "r") as zin:
        entries = {name: zin.read(name) for name in zin.namelist()}
    sheet_name = "xl/worksheets/sheet1.xml"
    sheet_xml = entries[sheet_name].decode()
    auto_filter_xml = (
        '<autoFilter ref="A1:D6">'
        '<filterColumn colId="0"><filters><filter val="Label"/></filters></filterColumn>'
        '<filterColumn colId="1"><customFilters and="1">'
        '<customFilter operator="greaterThan" val="10"/>'
        '</customFilters></filterColumn>'
        '<sortState ref="A2:D6"><sortCondition ref="B2:B6" descending="1"/></sortState>'
        "</autoFilter>"
    )
    sheet_xml, count = re.subn(r'<autoFilter ref="A1:D6"\s*/>', auto_filter_xml, sheet_xml)
    assert count == 1
    entries[sheet_name] = sheet_xml.encode()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in entries.items():
            zout.writestr(name, data)


def _make_shared_formula_xlsx(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    for row in range(1, 4):
        ws[f"A{row}"] = row
        ws[f"B{row}"] = f"=A{row}*2"
    wb.save(path)
    wb.close()

    with zipfile.ZipFile(path, "r") as zin:
        entries = {name: zin.read(name) for name in zin.namelist()}
    sheet_name = "xl/worksheets/sheet1.xml"
    sheet_xml = entries[sheet_name].decode()
    replacements = {
        r'<c r="B1"[^>]*>\s*<f>A1\*2</f>\s*(?:<v(?:>[^<]*</v>|\s*/>)\s*)?</c>': (
            '<c r="B1"><f t="shared" si="0" ref="B1:B3">A1*2</f><v/></c>'
        ),
        r'<c r="B2"[^>]*>\s*<f>A2\*2</f>\s*(?:<v(?:>[^<]*</v>|\s*/>)\s*)?</c>': (
            '<c r="B2"><f t="shared" si="0"/><v/></c>'
        ),
        r'<c r="B3"[^>]*>\s*<f>A3\*2</f>\s*(?:<v(?:>[^<]*</v>|\s*/>)\s*)?</c>': (
            '<c r="B3"><f t="shared" si="0"/><v/></c>'
        ),
    }
    for pattern, replacement in replacements.items():
        sheet_xml, count = re.subn(pattern, replacement, sheet_xml)
        assert count == 1
    entries[sheet_name] = sheet_xml.encode()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in entries.items():
            zout.writestr(name, data)


def _make_image_xlsx(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Images"
    ws["A1"] = "with image"
    ws.add_image(OpenpyxlImage(PNG_PATH), "B5")
    wb.save(path)
    wb.close()


def _make_wolfxl_anchor_image_xlsx(path: Path) -> None:
    from wolfxl.drawing import Image as WolfImage
    from wolfxl.drawing.spreadsheet_drawing import (
        AbsoluteAnchor,
        AnchorMarker,
        TwoCellAnchor,
        XDRPoint2D,
        XDRPositiveSize2D,
    )

    wb = wolfxl.Workbook()
    ws = wb.active
    ws.title = "Anchors"
    ws.add_image(
        WolfImage(PNG_PATH),
        TwoCellAnchor(
            _from=AnchorMarker(col=1, row=1, colOff=10, rowOff=20),
            to=AnchorMarker(col=3, row=4, colOff=30, rowOff=40),
            editAs="twoCell",
        ),
    )
    ws.add_image(
        WolfImage(PNG_PATH),
        AbsoluteAnchor(
            pos=XDRPoint2D(x=1000, y=2000),
            ext=XDRPositiveSize2D(cx=3000, cy=4000),
        ),
    )
    wb.save(path)


def _make_sheet_state_xlsx(path: Path) -> None:
    wb = openpyxl.Workbook()
    wb.active.title = "Visible"
    hidden = wb.create_sheet("Hidden")
    hidden.sheet_state = "hidden"
    very_hidden = wb.create_sheet("VeryHidden")
    very_hidden.sheet_state = "veryHidden"
    wb.save(path)
    wb.close()


def _make_print_area_xlsx(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Print"
    ws.print_area = "A1:D10"
    wb.save(path)
    wb.close()


def _make_print_setup_xlsx(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Setup"
    ws.page_margins.left = 1.1
    ws.page_margins.right = 1.2
    ws.page_margins.top = 1.3
    ws.page_margins.bottom = 1.4
    ws.page_margins.header = 0.5
    ws.page_margins.footer = 0.6
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = 9
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.scale = 75
    ws.oddHeader.left.text = "Left"
    ws.oddHeader.center.text = "Center"
    ws.oddFooter.right.text = "Page &P"
    ws.HeaderFooter.differentFirst = True
    ws.HeaderFooter.alignWithMargins = False
    wb.save(path)
    wb.close()


def _make_chart_xlsx(path: Path) -> None:
    from openpyxl.chart import BarChart, Reference

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Charts"
    rows = [
        ("Month", "Sales"),
        ("Jan", 10),
        ("Feb", 20),
        ("Mar", 30),
    ]
    for row in rows:
        ws.append(row)
    chart = BarChart()
    chart.title = "Sales Trend"
    chart.style = 10
    data = Reference(ws, min_col=2, min_row=1, max_row=4)
    cats = Reference(ws, min_col=1, min_row=2, max_row=4)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    ws.add_chart(chart, "D5")
    wb.save(path)
    wb.close()


def test_native_reader_flag_loads_path_values(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    path = tmp_path / "native-smoke.xlsx"
    _make_basic_xlsx(path)

    monkeypatch.setenv("WOLFXL_NATIVE_READER", "1")
    wb = wolfxl.load_workbook(path, data_only=False)
    try:
        assert wb._rust_reader.__class__.__name__ == "NativeXlsxBook"  # noqa: SLF001
        assert wb.sheetnames == ["Data"]
        ws = wb["Data"]
        assert ws["A1"].value == "Label"
        assert ws["B1"].value == 42
        assert ws["B1"].number_format == "#,##0.00"
        assert ws["B1"].font.name == "Arial"
        assert ws["B1"].font.size == 14.0
        assert ws["B1"].font.bold is True
        assert ws["B1"].font.italic is True
        assert ws["B1"].font.underline == "single"
        assert ws["B1"].font.strike is True
        assert ws["B1"].font.color == "#123456"
        assert ws["B1"].fill.patternType == "solid"
        assert ws["B1"].fill.fgColor == "#ABCDEF"
        assert ws["B1"].alignment.horizontal == "center"
        assert ws["B1"].alignment.vertical == "top"
        assert ws["B1"].alignment.wrap_text is True
        assert ws["B1"].alignment.text_rotation == 45
        assert ws["B1"].alignment.indent == 2
        assert ws["B1"].border.left.style == "thin"
        assert ws["B1"].border.left.color == "#FF0000"
        assert ws["B1"].border.right.style == "medium"
        assert ws["B1"].border.right.color == "#000000"
        assert ws["B1"].border.top.style == "double"
        assert ws["B1"].border.top.color == "#00FF00"
        assert ws["B1"].border.bottom.style == "dashed"
        assert ws["B1"].border.bottom.color == "#0000FF"
        assert ws["A1"].border.left.style is None
        assert ws["C1"].value is True
        assert ws["B2"].value == "=B1*2"
        assert wb._rust_reader.read_cell_formula("Data", "B2") == {  # noqa: SLF001
            "type": "formula",
            "formula": "=B1*2",
            "value": "=B1*2",
        }
        assert wb._rust_reader.read_cell_formula("Data", "A1") is None  # noqa: SLF001
        assert ws["A3"].value == dt.datetime(2024, 1, 15, 12, 30)
        assert ws["B3"].value == dt.datetime(2024, 6, 1, 0, 0)
        assert ws["A5"].hyperlink is not None
        assert ws["A5"].hyperlink.target == "https://example.com/report"
        assert ws["A5"].hyperlink.display == "External"
        assert ws["A5"].hyperlink.tooltip == "Example report"
        assert ws["B5"].hyperlink is not None
        assert ws["B5"].hyperlink.target is None
        assert ws["B5"].hyperlink.location == "Data!A1"
        assert ws["B5"].hyperlink.display == "Jump"
        assert ws["A6"].comment is not None
        assert ws["A6"].comment.text == "Native reader note"
        assert ws["A6"].comment.author == "Wolf"
        assert ws.row_dimensions[6].height == 24
        assert ws.column_dimensions["C"].width == 18
        validations = list(ws.data_validations)
        assert len(validations) == 1
        assert validations[0].type == "list"
        assert validations[0].formula1 == '="Red,Blue"'
        assert validations[0].allowBlank is True
        assert validations[0].sqref == "C2:C6"
        assert ws.freeze_panes == "B2"
        assert ws.auto_filter.ref == "A1:D6"
        assert len(ws.auto_filter.filter_columns) == 2
        assert ws.auto_filter.filter_columns[0].col_id == 0
        assert ws.auto_filter.filter_columns[0].filter.values == ["Label"]
        assert ws.auto_filter.filter_columns[1].filter.and_ is True
        assert ws.auto_filter.filter_columns[1].filter.customFilter[0].operator == "greaterThan"
        assert ws.auto_filter.filter_columns[1].filter.customFilter[0].val == "10"
        assert ws.auto_filter.sort_state is not None
        assert ws.auto_filter.sort_state.ref == "A2:D6"
        assert ws.auto_filter.sort_state.sort_conditions[0].ref == "B2:B6"
        assert ws.auto_filter.sort_state.sort_conditions[0].descending is True
        assert ws.protection.sheet is True
        assert ws.protection.objects is True
        assert ws.protection.formatCells is False
        assert ws.protection.sort is False
        assert ws.protection.password == "C258"
        assert wb.security is not None
        assert wb.security.lock_structure is True
        assert wb.security.workbook_algorithm_name == "SHA-512"
        assert wb.security.workbook_hash_value == "HASH"
        assert wb.security.workbook_salt_value == "SALT"
        assert wb.security.workbook_spin_count == 100000
        assert wb.fileSharing is not None
        assert wb.fileSharing.read_only_recommended is True
        assert wb.fileSharing.user_name == "Wolf"
        assert wb.fileSharing.algorithm_name == "SHA-512"
        assert wb.fileSharing.hash_value == "FILEHASH"
        assert wb.fileSharing.salt_value == "FILESALT"
        assert wb.fileSharing.spin_count == 100000
        assert {str(r) for r in ws.merged_cells.ranges} == {"D1:E1"}
        assert ws["D1"].number_format == "#,##0"
        assert ws["E1"].number_format is None
        assert ws["E1"].font.name is None
        records = {
            record["coordinate"]: record
            for record in ws.cell_records(include_format=True, include_empty=True)
        }
        assert records["B1"]["number_format"] == "#,##0.00"
        assert "number_format" not in records["E1"]
        assert records["A3"]["data_type"] == "datetime"
    finally:
        wb.close()


def test_native_reader_loads_workbook_sheet_states(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    path = tmp_path / "native-sheet-states.xlsx"
    _make_sheet_state_xlsx(path)

    monkeypatch.setenv("WOLFXL_NATIVE_READER", "1")
    wb = wolfxl.load_workbook(path)
    try:
        assert wb.sheetnames == ["Visible", "Hidden", "VeryHidden"]
        assert wb["Visible"].sheet_state == "visible"
        assert wb["Hidden"].sheet_state == "hidden"
        assert wb["VeryHidden"].sheet_state == "veryHidden"
    finally:
        wb.close()


def test_native_reader_loads_print_area(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    path = tmp_path / "native-print-area.xlsx"
    _make_print_area_xlsx(path)
    expected = openpyxl.load_workbook(path)
    try:
        expected_print_area = expected["Print"].print_area
    finally:
        expected.close()

    monkeypatch.setenv("WOLFXL_NATIVE_READER", "1")
    wb = wolfxl.load_workbook(path)
    try:
        assert wb["Print"].print_area == expected_print_area
    finally:
        wb.close()


def test_native_reader_loads_print_setup_metadata(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    path = tmp_path / "native-print-setup.xlsx"
    _make_print_setup_xlsx(path)

    monkeypatch.setenv("WOLFXL_NATIVE_READER", "1")
    wb = wolfxl.load_workbook(path)
    try:
        ws = wb["Setup"]
        assert ws.page_margins.left == 1.1
        assert ws.page_margins.right == 1.2
        assert ws.page_margins.top == 1.3
        assert ws.page_margins.bottom == 1.4
        assert ws.page_margins.header == 0.5
        assert ws.page_margins.footer == 0.6
        assert ws.page_setup.orientation == "landscape"
        assert ws.page_setup.paperSize == 9
        assert ws.page_setup.fitToWidth == 1
        assert ws.page_setup.fitToHeight == 0
        assert ws.page_setup.scale == 75
        assert ws.header_footer.odd_header.left == "Left"
        assert ws.header_footer.odd_header.center == "Center"
        assert ws.header_footer.odd_footer.right == "Page &P"
        assert ws.header_footer.different_first is True
        assert ws.header_footer.align_with_margins is False
    finally:
        wb.close()


def test_native_reader_loads_drawing_images(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    path = tmp_path / "native-images.xlsx"
    _make_image_xlsx(path)

    monkeypatch.setenv("WOLFXL_NATIVE_READER", "1")
    wb = wolfxl.load_workbook(path)
    try:
        ws = wb["Images"]
        images = ws._images  # noqa: SLF001
        assert len(images) == 1
        img = images[0]
        assert img.format == "png"
        assert img._data == PNG_PATH.read_bytes()  # noqa: SLF001

        from wolfxl.drawing.spreadsheet_drawing import OneCellAnchor

        assert isinstance(img.anchor, OneCellAnchor)
        assert img.anchor._from.col == 1
        assert img.anchor._from.row == 4
        assert img.anchor.ext is not None
        assert img.anchor.ext.cx > 0
        assert img.anchor.ext.cy > 0
    finally:
        wb.close()


def test_native_reader_loads_drawing_charts(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    path = tmp_path / "native-charts.xlsx"
    _make_chart_xlsx(path)

    monkeypatch.setenv("WOLFXL_NATIVE_READER", "1")
    wb = wolfxl.load_workbook(path)
    try:
        ws = wb["Charts"]
        charts = ws._charts  # noqa: SLF001
        assert len(charts) == 1
        chart = charts[0]
        assert chart.__class__.__name__ == "BarChart"
        assert chart.title.tx.rich.paragraphs[0].r[0].t == "Sales Trend"
        assert chart.style == 10
        assert len(chart.series) == 1
        series = chart.series[0]
        assert series.tx.strRef.f == "'Charts'!B1"
        assert series.cat.strRef.f == "'Charts'!$A$2:$A$4"
        assert series.val.numRef.f == "'Charts'!$B$2:$B$4"

        from wolfxl.drawing.spreadsheet_drawing import OneCellAnchor

        assert isinstance(chart._anchor, OneCellAnchor)  # noqa: SLF001
        assert chart._anchor._from.col == 3  # noqa: SLF001
        assert chart._anchor._from.row == 4  # noqa: SLF001
    finally:
        wb.close()


def test_native_reader_preserves_image_anchor_types(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    path = tmp_path / "native-image-anchors.xlsx"
    _make_wolfxl_anchor_image_xlsx(path)

    monkeypatch.setenv("WOLFXL_NATIVE_READER", "1")
    wb = wolfxl.load_workbook(path)
    try:
        ws = wb["Anchors"]
        images = ws._images  # noqa: SLF001
        assert len(images) == 2

        from wolfxl.drawing.spreadsheet_drawing import AbsoluteAnchor, TwoCellAnchor

        two_cell = images[0].anchor
        assert isinstance(two_cell, TwoCellAnchor)
        assert two_cell._from.col == 1
        assert two_cell._from.row == 1
        assert two_cell._from.colOff == 10
        assert two_cell._from.rowOff == 20
        assert two_cell.to.col == 3
        assert two_cell.to.row == 4
        assert two_cell.to.colOff == 30
        assert two_cell.to.rowOff == 40
        assert two_cell.editAs == "twoCell"

        absolute = images[1].anchor
        assert isinstance(absolute, AbsoluteAnchor)
        assert absolute.pos.x == 1000
        assert absolute.pos.y == 2000
        assert absolute.ext.cx == 3000
        assert absolute.ext.cy == 4000
    finally:
        wb.close()


def test_native_reader_expands_shared_formula_children(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    path = tmp_path / "native-shared-formulas.xlsx"
    _make_shared_formula_xlsx(path)
    expected_book = openpyxl.load_workbook(path, data_only=False)
    try:
        expected = [expected_book.active[f"B{row}"].value for row in range(1, 4)]
    finally:
        expected_book.close()

    monkeypatch.setenv("WOLFXL_NATIVE_READER", "1")
    wb = wolfxl.load_workbook(path, data_only=False)
    try:
        ws = wb["Sheet1"]
        assert [ws[f"B{row}"].value for row in range(1, 4)] == expected
        assert wb._rust_reader.read_cell_formula("Sheet1", "B2") == {  # noqa: SLF001
            "type": "formula",
            "formula": "=A2*2",
            "value": "=A2*2",
        }
    finally:
        wb.close()


def test_native_reader_flag_loads_bytes(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    path = tmp_path / "native-bytes.xlsx"
    _make_basic_xlsx(path)

    monkeypatch.setenv("WOLFXL_NATIVE_READER", "1")
    wb = wolfxl.load_workbook(path.read_bytes())
    try:
        assert wb._rust_reader.__class__.__name__ == "NativeXlsxBook"  # noqa: SLF001
        assert wb._rust_reader.opened_from_bytes() is True  # noqa: SLF001
        assert wb["Data"].iter_rows(values_only=True).__next__() == (
            "Label",
            42,
            True,
            "Anchor",
            None,
        )
        assert wb["Data"]["A5"].hyperlink.target == "https://example.com/report"
        assert wb["Data"]["A6"].comment.text == "Native reader note"
        assert wb["Data"].column_dimensions["C"].width == 18
        assert len(list(wb["Data"].data_validations)) == 1
    finally:
        wb.close()
