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
wolfxl = pytest.importorskip("wolfxl")

DataValidation = openpyxl_datavalidation.DataValidation
Hyperlink = openpyxl_hyperlink.Hyperlink
Border = openpyxl.styles.Border
Font = openpyxl.styles.Font
PatternFill = openpyxl.styles.PatternFill
Side = openpyxl.styles.Side
Alignment = openpyxl.styles.Alignment


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
    ws.protection.sheet = True
    ws.protection.objects = True
    ws.protection.formatCells = False
    ws.protection.sort = False
    ws.protection.set_password("hunter2")
    wb.save(path)
    wb.close()
    _inject_workbook_security(path)
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
