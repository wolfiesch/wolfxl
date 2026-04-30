"""Opt-in smoke tests for the native XLSX reader.

The native reader is intentionally hidden behind ``WOLFXL_NATIVE_READER`` while
it grows to parity. These tests pin the first public seam without changing the
default calamine-backed path.
"""

from __future__ import annotations

import datetime as dt
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
        assert {str(r) for r in ws.merged_cells.ranges} == {"D1:E1"}
        records = {record["coordinate"]: record for record in ws.cell_records(include_format=True)}
        assert records["B1"]["number_format"] == "#,##0.00"
        assert records["A3"]["data_type"] == "datetime"
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
            None,
            None,
        )
        assert wb["Data"]["A5"].hyperlink.target == "https://example.com/report"
        assert wb["Data"]["A6"].comment.text == "Native reader note"
        assert wb["Data"].column_dimensions["C"].width == 18
        assert len(list(wb["Data"].data_validations)) == 1
    finally:
        wb.close()
