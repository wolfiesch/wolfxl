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
openpyxl_hyperlink = pytest.importorskip("openpyxl.worksheet.hyperlink")
wolfxl = pytest.importorskip("wolfxl")

Hyperlink = openpyxl_hyperlink.Hyperlink


def _make_basic_xlsx(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"] = "Label"
    ws["B1"] = 42
    ws["B1"].number_format = "#,##0.00"
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
    finally:
        wb.close()
