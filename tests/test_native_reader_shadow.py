"""Native reader shadow comparisons against the current default reader."""

from __future__ import annotations

import datetime as dt
from pathlib import Path
from typing import Any

import pytest

openpyxl = pytest.importorskip("openpyxl")
openpyxl_datavalidation = pytest.importorskip("openpyxl.worksheet.datavalidation")
openpyxl_hyperlink = pytest.importorskip("openpyxl.worksheet.hyperlink")
openpyxl_table = pytest.importorskip("openpyxl.worksheet.table")
wolfxl = pytest.importorskip("wolfxl")

DataValidation = openpyxl_datavalidation.DataValidation
Hyperlink = openpyxl_hyperlink.Hyperlink
Table = openpyxl_table.Table
TableStyleInfo = openpyxl_table.TableStyleInfo


def _make_shadow_xlsx(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Main"
    ws["A1"] = "label"
    ws["B1"] = "amount"
    ws["A2"] = "alpha"
    ws["B2"] = 1234.5
    ws["B2"].number_format = "#,##0.00"
    ws["A3"] = "when"
    ws["B3"] = dt.datetime(2024, 3, 5, 9, 45)
    ws["A4"] = "formula"
    ws["B4"] = "=B2*2"
    ws["A5"] = "link"
    ws["A5"].hyperlink = Hyperlink(
        ref="A5",
        target="https://example.com/shadow",
        tooltip="Shadow link",
    )
    ws["B5"] = "jump"
    ws["B5"].hyperlink = Hyperlink(ref="B5", location="Other!A1", display="Other")
    ws["A6"] = "note"
    ws["A6"].comment = openpyxl.comments.Comment("Shadow note", "Wolf")
    ws.row_dimensions[6].height = 24
    ws.column_dimensions["C"].width = 18
    dv = DataValidation(type="list", formula1='"Red,Blue"', allow_blank=True)
    dv.add("C2:C6")
    ws.add_data_validation(dv)
    table = Table(displayName="ShadowTable", ref="A1:B4")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
    ws.add_table(table)
    ws.merge_cells("D1:E1")
    ws.freeze_panes = "B2"

    other = wb.create_sheet("Other")
    other["A1"] = True
    other["B1"] = "tail"

    wb.save(path)
    wb.close()


def _workbook_snapshot(wb: Any) -> dict[str, Any]:
    out: dict[str, Any] = {"sheetnames": list(wb.sheetnames), "sheets": {}}
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        values = list(ws.iter_rows(values_only=True))
        number_formats = {
            coord: ws[coord].number_format
            for coord in ("B2", "B3")
            if ws.max_row >= int(coord[1:])
        }
        merged = {str(r) for r in ws.merged_cells.ranges}
        hyperlinks = {
            coord: (
                ws[coord].hyperlink.target,
                ws[coord].hyperlink.location,
                ws[coord].hyperlink.display,
                ws[coord].hyperlink.tooltip,
            )
            for coord in ("A5", "B5")
            if ws[coord].hyperlink is not None
        }
        comments = {
            coord: (ws[coord].comment.text, ws[coord].comment.author)
            for coord in ("A6",)
            if ws[coord].comment is not None
        }
        validations = [
            (dv.type, dv.formula1, dv.allowBlank, str(dv.sqref))
            for dv in ws.data_validations
        ]
        tables = {
            name: (
                table.ref,
                [column.name for column in table.tableColumns],
                table.tableStyleInfo.name if table.tableStyleInfo else None,
            )
            for name, table in ws.tables.items()
        }
        out["sheets"][sheet_name] = {
            "values": values,
            "number_formats": number_formats,
            "merged": merged,
            "hyperlinks": hyperlinks,
            "comments": comments,
            "freeze_panes": ws.freeze_panes,
            "row_height": ws.row_dimensions[6].height,
            "column_width": ws.column_dimensions["C"].width,
            "data_validations": validations,
            "tables": tables,
        }
    return out


def test_native_reader_shadow_matches_default_reader(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    path = tmp_path / "shadow.xlsx"
    _make_shadow_xlsx(path)

    monkeypatch.delenv("WOLFXL_NATIVE_READER", raising=False)
    default = wolfxl.load_workbook(path, data_only=False)
    try:
        default_snapshot = _workbook_snapshot(default)
    finally:
        default.close()

    monkeypatch.setenv("WOLFXL_NATIVE_READER", "1")
    native = wolfxl.load_workbook(path, data_only=False)
    try:
        assert native._rust_reader.__class__.__name__ == "NativeXlsxBook"  # noqa: SLF001
        native_snapshot = _workbook_snapshot(native)
    finally:
        native.close()

    assert native_snapshot == default_snapshot
