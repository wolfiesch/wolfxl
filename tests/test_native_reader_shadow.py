"""Native reader shadow comparisons against the current default reader."""

from __future__ import annotations

import datetime as dt
from pathlib import Path
from typing import Any

import pytest

openpyxl = pytest.importorskip("openpyxl")
wolfxl = pytest.importorskip("wolfxl")


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
    ws.merge_cells("D1:E1")

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
        out["sheets"][sheet_name] = {
            "values": values,
            "number_formats": number_formats,
            "merged": merged,
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
