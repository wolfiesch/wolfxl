"""T1 PR5 — Worksheet writes: tables, data_validations, conditional_formatting.

Write mode (``Workbook()`` → ``ws.add_table``, ``ws.data_validations.append``,
``ws.conditional_formatting.add`` → ``wb.save``) must produce a valid xlsx
that openpyxl can read with all three features intact. Modify mode raises
``NotImplementedError`` with a T1.5 pointer.
"""

from __future__ import annotations

from pathlib import Path

import pytest
from wolfxl.formatting.rule import CellIsRule
from wolfxl.worksheet.datavalidation import DataValidation
from wolfxl.worksheet.table import Table, TableColumn, TableStyleInfo

from wolfxl import Workbook

openpyxl = pytest.importorskip("openpyxl")


def test_add_table_round_trip(tmp_path: Path) -> None:
    path = tmp_path / "wolfxl_table.xlsx"
    wb = Workbook()
    ws = wb.active
    # Build a minimal data block for the table to cover.
    headers = ["Name", "Amount"]
    ws.append(headers)
    for i in range(1, 4):
        ws.append([f"R{i}", i * 10])

    t = Table(
        name="Sales",
        displayName="Sales",
        ref="A1:B4",
        tableStyleInfo=TableStyleInfo(name="TableStyleLight9"),
        tableColumns=[TableColumn(id=1, name="Name"), TableColumn(id=2, name="Amount")],
    )
    ws.add_table(t)
    # Pre-save: ws.tables already knows about it.
    assert "Sales" in ws.tables
    assert ws.tables["Sales"].ref == "A1:B4"
    wb.save(str(path))

    op_wb = openpyxl.load_workbook(path)
    op_ws = op_wb.active
    assert "Sales" in op_ws.tables
    assert op_ws.tables["Sales"].ref == "A1:B4"


def test_data_validation_append_round_trip(tmp_path: Path) -> None:
    path = tmp_path / "wolfxl_dv.xlsx"
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "Pick"
    dv = DataValidation(
        type="list",
        formula1='"Red,Blue,Green"',
        sqref="A2:A10",
        allowBlank=True,
    )
    ws.data_validations.append(dv)
    # Pre-save visibility.
    assert len(ws.data_validations) == 1
    wb.save(str(path))

    op_wb = openpyxl.load_workbook(path)
    op_ws = op_wb.active
    assert len(op_ws.data_validations.dataValidation) == 1
    op_dv = op_ws.data_validations.dataValidation[0]
    assert op_dv.type == "list"
    assert op_dv.formula1 == '"Red,Blue,Green"'


def test_conditional_formatting_add_round_trip(tmp_path: Path) -> None:
    path = tmp_path / "wolfxl_cf.xlsx"
    wb = Workbook()
    ws = wb.active
    for i in range(1, 11):
        ws[f"B{i}"] = i * 5
    rule = CellIsRule(operator="greaterThan", formula=["50"])
    ws.conditional_formatting.add("B2:B10", rule)
    assert len(ws.conditional_formatting) == 1
    wb.save(str(path))

    op_wb = openpyxl.load_workbook(path)
    op_ws = op_wb.active
    entries = list(op_ws.conditional_formatting)
    assert any(
        any(getattr(r, "type", None) == "cellIs" for r in e.rules)
        for e in entries
    ), entries


def test_add_table_rejects_non_table() -> None:
    wb = Workbook()
    ws = wb.active
    with pytest.raises(TypeError, match="Table"):
        ws.add_table("not a table")  # type: ignore[arg-type]


def test_modify_mode_raises_on_add_table(tmp_path: Path) -> None:
    path = tmp_path / "exists.xlsx"
    openpyxl.Workbook().save(path)

    wb = Workbook._from_patcher(str(path))
    ws = wb.active
    t = Table(name="X", ref="A1:B2")
    with pytest.raises(NotImplementedError, match="T1.5"):
        ws.add_table(t)
    wb.close()


def test_modify_mode_raises_on_dv_append(tmp_path: Path) -> None:
    path = tmp_path / "exists_dv.xlsx"
    openpyxl.Workbook().save(path)

    wb = Workbook._from_patcher(str(path))
    ws = wb.active
    dv = DataValidation(type="list", formula1='"a,b"', sqref="A1:A2")
    with pytest.raises(NotImplementedError, match="T1.5"):
        ws.data_validations.append(dv)
    wb.close()


def test_modify_mode_raises_on_cf_add(tmp_path: Path) -> None:
    path = tmp_path / "exists_cf.xlsx"
    openpyxl.Workbook().save(path)

    wb = Workbook._from_patcher(str(path))
    ws = wb.active
    with pytest.raises(NotImplementedError, match="T1.5"):
        ws.conditional_formatting.add("A1:A10", CellIsRule(operator="equal", formula=["1"]))
    wb.close()
