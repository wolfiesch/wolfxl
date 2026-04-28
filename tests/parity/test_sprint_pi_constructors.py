"""Sprint Π constructor ratchet for formerly stubbed openpyxl paths."""

from __future__ import annotations

import importlib
from collections.abc import Callable
from typing import Any

import pytest


def _construct_dimension_holder(cls: type[Any]) -> Any:
    from wolfxl import Workbook

    return cls(Workbook().active)


def _construct_merge_cell(cls: type[Any]) -> Any:
    return cls("A1:B2")


def _construct_no_args(cls: type[Any]) -> Any:
    return cls()


def _construct_worksheet_copy(cls: type[Any]) -> Any:
    from wolfxl import Workbook

    wb = Workbook()
    source = wb.active
    target = type("TargetSheet", (), {"title": "Copied"})()
    return cls(source, target)


# Only include pods that have landed on feat/native-writer. Later Sprint Π pods
# should append their symbols here as they replace the remaining stubs.
SPRINT_PI_LANDED_CONSTRUCTORS: tuple[
    tuple[str, str, Callable[[type[Any]], Any]],
    ...,
] = (
    # RFC-066 / Π-epsilon: re-route to existing real page_setup classes.
    ("wolfxl.worksheet.page", "PageMargins", _construct_no_args),
    ("wolfxl.worksheet.page", "PrintOptions", _construct_no_args),
    ("wolfxl.worksheet.page", "PrintPageSetup", _construct_no_args),
    # RFC-062 / Π-alpha: page breaks + dimensions.
    ("wolfxl.worksheet.pagebreak", "Break", _construct_no_args),
    ("wolfxl.worksheet.pagebreak", "ColBreak", _construct_no_args),
    ("wolfxl.worksheet.pagebreak", "RowBreak", _construct_no_args),
    ("wolfxl.worksheet.dimensions", "DimensionHolder", _construct_dimension_holder),
    ("wolfxl.worksheet.dimensions", "SheetFormatProperties", _construct_no_args),
    ("wolfxl.worksheet.dimensions", "SheetDimension", _construct_no_args),
    # RFC-063 / Π-beta: merge, table-list, and copier support types.
    ("wolfxl.worksheet.merge", "MergeCell", _construct_merge_cell),
    ("wolfxl.worksheet.merge", "MergeCells", _construct_no_args),
    ("wolfxl.worksheet.copier", "WorksheetCopy", _construct_worksheet_copy),
    ("wolfxl.worksheet.table", "TableList", _construct_no_args),
    ("wolfxl.worksheet.table", "TablePartList", _construct_no_args),
    ("wolfxl.worksheet.table", "Related", _construct_no_args),
    ("wolfxl.worksheet.table", "XMLColumnProps", _construct_no_args),
    # RFC-065 / Π-delta: workbook calculation + workbook properties.
    ("wolfxl.workbook.properties", "CalcProperties", _construct_no_args),
    ("wolfxl.workbook.properties", "WorkbookProperties", _construct_no_args),
)


@pytest.mark.parametrize(
    ("module_path", "symbol_name", "factory"),
    SPRINT_PI_LANDED_CONSTRUCTORS,
)
def test_landed_sprint_pi_constructors_are_not_stubs(
    module_path: str,
    symbol_name: str,
    factory: Callable[[type[Any]], Any],
) -> None:
    module = importlib.import_module(module_path)
    cls = getattr(module, symbol_name)

    try:
        instance = factory(cls)
    except NotImplementedError as exc:  # pragma: no cover - regression path
        pytest.fail(f"{module_path}.{symbol_name} still raises NotImplementedError: {exc}")

    assert instance is not None


def test_mergecells_unbound_container_is_set_like() -> None:
    from wolfxl.worksheet.merge import MergeCell, MergeCells

    merges = MergeCells(mergeCell=["A1:B2", MergeCell("C3:D4"), "A1:B2"])

    assert merges.count == 2
    assert "A1:B2" in merges
    assert [cell.ref for cell in merges] == ["A1:B2", "C3:D4"]

    merges.remove("A1:B2")

    assert merges.count == 1
    assert [cell.coord for cell in merges.mergeCell] == ["C3:D4"]


def test_mergecells_bound_container_updates_worksheet_ranges() -> None:
    from wolfxl.worksheet.merge import MergeCells

    from wolfxl import Workbook

    ws = Workbook().active
    merges = MergeCells(ws)

    merges.append("B2:C3")

    assert "B2:C3" in ws._merged_ranges  # noqa: SLF001
    assert [cell.ref for cell in merges] == ["B2:C3"]


def test_table_list_unbound_and_table_part_helpers() -> None:
    from wolfxl.worksheet.table import Related, Table, TableList, TablePartList, XMLColumnProps

    tables = TableList()
    table = Table(name="Sales", ref="A1:B2")

    tables.add(table)

    assert len(tables) == 1
    assert "Sales" in tables
    assert tables["Sales"] is table
    assert tables.items() == [("Sales", table)]

    parts = TablePartList(tablePart=[Related("rId1")])
    parts.append(Related("rId2"))

    assert parts.count == 2
    assert [part.id for part in parts] == ["rId1", "rId2"]
    assert XMLColumnProps(mapId=7, xpath="/root/item").mapId == 7


def test_table_list_bound_container_queues_tables_on_worksheet() -> None:
    from wolfxl.worksheet.table import Table, TableList

    from wolfxl import Workbook

    ws = Workbook().active
    table = Table(name="Sales", ref="A1:B2")

    tables = TableList(ws)
    tables.add(table)

    assert tables["Sales"] is table
    assert ws.tables["Sales"] is table
    assert table in ws._pending_tables  # noqa: SLF001


def test_worksheetcopy_delegates_to_workbook_copy_worksheet() -> None:
    from wolfxl.worksheet.copier import WorksheetCopy

    from wolfxl import Workbook

    wb = Workbook()
    source = wb.active
    source["A1"] = "copied"
    source.merge_cells("B2:C3")
    target = type("TargetSheet", (), {"title": "Copied"})()

    copied = WorksheetCopy(source, target).copy_worksheet()

    assert copied.title == "Copied"
    assert copied["A1"].value == "copied"
    assert "B2:C3" in copied._merged_ranges  # noqa: SLF001


def test_workbook_property_dataclasses_export_rust_contract() -> None:
    from wolfxl.workbook.properties import CalcProperties, WorkbookProperties

    calc = CalcProperties(calcId=191029, calcMode="manual", forceFullCalc=True)
    props = WorkbookProperties(date1904=True, codeName="ThisWorkbook")

    assert calc.to_rust_dict()["calc_id"] == 191029
    assert calc.to_rust_dict()["calc_mode"] == "manual"
    assert calc.to_rust_dict()["force_full_calc"] is True
    assert props.to_rust_dict()["date1904"] is True
    assert props.to_rust_dict()["code_name"] == "ThisWorkbook"
