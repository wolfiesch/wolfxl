"""T1 PR6 — Workbook writes: DocumentProperties + DefinedName.

Write mode accepts ``wb.properties.title = ...`` and ``wb.defined_names[k] = DefinedName(...)``.
Modify mode now ships properties round-trip (RFC-020); ``defined_names`` is
still T1.5 (RFC-021) so its raise-contract test stays.
"""

from __future__ import annotations

from pathlib import Path

import pytest
from wolfxl.packaging.core import DocumentProperties
from wolfxl.workbook.defined_name import DefinedName

from wolfxl import Workbook

openpyxl = pytest.importorskip("openpyxl")


def test_properties_write_round_trip_via_openpyxl(tmp_path: Path) -> None:
    """wolfxl write → openpyxl read on the fields rust_xlsxwriter supports."""
    path = tmp_path / "wolfxl_props.xlsx"
    wb = Workbook()
    wb.properties.title = "Q3 Report"
    wb.properties.creator = "Alice"
    wb.properties.subject = "Revenue"
    wb.properties.keywords = "revenue, q3"
    wb.properties.description = "Quarterly revenue snapshot"
    wb.properties.category = "Finance"
    wb.save(str(path))

    op_wb = openpyxl.load_workbook(path)
    assert op_wb.properties.title == "Q3 Report"
    assert op_wb.properties.creator == "Alice"
    assert op_wb.properties.subject == "Revenue"
    assert op_wb.properties.keywords == "revenue, q3"
    assert op_wb.properties.description == "Quarterly revenue snapshot"
    assert op_wb.properties.category == "Finance"


def test_properties_setter_marks_dirty() -> None:
    wb = Workbook()
    wb.properties = DocumentProperties(title="replaced", creator="bob")
    assert wb._properties_dirty is True
    assert wb.properties.title == "replaced"
    assert wb.properties.creator == "bob"


def test_properties_write_round_trip_via_wolfxl(tmp_path: Path) -> None:
    """wolfxl write → wolfxl read (covers our own reader parity)."""
    path = tmp_path / "wolfxl_props_self.xlsx"
    wb = Workbook()
    wb.properties.title = "My Report"
    wb.properties.creator = "Me"
    wb.save(str(path))

    rb = Workbook._from_reader(str(path))
    assert rb.properties.title == "My Report"
    assert rb.properties.creator == "Me"


def test_defined_name_write_round_trip(tmp_path: Path) -> None:
    path = tmp_path / "wolfxl_names.xlsx"
    wb = Workbook()
    ws = wb.active
    for i in range(1, 6):
        ws[f"A{i}"] = i
    wb.defined_names["Region"] = DefinedName(
        name="Region",
        value="Sheet!$A$1:$A$5",
    )
    # In-session access works immediately.
    assert wb.defined_names["Region"].value == "Sheet!$A$1:$A$5"
    wb.save(str(path))

    op_wb = openpyxl.load_workbook(path)
    assert "Region" in op_wb.defined_names
    refers = op_wb.defined_names["Region"].value
    # rust_xlsxwriter may or may not add a leading "=" to refers_to — accept both.
    assert refers.lstrip("=").replace("'", "") in {
        "Sheet!$A$1:$A$5",
        "Sheet1!$A$1:$A$5",
    } or "A1:A5" in refers.replace("'", "")


def test_defined_name_add_helper_round_trip(tmp_path: Path) -> None:
    path = tmp_path / "wolfxl_names_add.xlsx"
    wb = Workbook()
    ws = wb.active
    ws["B1"] = 99
    wb.defined_names.add(DefinedName(name="Top", value="Sheet!$B$1"))
    wb.save(str(path))

    op_wb = openpyxl.load_workbook(path)
    assert "Top" in op_wb.defined_names


def test_defined_names_empty_workbook_saves(tmp_path: Path) -> None:
    """A workbook with no defined names should save without trouble."""
    path = tmp_path / "no_names.xlsx"
    wb = Workbook()
    wb.active["A1"] = 1
    wb.save(str(path))
    # Reopen — defined_names is empty dict-like.
    rb = Workbook._from_reader(str(path))
    assert len(rb.defined_names) == 0


def test_properties_modify_mode_round_trips(tmp_path: Path) -> None:
    """RFC-020: the patcher rewrites docProps/core.xml in modify mode.

    Replaces the prior T1.5-raise contract. The full positive-coverage
    surface lives in tests/test_modify_properties.py; this one just
    pins the contract that save() does NOT raise.
    """
    src = tmp_path / "exists_props.xlsx"
    out = tmp_path / "out_props.xlsx"
    openpyxl.Workbook().save(src)
    wb = Workbook._from_patcher(str(src))
    wb.properties.title = "new"
    wb.save(str(out))
    wb.close()
    # Re-open via openpyxl to verify the title round-trips.
    wb_check = openpyxl.load_workbook(str(out))
    assert wb_check.properties.title == "new"


def test_defined_names_modify_mode_raises(tmp_path: Path) -> None:
    path = tmp_path / "exists_names.xlsx"
    openpyxl.Workbook().save(path)
    wb = Workbook._from_patcher(str(path))
    wb.defined_names["New"] = DefinedName(name="New", value="Sheet!$A$1")
    with pytest.raises(NotImplementedError, match="T1.5"):
        wb.save(str(path))
    wb.close()
