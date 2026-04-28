"""RFC-059 §2.3 (Sprint Ο Pod-1E): MergedCell + WriteOnlyCell.

Pins the cell-class compatibility shims so user code that does
``isinstance(cell, MergedCell)`` to detect non-anchor positions
inside merged ranges, or constructs ``WriteOnlyCell(value=42)``
to hand to ``ws.append([...])``, can migrate to wolfxl with a
one-line import swap.
"""

from __future__ import annotations

import pytest

from wolfxl.cell import MergedCell, WriteOnlyCell
from wolfxl.cell._merged import MergedCell as MergedCellDirect
from wolfxl.cell._write_only import WriteOnlyCell as WriteOnlyCellDirect


# ---------------------------------------------------------------------------
# MergedCell
# ---------------------------------------------------------------------------


def test_merged_cell_value_is_none() -> None:
    mc = MergedCell(parent=None, row=2, column=3)
    assert mc.value is None


def test_merged_cell_setter_raises_attribute_error() -> None:
    mc = MergedCell(parent=None, row=2, column=3)
    with pytest.raises(AttributeError, match="merged range"):
        mc.value = "anything"


def test_merged_cell_coordinate() -> None:
    mc = MergedCell(parent=None, row=3, column=2)
    assert mc.coordinate == "B3"
    assert mc.row == 3
    assert mc.column == 2


def test_merged_cell_reexport_path_matches_direct() -> None:
    """``wolfxl.cell.MergedCell`` and ``wolfxl.cell._merged.MergedCell``
    must be the same class object (re-export, not a copy)."""
    assert MergedCell is MergedCellDirect


# ---------------------------------------------------------------------------
# WriteOnlyCell
# ---------------------------------------------------------------------------


def test_write_only_cell_default_construction() -> None:
    wc = WriteOnlyCell()
    assert wc.parent is None
    assert wc.value is None
    assert wc.font is None
    assert wc.fill is None
    assert wc.number_format is None


def test_write_only_cell_value_and_style_passthrough() -> None:
    """Construction-time fields stick on the instance."""
    sentinel_font = object()
    sentinel_fill = object()
    wc = WriteOnlyCell(
        ws=None,
        value=42,
        font=sentinel_font,
        fill=sentinel_fill,
        number_format="0.00",
    )
    assert wc.value == 42
    assert wc.font is sentinel_font
    assert wc.fill is sentinel_fill
    assert wc.number_format == "0.00"


def test_write_only_cell_value_is_settable() -> None:
    """Unlike MergedCell, WriteOnlyCell.value is mutable."""
    wc = WriteOnlyCell(value="initial")
    assert wc.value == "initial"
    wc.value = "updated"
    assert wc.value == "updated"


def test_write_only_cell_reexport_path_matches_direct() -> None:
    assert WriteOnlyCell is WriteOnlyCellDirect
