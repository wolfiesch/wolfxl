"""Data-validation case — list with literal value set."""
from __future__ import annotations

from typing import Any


def _build_dv_list(wb: Any) -> None:
    ws = wb.active
    ws["A1"] = "pick one"
    w = wb._rust_writer
    w.add_data_validation(ws.title, {
        "range": "B1:B10",
        "validation_type": "list",
        "formula1": '"Red,Green,Blue"',
        "allow_blank": True,
    })


CASES = [
    ("data_validation_list_literal", _build_dv_list),
]
