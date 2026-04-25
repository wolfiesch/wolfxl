"""Data-validation cases.

Two fixtures here:

1. ``data_validation_list_literal`` — happy path with ``allow_blank: True``
   set explicitly. Pins that both backends produce identical XML for the
   common case.

2. ``data_validation_omit_allow_blank`` — W4E.P3 regression. The original
   review flagged that native unconditionally defaulted ``allow_blank``
   to ``True`` while oracle left the OOXML default in place. This case
   builds a list-DV without the key set and asserts cross-backend
   structural parity through the harness Layer 2 + Layer 3 gates. If
   either backend changes its default, the harness flags it loudly
   instead of letting it drift silently.
"""
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


def _build_dv_omit_allow_blank(wb: Any) -> None:
    """W4E.P3: omit the ``allow_blank`` key entirely. Both backends must
    pick the same default — the harness reports any divergence."""
    ws = wb.active
    ws["A1"] = "pick one"
    w = wb._rust_writer
    w.add_data_validation(ws.title, {
        "range": "B1:B10",
        "validation_type": "list",
        "formula1": '"Red,Green,Blue"',
    })


CASES = [
    ("data_validation_list_literal", _build_dv_list),
    ("data_validation_omit_allow_blank", _build_dv_omit_allow_blank),
]
