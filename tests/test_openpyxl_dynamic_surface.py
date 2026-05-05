"""Dynamic public-surface ratchet against openpyxl core objects."""

from __future__ import annotations

import pytest

import wolfxl

openpyxl = pytest.importorskip("openpyxl")


@pytest.mark.parametrize(
    ("label", "openpyxl_obj", "wolfxl_obj"),
    [
        ("Workbook", openpyxl.Workbook(), wolfxl.Workbook()),
        ("Worksheet", openpyxl.Workbook().active, wolfxl.Workbook().active),
        ("Cell", openpyxl.Workbook().active["A1"], wolfxl.Workbook().active["A1"]),
    ],
)
def test_core_objects_expose_openpyxl_public_surface(
    label: str,
    openpyxl_obj: object,
    wolfxl_obj: object,
) -> None:
    """Core openpyxl objects should not gain untracked public attr gaps."""
    missing_callables: list[str] = []
    missing_values: list[str] = []
    type_mismatches: list[tuple[str, str]] = []

    for name in dir(openpyxl_obj):
        if name.startswith("_"):
            continue
        try:
            openpyxl_value = getattr(openpyxl_obj, name)
        except Exception:  # noqa: BLE001 - some descriptors depend on workbook state
            continue

        if not hasattr(wolfxl_obj, name):
            if callable(openpyxl_value):
                missing_callables.append(name)
            else:
                missing_values.append(name)
            continue

        wolfxl_value = getattr(wolfxl_obj, name)
        if callable(openpyxl_value) and not callable(wolfxl_value):
            type_mismatches.append((name, "openpyxl callable, wolfxl value"))
        elif not callable(openpyxl_value) and callable(wolfxl_value):
            type_mismatches.append((name, "openpyxl value, wolfxl callable"))

    assert missing_callables == [], f"{label} missing callables: {missing_callables}"
    assert missing_values == [], f"{label} missing values: {missing_values}"
    assert type_mismatches == [], f"{label} type mismatches: {type_mismatches}"
