"""Sprint Π constructor ratchet for formerly stubbed openpyxl paths."""

from __future__ import annotations

import importlib
from collections.abc import Callable
from typing import Any

import pytest


def _construct_dimension_holder(cls: type[Any]) -> Any:
    from wolfxl import Workbook

    return cls(Workbook().active)


def _construct_no_args(cls: type[Any]) -> Any:
    return cls()


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
