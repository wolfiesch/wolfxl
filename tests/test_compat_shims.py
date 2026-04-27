"""T0 compat shim module smoke tests.

Two goals:
1. Every openpyxl module path exposed by wolfxl imports cleanly (no
   ``ModuleNotFoundError`` for the common paths).
2. Every shim class raises ``NotImplementedError`` with a helpful message
   at the construction site - not at downstream attribute access.

A drop-in replacement that silently no-ops would be far worse than a
pointed error. These tests pin the error behavior so we don't regress.
"""

from __future__ import annotations

import importlib

import pytest

import wolfxl

# ---------------- real class re-exports ----------------


def test_real_styles_classes_importable() -> None:
    from wolfxl.styles import Alignment, Border, Color, Font, PatternFill, Side

    assert Font(bold=True).bold is True
    assert PatternFill(patternType="solid", fgColor="FFFF0000").patternType == "solid"
    assert Color(rgb="FF00FF00").rgb == "FF00FF00"
    assert Border().left == Side()
    assert Alignment(horizontal="center").horizontal == "center"


def test_utils_cell_reexports() -> None:
    from wolfxl.utils.cell import (
        column_index_from_string,
        coordinate_to_tuple,
        get_column_letter,
        range_boundaries,
    )

    assert get_column_letter(1) == "A"
    assert column_index_from_string("AA") == 27
    assert coordinate_to_tuple("B3") == (3, 2)
    assert range_boundaries("A1:B2") == (1, 1, 2, 2)


def test_utils_cell_lazy_reexports() -> None:
    """Higher-level helpers routed through lazy ``__getattr__``."""
    from wolfxl.utils.cell import (
        absolute_coordinate,
        cols_from_range,
        get_column_interval,
        quote_sheetname,
        range_to_tuple,
        rows_from_range,
    )

    assert absolute_coordinate("A1") == "$A$1"
    assert quote_sheetname("Data")  # always quoted in openpyxl
    assert range_to_tuple("Sheet1!A1:B2") == ("Sheet1", (1, 1, 2, 2))
    assert list(rows_from_range("A1:B1")) == [("A1", "B1")]
    assert list(cols_from_range("A1:A2")) == [("A1", "A2")]
    assert get_column_interval("A", "C") == ["A", "B", "C"]


# ---------------- module imports ----------------


@pytest.mark.parametrize(
    "module_path",
    [
        "wolfxl.styles",
        "wolfxl.styles.named_styles",
        "wolfxl.styles.differential",
        "wolfxl.utils.cell",
        "wolfxl.utils.dataframe",
        "wolfxl.comments",
        "wolfxl.chart",
        "wolfxl.drawing",
        "wolfxl.drawing.image",
        "wolfxl.worksheet",
        "wolfxl.worksheet.datavalidation",
        "wolfxl.worksheet.table",
        "wolfxl.worksheet.filters",
        "wolfxl.worksheet.hyperlink",
        "wolfxl.formatting",
        "wolfxl.formatting.rule",
        "wolfxl.workbook",
        "wolfxl.workbook.defined_name",
        "wolfxl.pivot",
    ],
)
def test_module_imports(module_path: str) -> None:
    importlib.import_module(module_path)


# ---------------- stubs raise at construction ----------------


STUB_CONSTRUCTORS: list[tuple[str, str]] = [
    ("wolfxl.styles", "NamedStyle"),
    ("wolfxl.styles", "Protection"),
    ("wolfxl.styles", "GradientFill"),
    ("wolfxl.styles.differential", "DifferentialStyle"),
    # T1 PR1 promoted: Comment, Hyperlink now real dataclasses.
    # T1 PR2 promoted: DataValidation, Table, TableStyleInfo, *Rule now real.
    ("wolfxl.chart", "BarChart"),
    ("wolfxl.chart", "LineChart"),
    ("wolfxl.chart", "PieChart"),
    ("wolfxl.chart", "Reference"),
    ("wolfxl.chart", "Series"),
    # Sprint Λ Pod-β (RFC-045) promoted: ``wolfxl.drawing.image.Image``
    # is now a real class — exercised in tests/test_images_write.py
    # and tests/test_images_modify.py.
    ("wolfxl.worksheet.filters", "AutoFilter"),
    # T1 PR3 promoted: DefinedName now real dataclass.
    ("wolfxl.pivot", "PivotTable"),
]


# Classes promoted from stub -> real as T1 PRs land. Parametrized to catch
# accidental regression back to a raising stub.
REAL_DATACLASSES: list[tuple[str, str, dict]] = [
    ("wolfxl.comments", "Comment", {"text": "hello", "author": "me"}),
    ("wolfxl.worksheet.hyperlink", "Hyperlink", {"target": "https://example.com"}),
    ("wolfxl.worksheet.datavalidation", "DataValidation", {"type": "list", "formula1": '"a,b,c"'}),
    ("wolfxl.worksheet.table", "Table", {"name": "MyTable", "ref": "A1:B2"}),
    ("wolfxl.worksheet.table", "TableStyleInfo", {"name": "TableStyleLight9"}),
    ("wolfxl.formatting.rule", "CellIsRule", {"operator": "greaterThan", "formula": ["10"]}),
    ("wolfxl.formatting.rule", "FormulaRule", {"formula": ["$A1>100"]}),
    ("wolfxl.formatting.rule", "ColorScaleRule", {"start_type": "min", "end_type": "max"}),
    ("wolfxl.workbook.defined_name", "DefinedName", {"name": "Totals", "value": "Sheet1!$A$1"}),
]


@pytest.mark.parametrize("module_path,class_name,kwargs", REAL_DATACLASSES)
def test_real_dataclass_constructs_cleanly(
    module_path: str, class_name: str, kwargs: dict
) -> None:
    """Each promoted class constructs cleanly without raising.

    We don't compare every kwarg against an attribute of the same name —
    some classes (color-scale/data-bar rules) stash options into an
    ``extra`` dict rather than exposing each as a top-level attribute.
    The goal is to pin the ``NotImplementedError`` regression risk, not
    to fully spec each constructor's introspection surface.
    """
    mod = importlib.import_module(module_path)
    cls = getattr(mod, class_name)
    instance = cls(**kwargs)
    assert instance is not None
    # Best-effort attribute check: only probe names that ARE attributes.
    for key, value in kwargs.items():
        if hasattr(instance, key):
            assert getattr(instance, key) == value


@pytest.mark.parametrize("module_path,class_name", STUB_CONSTRUCTORS)
def test_stub_raises_on_construct(module_path: str, class_name: str) -> None:
    mod = importlib.import_module(module_path)
    cls = getattr(mod, class_name)
    with pytest.raises(NotImplementedError) as excinfo:
        cls()
    # Message contains the class name and our GitHub compatibility anchor.
    msg = str(excinfo.value)
    assert class_name in msg
    assert "wolfxl" in msg.lower()


# ---------------- Color theme/indexed support ----------------


def test_color_theme_roundtrip() -> None:
    c = wolfxl.Color(theme=1, tint=-0.3)
    assert c.theme == 1
    assert c.tint == -0.3
    assert c.rgb is None
    assert c.type == "theme"


def test_color_indexed() -> None:
    c = wolfxl.Color(indexed=3)
    assert c.indexed == 3
    assert c.rgb is None
    assert c.type == "indexed"
    # Indexed 3 is 00FF00 in the openpyxl COLOR_INDEX table.
    assert c.to_hex() == "#00FF00"


def test_color_rgb_default() -> None:
    c = wolfxl.Color()
    assert c.rgb == "00000000"
    assert c.theme is None
    assert c.indexed is None
    assert c.type == "rgb"


def test_color_is_hashable() -> None:
    """Frozen dataclass contract - Colors must be hashable for set/dict membership."""
    a = wolfxl.Color(rgb="FF0000")
    b = wolfxl.Color(rgb="FF0000")
    assert {a, b} == {a}


def test_dataframe_to_rows_without_pandas_import() -> None:
    """Importing the module does not require pandas."""
    import wolfxl.utils.dataframe as dfmod

    assert hasattr(dfmod, "dataframe_to_rows")


def test_dataframe_to_rows_basic() -> None:
    pd = pytest.importorskip("pandas")
    from wolfxl.utils.dataframe import dataframe_to_rows

    df = pd.DataFrame({"a": [1, 2], "b": [3, 4]})
    rows = list(dataframe_to_rows(df, index=False, header=True))
    assert rows[0] == ["a", "b"]
    assert rows[1] == [1, 3]
    assert rows[2] == [2, 4]
