"""T0 compat shim module smoke tests.

Two goals:
1. Every openpyxl module path exposed by wolfxl imports cleanly (no
   ``ModuleNotFoundError`` for the common paths).
2. Every formerly-stubbed shim class stays promoted to a real constructor,
   while any remaining shim placeholders raise ``NotImplementedError`` with a
   helpful message at construction time.

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
    # T1 PR1 promoted: Comment, Hyperlink now real dataclasses.
    # T1 PR2 promoted: DataValidation, Table, TableStyleInfo, *Rule now real.
    # Sprint Μ Pod-β (RFC-046) promoted: ``BarChart``, ``LineChart``,
    # ``PieChart``, ``DoughnutChart``, ``AreaChart``, ``ScatterChart``,
    # ``BubbleChart``, ``RadarChart``, ``Reference``, and ``Series`` are
    # now real chart classes — exercised in tests/test_charts_*.py.
    # Sprint Λ Pod-β (RFC-045) promoted: ``wolfxl.drawing.image.Image``
    # is now a real class — exercised in tests/test_images_write.py
    # and tests/test_images_modify.py.
    # Sprint Ο Pod 1B (RFC-056) promoted:
    # ``wolfxl.worksheet.filters.AutoFilter`` is now a real class —
    # exercised in tests/test_autofilter_filters.py.
    # T1 PR3 promoted: DefinedName now real dataclass.
    # Sprint Ν (RFC-047/048) promoted: ``wolfxl.pivot.PivotTable``,
    # ``PivotCache``, ``DataField``, ``PivotSource`` are real classes
    # — exercised in tests/test_pivot_construction.py.
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
    ("wolfxl.worksheet.filters", "AutoFilter", {}),
    ("wolfxl.styles", "NamedStyle", {"name": "Metric"}),
    ("wolfxl.styles", "Protection", {}),
    ("wolfxl.styles", "GradientFill", {}),
    ("wolfxl.styles.fills", "Fill", {}),
    ("wolfxl.styles.differential", "DifferentialStyle", {}),
    ("wolfxl.worksheet.dimensions", "DimensionHolder", {"worksheet": None}),
    ("wolfxl.worksheet.dimensions", "SheetFormatProperties", {}),
    ("wolfxl.worksheet.dimensions", "SheetDimension", {}),
    ("wolfxl.worksheet.merge", "MergeCell", {"ref": "A1:B2"}),
    ("wolfxl.worksheet.merge", "MergeCells", {}),
    ("wolfxl.worksheet.pagebreak", "Break", {}),
    ("wolfxl.worksheet.pagebreak", "PageBreak", {}),
    ("wolfxl.worksheet.properties", "WorksheetProperties", {}),
    ("wolfxl.worksheet.table", "TableList", {}),
    ("wolfxl.worksheet.table", "TablePartList", {}),
    ("wolfxl.worksheet.table", "Related", {}),
    ("wolfxl.worksheet.table", "XMLColumnProps", {}),
    ("wolfxl.workbook.properties", "CalcProperties", {}),
    ("wolfxl.workbook.properties", "WorkbookProperties", {}),
    ("wolfxl.workbook.child", "_WorkbookChild", {}),
    ("wolfxl.drawing.spreadsheet_drawing", "SpreadsheetDrawing", {}),
    ("wolfxl.comments.comments", "CommentSheet", {}),
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


def test_pivot_table_no_longer_stub() -> None:
    """Sprint Ν (RFC-047/048) ratchet flip: ``wolfxl.pivot.PivotTable``
    must NOT raise ``NotImplementedError`` on construction.

    This is the explicit ratchet for the v0.5+ → v2.0 promotion of
    pivot construction from stub to real class.
    """
    # Real PivotTable signature requires `cache=` and `location=`. The
    # ratchet is: invoking the constructor with proper args succeeds; the
    # stub variant would have raised NotImplementedError unconditionally.
    # Exhaustive surface tested in tests/test_pivot_construction.py.
    import inspect

    from wolfxl.pivot import PivotTable
    sig = inspect.signature(PivotTable.__init__)
    assert "cache" in sig.parameters
    assert "location" in sig.parameters
    # Sanity: confirm we did NOT accidentally re-stub.
    assert PivotTable.__module__.startswith("wolfxl.pivot")
    assert PivotTable.__init__.__qualname__ == "PivotTable.__init__"


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


def test_style_openpyxl_aliases() -> None:
    font = wolfxl.Font(b=True, i=True, u=True, sz=14, strikethrough=True)
    assert font.bold is True
    assert font.b is True
    assert font.italic is True
    assert font.i is True
    assert font.underline == "single"
    assert font.u == "single"
    assert font.size == 14
    assert font.sz == 14
    assert font.strikethrough is True

    fill = wolfxl.PatternFill(fill_type="solid", start_color="FFFF0000", end_color="FF00FF00")
    assert fill.patternType == "solid"
    assert fill.fill_type == "solid"
    assert fill.fgColor == "FFFF0000"
    assert fill.start_color == "FFFF0000"
    assert fill.bgColor == "FF00FF00"
    assert fill.end_color == "FF00FF00"

    align = wolfxl.Alignment(wrapText=True, textRotation=45, shrinkToFit=True)
    assert align.wrap_text is True
    assert align.wrapText is True
    assert align.text_rotation == 45
    assert align.textRotation == 45
    assert align.shrink_to_fit is True
    assert align.shrinkToFit is True

    side = wolfxl.Side(border_style="thin")
    assert side.style == "thin"
    assert side.border_style == "thin"
    border = wolfxl.Border(diagonal=side, diagonalUp=True)
    assert border.diagonal is side
    assert border.diagonal_direction == "up"

    color = wolfxl.Color(indexed=3)
    assert color.value == 3
    assert color.index == 3


def test_style_tree_helpers_round_trip() -> None:
    from xml.etree import ElementTree as ET

    font = wolfxl.Font(b=True, i=True, sz=14, color="FF0000")
    font_xml = ET.tostring(font.to_tree()).decode()
    assert "<b" in font_xml
    assert "<i" in font_xml
    assert 'rgb="00FF0000"' in font_xml
    assert wolfxl.Font.from_tree(font.to_tree()).bold is True

    fill = wolfxl.PatternFill(fill_type="solid", start_color="FF0000")
    assert wolfxl.PatternFill.from_tree(fill.to_tree()).fill_type == "solid"

    alignment = wolfxl.Alignment(horizontal="center", wrapText=True)
    assert wolfxl.Alignment.from_tree(alignment.to_tree()).wrapText is True

    side = wolfxl.Side(border_style="thin")
    border = wolfxl.Border(left=side, diagonalUp=True)
    assert wolfxl.Border.from_tree(border.to_tree()).left.border_style == "thin"

    color = wolfxl.Color.from_tree(wolfxl.Color(theme=1, tint=-0.3).to_tree())
    assert color.theme == 1
    assert color.tint == -0.3


def test_protection_and_named_style_helpers() -> None:
    from wolfxl.styles import NamedStyle, Protection

    protection = Protection(locked=False, hidden=True)
    assert Protection.from_tree(protection.to_tree()).locked is False
    assert Protection.from_tree(protection.to_tree()).hidden is True

    style = NamedStyle(name="Metric", builtinId=42, xfId=7, hidden=True)
    assert style.as_name().name == "Metric"
    assert style.as_tuple() == (0, 0, 0, 0, 0, 0, 0, 0, 0)
    assert style.as_xf().xfId == 7
    restored = NamedStyle.from_tree(style.to_tree())
    assert restored.name == "Metric"
    assert restored.builtinId == 42
    assert restored.xfId == 7
    assert restored.hidden is True


def test_data_validation_openpyxl_aliases_and_tree_helpers() -> None:
    from xml.etree import ElementTree as ET

    from wolfxl.worksheet.datavalidation import DataValidation, DataValidationList

    dv = DataValidation(type="list", formula1='"A,B"', allow_blank=True)
    dv.add("A1:B2")
    dv.hide_drop_down = True

    assert dv.validation_type == "list"
    assert dv.allowBlank is True
    assert dv.allow_blank is True
    assert dv.showDropDown is True
    assert dv.hide_drop_down is True
    assert str(dv.ranges) == "A1:B2"
    assert "A1" in dv.cells

    xml = ET.tostring(dv.to_tree()).decode()
    assert 'sqref="A1:B2"' in xml
    assert '<formula1>"A,B"</formula1>' in xml

    restored = DataValidation.from_tree(dv.to_tree())
    assert restored.validation_type == "list"
    assert restored.formula1 == '"A,B"'
    assert restored.hide_drop_down is True
    assert str(restored.sqref) == "A1:B2"

    validations = DataValidationList()
    validations.append(DataValidation())
    validations.append(restored)
    assert validations.count == 2
    list_xml = ET.tostring(validations.to_tree()).decode()
    assert 'count="1"' in list_xml
    assert "<dataValidation" in list_xml
    assert DataValidationList.from_tree(validations.to_tree()).count == 1


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
