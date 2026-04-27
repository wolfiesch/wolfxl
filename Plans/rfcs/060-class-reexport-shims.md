# RFC-060 — openpyxl-shaped class re-export shims

> **Status**: Approved
> **Phase**: 5 (2.0 — Sprint Ο)
> **Depends-on**: 055, 056, 057, 058, 059 (Pod 2 runs after Tier 1 pods)
> **Unblocks**: v2.0.0 launch
> **Pod**: 2

## 1. Goal

Mechanical but pedantic: ~70 openpyxl module paths from which
user code does `from openpyxl.X import Y` need to work as
`from wolfxl.X import Y`. The classes already exist (or are
created by Pods 1A-1E + 3); this pod creates the
openpyxl-shaped import paths that re-export them.

This is the highest-leverage Tier 2 work for "drop-in"
credibility: a user search-and-replace from `openpyxl` → `wolfxl`
should produce a working program.

## 2. Path map

### 2.1 worksheet.*

```python
# python/wolfxl/worksheet/cell_range.py
from ._cell_range import CellRange, MultiCellRange  # NEW class
__all__ = ["CellRange", "MultiCellRange"]

# python/wolfxl/worksheet/dimensions.py
from wolfxl._worksheet import (
    _ColumnDimension as ColumnDimension,
    _RowDimension as RowDimension,
    _DimensionHolder as DimensionHolder,
    _SheetFormatProperties as SheetFormatProperties,
    _SheetDimension as SheetDimension,
)
class Dimension: pass
__all__ = ["Dimension", "ColumnDimension", "RowDimension",
           "DimensionHolder", "SheetFormatProperties", "SheetDimension"]

# python/wolfxl/worksheet/merge.py
from ._merge import MergeCell, MergeCells, MergedCell, MergedCellRange

# python/wolfxl/worksheet/views.py — populated by Pod 1A
from .views import SheetView, Pane, Selection, SheetViewList

# python/wolfxl/worksheet/pagebreak.py — class re-exports
from ._pagebreak import Break, ColBreak, RowBreak, PageBreak
__all__ = ["Break", "ColBreak", "RowBreak", "PageBreak"]

# python/wolfxl/worksheet/print_settings.py — populated by Pod 1A
from .print_settings import PrintArea, PrintTitles, ColRange, RowRange

# python/wolfxl/worksheet/formula.py — populated by Pod 1C
from wolfxl.cell.cell import ArrayFormula, DataTableFormula

# python/wolfxl/worksheet/hyperlink.py
from ._hyperlink import HyperlinkList
from wolfxl.cell.hyperlink import Hyperlink

# python/wolfxl/worksheet/properties.py — populated by Pod 1A
from .properties import WorksheetProperties, PageSetupProperties, Outline

# python/wolfxl/worksheet/protection.py — populated by Pod 1A
from .protection import SheetProtection

# python/wolfxl/worksheet/table.py — extend existing
from ._table import (
    Table, TableColumn, TableStyleInfo,    # already exist
    TableList, TablePartList,               # NEW
    AutoFilter, SortState,                  # populated by Pod 1B
    Related, XMLColumnProps,                # NEW
)

# python/wolfxl/worksheet/page.py — populated by Pod 1A
from .page_setup import PageMargins, PrintOptions, PrintPageSetup

# python/wolfxl/worksheet/header_footer.py — populated by Pod 1A
from .header_footer import HeaderFooter, HeaderFooterItem

# python/wolfxl/worksheet/filters.py — populated by Pod 1B
from .filters import (
    BlankFilter, ColorFilter, CustomFilter, CustomFilters,
    DateGroupItem, DynamicFilter, IconFilter, NumberFilter,
    StringFilter, Top10, SortCondition, SortState,
    FilterColumn, AutoFilter,
)
```

### 2.2 cell.*

```python
# python/wolfxl/cell/cell.py — extended re-exports
from .hyperlink import Hyperlink
from ._merged import MergedCell                       # populated by Pod 1E
from ._write_only import WriteOnlyCell                # populated by Pod 1E
from wolfxl.utils.exceptions import IllegalCharacterError
from .rich_text import CellRichText
from wolfxl.cell.cell import ArrayFormula, DataTableFormula  # Pod 1C
__all__ = ["Cell", "Hyperlink", "MergedCell", "WriteOnlyCell",
           "IllegalCharacterError", "CellRichText",
           "ArrayFormula", "DataTableFormula", ...]
```

### 2.3 styles.*

```python
# python/wolfxl/styles/__init__.py — extend existing
from ._fill_base import Fill                # NEW abstract base
__all__ += ["Fill"]
```

### 2.4 formatting.rule.*

```python
# python/wolfxl/formatting/rule.py — NEW user-friendly classes
class ColorScale:
    """Color-scale CF rule with 2 or 3 stops."""
    def __init__(self, start_type="min", start_value=None, start_color=None,
                 mid_type=None, mid_value=None, mid_color=None,
                 end_type="max", end_value=None, end_color=None): ...
    def to_rust_dict(self) -> dict: ...

class DataBar:
    def __init__(self, color, min_type="min", min_value=None,
                 max_type="max", max_value=None, show_value=True): ...
    def to_rust_dict(self) -> dict: ...

class IconSet:
    def __init__(self, icon_set="3TrafficLights1", values=None,
                 reverse=False, show_value=True): ...
    def to_rust_dict(self) -> dict: ...

class DifferentialStyle:
    """Styling carried by CF rules' dxfId reference."""
    def __init__(self, font=None, fill=None, border=None,
                 alignment=None, number_format=None): ...

class RuleType:
    """Marker for parametrized CF rule kinds. Constants:
    AVERAGE, COLOR_SCALE, DATA_BAR, ICON_SET, FORMULA,
    EXPRESSION, DUPLICATE_VALUES, UNIQUE_VALUES,
    CONTAINS_TEXT, NOT_CONTAINS_TEXT, BEGINS_WITH, ENDS_WITH,
    CONTAINS_BLANKS, CONTAINS_NO_BLANKS,
    CONTAINS_ERRORS, CONTAINS_NO_ERRORS,
    TIME_PERIOD, ABOVE_AVERAGE, TOP10, CELL_IS."""
```

These wrap the existing dict-shape API so user code that
constructed an openpyxl `ColorScale` works as-is. Each
`.to_rust_dict()` produces the same shape as the existing
`add_conditional_formatting(...)` accepts.

### 2.5 utils.*

```python
# python/wolfxl/utils/cell.py — extend existing
from wolfxl.utils.exceptions import CellCoordinatesException
__all__ += ["CellCoordinatesException"]

# python/wolfxl/utils/indexed_list.py — populated by Pod 1E

# python/wolfxl/utils/exceptions.py — populated by Pod 1E
```

### 2.6 workbook.*

```python
# python/wolfxl/workbook/defined_name.py — extend existing
from ._defined_name_list import DefinedNameList   # NEW wrapper

# python/wolfxl/workbook/protection.py — populated by Pod 1D
from .protection import WorkbookProtection, FileSharing
```

## 3. CellRange class

The `wolfxl.worksheet.cell_range.CellRange` class is the most
substantial new addition. openpyxl's `CellRange` represents a
range like "A1:B10" with arithmetic ops:

```python
class CellRange:
    title: str | None = None
    min_col: int
    min_row: int
    max_col: int
    max_row: int

    def __init__(self, range_string=None, *,
                 min_col=None, min_row=None,
                 max_col=None, max_row=None, title=None): ...

    @property
    def coord(self) -> str: ...           # "A1:B10"
    @property
    def bounds(self) -> tuple[int,int,int,int]: ...
    @property
    def size(self) -> dict: ...           # {"rows": ..., "cols": ...}

    def expand(self, right=0, down=0, left=0, up=0) -> None: ...
    def shrink(self, right=0, bottom=0, left=0, top=0) -> None: ...
    def shift(self, col_shift=0, row_shift=0) -> None: ...
    def __contains__(self, coord: str) -> bool: ...
    def issubset(self, other: "CellRange") -> bool: ...
    def isdisjoint(self, other: "CellRange") -> bool: ...
    def union(self, other) -> "MultiCellRange": ...
    def intersection(self, other) -> "CellRange | None": ...
    def __eq__(self, other) -> bool: ...
    def __repr__(self) -> str: ...

class MultiCellRange:
    """Set-like collection of CellRange objects."""
    ranges: list[CellRange]
    def add(self, range): ...
    def remove(self, range): ...
    def __contains__(self, coord): ...
    def __iter__(self): ...
```

Many openpyxl APIs accept either a string or a `CellRange`
(merged_cells, sqref, conditional formatting ranges).
Wolfxl's existing string-based APIs continue to work; the
new `CellRange` class is accepted alongside via duck-typing
on `.coord`.

## 4. Drop-in test harness

```python
# tests/parity/test_dropin_imports.py
import pytest

DROP_IN_PAIRS = [
    # (openpyxl_path, symbol_name)
    ("openpyxl.worksheet.cell_range", "CellRange"),
    ("openpyxl.worksheet.cell_range", "MultiCellRange"),
    ("openpyxl.worksheet.dimensions", "ColumnDimension"),
    ("openpyxl.worksheet.dimensions", "RowDimension"),
    ("openpyxl.worksheet.dimensions", "DimensionHolder"),
    # ... (~70 entries)
]

@pytest.mark.parametrize("openpyxl_path,symbol", DROP_IN_PAIRS)
def test_wolfxl_provides_openpyxl_shaped_path(openpyxl_path, symbol):
    """Search-and-replace 'openpyxl' → 'wolfxl' in the import statement
    must produce a working import."""
    import importlib
    wolfxl_path = openpyxl_path.replace("openpyxl", "wolfxl", 1)
    mod = importlib.import_module(wolfxl_path)
    assert hasattr(mod, symbol), \
        f"{wolfxl_path}.{symbol} missing — drop-in claim broken"
    cls = getattr(mod, symbol)
    # Class must be instantiable with no args (or with same defaults
    # as openpyxl's class, modulo our duck-type contract).
```

## 5. Out of scope

Some openpyxl symbols aren't worth re-exporting:
- Internal descriptor primitives (`Bool`, `Integer`, `Typed`,
  `Sequence`) — user code never imports these.
- Private namespace constants (`SHEET_MAIN_NS`, regex patterns).
- stdlib re-exports (`BytesIO`, `chain`, `defaultdict`).

The audit script categorizes these as Tier 4 ("doesn't
matter"). Skip.

## 6. Testing

- `tests/parity/test_dropin_imports.py` — ~70 parametrized tests.
- `tests/test_cell_range.py` — full CellRange API (~30 tests).
- `tests/test_multi_cell_range.py` (~10 tests).
- `tests/test_color_scale_rule.py` (~10 tests).
- `tests/test_data_bar_rule.py` (~10 tests).
- `tests/test_icon_set_rule.py` (~10 tests).
- `tests/test_differential_style.py` (~6 tests).

## 10. Dict contract

CellRange does not have a dict contract — it's a pure-Python
class. The CF rule classes (ColorScale, DataBar, IconSet,
DifferentialStyle) emit dicts compatible with the existing
`add_conditional_formatting` API.

## 11. Acceptance

- `tests/parity/test_dropin_imports.py` 70/70 green.
- `from openpyxl.X import Y as Z` → `from wolfxl.X import Y as Z`
  works for every path in §2.
- `CellRange("A1:B10").bounds == (1, 1, 2, 10)`.
- `ColorScale(start_color="FFFF00", end_color="FF0000")` produces
  the same XML as openpyxl when fed to
  `ws.conditional_formatting.add(...)`.
