# RFC-062 — Page breaks + dimensions (Sprint Π Pod Π-α)

> Closes 6 stubs: `Break`, `ColBreak`, `RowBreak` (page breaks);
> `DimensionHolder`, `SheetFormatProperties`, `SheetDimension` (dimensions).

## 1. Goal

Land Python construction + Rust emit + patcher Phase 2.5r drainage
for `<rowBreaks>` / `<colBreaks>` and the three dimension proxies.

## 2. Python class layer

### 2.1 `python/wolfxl/worksheet/pagebreak.py` (replaces stubs)

```python
from dataclasses import dataclass
from typing import Any

@dataclass
class Break:
    """Single page break (CT_Break, ECMA-376 §18.3.1.1)."""
    id: int = 0                  # 1-based row or column index
    min: int | None = None       # min cell in the break range
    max: int | None = None       # max cell in the break range
    man: bool = True             # manual break (vs auto)
    pt: bool = False             # printer-fitted break

    def to_rust_dict(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "min": self.min,
            "max": self.max,
            "man": self.man,
            "pt": self.pt,
        }


class RowBreak(Break):
    """Page break between rows. ``id`` is the row above the break."""


class ColBreak(Break):
    """Page break between columns. ``id`` is the column to the left of the break."""


PageBreak = RowBreak  # openpyxl alias


@dataclass
class PageBreakList:
    """Container — backs ``ws.row_breaks`` / ``ws.col_breaks``."""
    breaks: list[Break]
    count: int = 0
    manualBreakCount: int = 0  # noqa: N815

    def __init__(self) -> None:
        self.breaks = []

    def append(self, brk: Break) -> None:
        self.breaks.append(brk)
        self.count = len(self.breaks)
        self.manualBreakCount = sum(1 for b in self.breaks if b.man)

    def __len__(self) -> int:
        return len(self.breaks)

    def __iter__(self):
        return iter(self.breaks)

    def to_rust_dict(self) -> dict[str, Any]:
        return {
            "count": self.count,
            "manual_break_count": self.manualBreakCount,
            "breaks": [b.to_rust_dict() for b in self.breaks],
        }
```

### 2.2 `python/wolfxl/worksheet/dimensions.py` (replaces stubs)

```python
@dataclass
class DimensionHolder:
    """openpyxl-shape mapping wrapper over ``ws.row_dimensions`` /
    ``ws.column_dimensions``. Acts as a dict-like view; mutations
    flow back to the worksheet's existing dim proxies."""
    worksheet: Any
    default_factory: Any = None
    max_outline: int = 0

    def __getitem__(self, key): ...
    def __setitem__(self, key, value): ...
    def __iter__(self): ...
    def __len__(self): ...


@dataclass
class SheetFormatProperties:
    """`<sheetFormatPr>` defaults (CT_SheetFormatPr §18.3.1.81)."""
    baseColWidth: int = 8                # noqa: N815
    defaultColWidth: float | None = None # noqa: N815
    defaultRowHeight: float = 15.0       # noqa: N815
    customHeight: bool = False           # noqa: N815
    zeroHeight: bool = False             # noqa: N815
    thickTop: bool = False               # noqa: N815
    thickBottom: bool = False            # noqa: N815
    outlineLevelRow: int = 0             # noqa: N815
    outlineLevelCol: int = 0             # noqa: N815

    def to_rust_dict(self) -> dict[str, Any]: ...


@dataclass
class SheetDimension:
    """`<dimension ref="A1:Z100">`. Auto-computed by wolfxl from cell
    bounds; users can override via construction."""
    ref: str = "A1"

    def to_rust_dict(self) -> dict[str, Any]:
        return {"ref": self.ref}
```

## 3. Worksheet integration

```python
class Worksheet:
    @property
    def row_breaks(self) -> PageBreakList:
        if self._row_breaks is None:
            self._row_breaks = PageBreakList()
        return self._row_breaks

    @property
    def col_breaks(self) -> PageBreakList:
        if self._col_breaks is None:
            self._col_breaks = PageBreakList()
        return self._col_breaks

    @property
    def sheet_format(self) -> SheetFormatProperties:
        if self._sheet_format is None:
            self._sheet_format = SheetFormatProperties()
        return self._sheet_format

    @property
    def dimension_holder(self) -> DimensionHolder:
        return DimensionHolder(self)
```

New `__slots__` entries on Worksheet: `_row_breaks`, `_col_breaks`,
`_sheet_format`. Initialized to `None`.

## 4. Rust emit

`crates/wolfxl-writer/src/emit/page_breaks.rs` (new):

```rust
pub fn emit_row_breaks(spec: &PageBreakList) -> Vec<u8> { ... }
pub fn emit_col_breaks(spec: &PageBreakList) -> Vec<u8> { ... }
```

Wires into `sheet_xml::emit` AFTER `<headerFooter>`, BEFORE `<extLst>`.
ECMA-376 §18.3.1 child ordinal: rowBreaks=15, colBreaks=16.

## 5. Patcher Phase 2.5r

New module `src/wolfxl/page_breaks.rs`:

- `XlsxPatcher::queue_page_breaks_update(sheet, dict)` PyO3 method
- `apply_page_breaks_phase` runs AFTER Phase 2.5n (sheet-setup),
  BEFORE Phase 2.5p (slicers)
- Uses `wolfxl_merger::merge_blocks` with new SheetBlock variants
  `RowBreaks(bytes)` and `ColBreaks(bytes)` (added in Pod-α to
  `wolfxl-merger`).

## 6. Phase ordering

`2.5m → 2.5n → 2.5r (NEW) → 2.5p → 2.5o → 2.5q → 2.5h`

## 7. RFC-035 deep-clone

`crates/wolfxl-structural/src/sheet_copy.rs`: copy `_row_breaks`,
`_col_breaks`, `_sheet_format` slots when cloning.

## 8. §10 dict contracts

```python
ws.to_rust_page_breaks_dict() = {
    "row_breaks": {
        "count": 2,
        "manual_break_count": 2,
        "breaks": [{"id": 5, "min": 0, "max": 16383, "man": True, "pt": False}, ...],
    },
    "col_breaks": {...},  # same shape
}
ws.to_rust_sheet_format_dict() = {
    "base_col_width": 8,
    "default_col_width": None,
    "default_row_height": 15.0,
    ...
}
```

## 9. Tests

- `tests/test_page_breaks.py` — construction + attribute access (~30)
- `tests/test_dimension_helpers.py` — DimensionHolder/SheetFormatProperties/SheetDimension (~20)
- `tests/diffwriter/test_page_breaks_bytes.py` (`WOLFXL_TEST_EPOCH=0`)
- `tests/parity/test_page_breaks_parity.py` — round-trip via openpyxl

## 10. Tolerable pre-existing failures

None.
