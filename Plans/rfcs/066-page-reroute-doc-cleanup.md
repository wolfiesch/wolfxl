# RFC-066 — Re-routes + RFC-060 doc cleanup (Sprint Π Pod Π-ε)

> Closes 3 stubs: `wolfxl.worksheet.page.{PageMargins, PrintOptions, PrintPageSetup}`.
>
> Smallest pod in Sprint Π — pure re-routing + documentation. Can land first
> since it has zero dependencies on other Π pods.

## 1. Goal

Re-route 3 stubs to existing real implementations from Sprint-Ο Pod-1A.5,
and flip stale "stub" annotations in RFC-060 §12 / §12.1 to ✅.

## 2. Re-route work

### 2.1 `python/wolfxl/worksheet/page.py` (replaces stubs)

```python
"""``openpyxl.worksheet.page`` — page-margin / print-options / page-setup.

Re-exports real implementations from :mod:`wolfxl.worksheet.page_setup`
(landed by Sprint-Ο Pod-1A.5 / RFC-055).
"""

from __future__ import annotations

from wolfxl.worksheet.page_setup import (
    PageMargins,
    PageSetup,
    PrintOptions,
    PrintPageSetup,
)

__all__ = ["PageMargins", "PageSetup", "PrintOptions", "PrintPageSetup"]
```

Zero new logic — `page_setup.py` already implements all four classes
with full §10 dict contracts and 16 + 9 parity / diffwriter tests
passing from Pod-1A.5.

## 3. RFC-060 doc cleanup

### 3.1 `Plans/rfcs/060-class-reexport-shims.md §12`

Find every row containing "(stub)" and audit against current state.
Flip the following classes from "(stub)" → ✅ real:

- `wolfxl.worksheet.views.{SheetView, Pane, Selection, SheetViewList}` — landed by Sprint-Ο Pod-1A.5
- `wolfxl.worksheet.protection.SheetProtection` — landed by Sprint-Ο Pod-1A.5
- `wolfxl.worksheet.print_settings.{PrintArea, PrintTitles, ColRange, RowRange}` — landed by Sprint-Ο Pod-1A.5
- `wolfxl.worksheet.header_footer.HeaderFooter` — landed by Sprint-Ο Pod-1A.5
- `wolfxl.worksheet.page.{PageMargins, PrintOptions, PrintPageSetup}` — landed by Pod Π-ε (this RFC)

### 3.2 `Plans/rfcs/060-class-reexport-shims.md §12.1`

The 4-item "Tier-1.5 follow-up candidates" list is fully closed by
Sprint-Ο Pod-1A.5 + Sprint-Π. Replace with a closure note:

```markdown
### 12.1 Closure status

All Tier-1.5 follow-up candidates from the Sprint-Ο audit are closed
as of Sprint Π:

| Candidate | Closed by | SHA |
|---|---|---|
| SheetProtection | Sprint-Ο Pod-1A.5 | edcbb2c |
| PageMargins / PrintOptions / PrintPageSetup | Sprint-Π Pod-Π-ε | <fill at merge> |
| HeaderFooter / HeaderFooterItem | Sprint-Ο Pod-1A.5 | edcbb2c |
| SheetView / Pane / Selection | Sprint-Ο Pod-1A.5 | edcbb2c |

Sprint Π closes the remaining 23 Tier-1.5 stubs across Pods Π-α
(page breaks + dimensions), Π-β (merge + tables + copier), Π-γ
(NamedStyle + GradientFill + Protection + DifferentialStyle + Fill),
and Π-δ (workbook properties + internals).
```

## 4. Tests

`tests/parity/test_pod2_stubs_now_real.py` (NEW):

```python
"""Verify all 26 Sprint-Ο Pod-2 construction stubs are now real after Sprint Π."""

import pytest

# Format: (mod_path, symbol, no_arg_construction_works_or_None)
SPRINT_PI_REAL_STUBS = [
    # Pod Π-ε (this RFC)
    ("wolfxl.worksheet.page", "PageMargins", True),
    ("wolfxl.worksheet.page", "PrintOptions", True),
    ("wolfxl.worksheet.page", "PrintPageSetup", True),
    # Pod Π-α (RFC-062)
    ("wolfxl.worksheet.pagebreak", "Break", True),
    ("wolfxl.worksheet.pagebreak", "ColBreak", True),
    ("wolfxl.worksheet.pagebreak", "RowBreak", True),
    ("wolfxl.worksheet.dimensions", "DimensionHolder", False),  # needs worksheet arg
    ("wolfxl.worksheet.dimensions", "SheetFormatProperties", True),
    ("wolfxl.worksheet.dimensions", "SheetDimension", True),
    # Pod Π-β (RFC-063)
    ("wolfxl.worksheet.merge", "MergeCell", False),  # needs ref arg
    ("wolfxl.worksheet.merge", "MergeCells", True),
    ("wolfxl.worksheet.copier", "WorksheetCopy", False),  # needs source/target
    ("wolfxl.worksheet.table", "TableList", True),
    ("wolfxl.worksheet.table", "TablePartList", True),
    ("wolfxl.worksheet.table", "Related", True),
    ("wolfxl.worksheet.table", "XMLColumnProps", False),  # needs mapId/xpath
    # Pod Π-γ (RFC-064)
    ("wolfxl.styles", "NamedStyle", True),
    ("wolfxl.styles", "Protection", True),
    ("wolfxl.styles", "GradientFill", True),
    ("wolfxl.styles.differential", "DifferentialStyle", True),
    ("wolfxl.styles.fills", "Fill", True),  # abstract base; should construct empty
    # Pod Π-δ (RFC-065)
    ("wolfxl.workbook.child", "_WorkbookChild", True),
    ("wolfxl.workbook.properties", "CalcProperties", True),
    ("wolfxl.workbook.properties", "WorkbookProperties", True),
    ("wolfxl.comments.comments", "CommentSheet", True),
    ("wolfxl.drawing.spreadsheet_drawing", "SpreadsheetDrawing", True),
]


@pytest.mark.parametrize("mod_path,name,no_arg_works", SPRINT_PI_REAL_STUBS)
def test_stub_now_real(mod_path: str, name: str, no_arg_works: bool) -> None:
    import importlib
    mod = importlib.import_module(mod_path)
    cls = getattr(mod, name)
    if no_arg_works:
        # Must not raise NotImplementedError
        cls()
```

## 5. Tolerable pre-existing failures

None.

## 6. Calendar

½ day. Can land before any other Π pod since it has zero deps.
