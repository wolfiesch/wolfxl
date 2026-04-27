# RFC-055 — Print / view / sheet protection

> **Status**: Approved
> **Phase**: 5 (2.0 — Sprint Ο)
> **Depends-on**: 010 (rels), 011 (xml-merger), 013 (patcher extensions), 035 (copy_worksheet)
> **Unblocks**: 060 (re-export shims)
> **Pod**: 1A

## 1. Goal

Close the largest user-visible Tier 1 gap surfaced by the `from
openpyxl.worksheet.X import Y` audit: `ws.page_setup`,
`ws.page_margins`, `ws.header_footer`, `ws.print_title_rows`,
`ws.print_title_cols`, `ws.sheet_view`, `ws.protection` all raise
`AttributeError` today. Excel's print-preview / view-preference
/ sheet-protection features are first-class user-facing surface;
"full openpyxl replacement" is not credible without them.

## 2. Public API

### 2.1 PageSetup

```python
class PageSetup:
    orientation: Literal["default", "portrait", "landscape"] = "default"
    paper_size: int | None = None  # 1=Letter, 9=A4, … (118 enum values)
    fit_to_width: int | None = None
    fit_to_height: int | None = None
    scale: int | None = None  # 10..400
    first_page_number: int | None = None
    horizontal_dpi: int | None = None
    vertical_dpi: int | None = None
    cell_comments: Literal["asDisplayed", "atEnd", "none"] | None = None
    errors: Literal["displayed", "blank", "dash", "NA"] | None = None
    use_first_page_number: bool | None = None
    use_printer_defaults: bool | None = None
    black_and_white: bool | None = None
    draft: bool | None = None
```

Accessor: `ws.page_setup` returns the singleton instance attached
to the worksheet (lazy-initialized; mutations are observable).

### 2.2 PageMargins

```python
class PageMargins:
    top: float = 0.75       # inches
    bottom: float = 0.75
    left: float = 0.7
    right: float = 0.7
    header: float = 0.3
    footer: float = 0.3
```

Accessor: `ws.page_margins`. All values in inches per OOXML
contract.

### 2.3 HeaderFooter

```python
class HeaderFooterItem:
    left: str | None = None
    center: str | None = None
    right: str | None = None

class HeaderFooter:
    odd_header: HeaderFooterItem
    odd_footer: HeaderFooterItem
    even_header: HeaderFooterItem
    even_footer: HeaderFooterItem
    first_header: HeaderFooterItem
    first_footer: HeaderFooterItem
    different_odd_even: bool = False
    different_first: bool = False
    scale_with_doc: bool = True
    align_with_margins: bool = True
```

Accessor: `ws.header_footer`. Header/footer text supports the
OOXML format-code grammar:

| Code | Meaning |
|---|---|
| `&L` / `&C` / `&R` | Left / center / right alignment switch |
| `&P` | Page number |
| `&N` | Total page count |
| `&D` | Date |
| `&T` | Time |
| `&F` | File name |
| `&A` | Sheet name |
| `&Z` | File path |
| `&G` | Picture (out of scope for v2.0 — round-trips bytes only) |
| `&"font,style"` | Font select |
| `&NN` | Font size (NN integer) |
| `&K{RRGGBB}` | Font color hex |
| `&B` / `&I` / `&U` | Bold / italic / underline toggle |
| `&S` | Strikethrough toggle |
| `&X` / `&Y` | Superscript / subscript |
| `&&` | Literal ampersand |

Wolfxl validates these on emit; unknown `&X` codes pass through
unchanged for forward compat.

### 2.4 print_title_rows / print_title_cols

```python
ws.print_title_rows = "1:2"      # repeat rows 1-2
ws.print_title_cols = "A:B"      # repeat cols A-B
ws.print_title_rows = None       # clear
```

Stored as a workbook-level `<definedName>` per OOXML:
`<definedName name="_xlnm.Print_Titles" localSheetId="N">SheetName!$1:$2,SheetName!$A:$B</definedName>`.

Pod 1A composes with RFC-021 (defined-names mutation) — the
emitted definedName uses RFC-021's queue path.

### 2.5 SheetView

```python
class Pane:
    x_split: float = 0.0
    y_split: float = 0.0
    top_left_cell: str = "A1"
    active_pane: Literal["bottomLeft", "bottomRight", "topLeft", "topRight"] = "topLeft"
    state: Literal["frozen", "split", "frozenSplit"] = "frozen"

class Selection:
    active_cell: str = "A1"
    sqref: str = "A1"
    pane: Literal["bottomLeft", "bottomRight", "topLeft", "topRight"] | None = None

class SheetView:
    zoom_scale: int = 100        # 10..400
    zoom_scale_normal: int = 100
    view: Literal["normal", "pageBreakPreview", "pageLayout"] = "normal"
    show_grid_lines: bool = True
    show_row_col_headers: bool = True
    show_outline_symbols: bool = True
    show_zeros: bool = True
    right_to_left: bool = False
    tab_selected: bool = False
    top_left_cell: str | None = None
    pane: Pane | None = None
    selection: list[Selection] = []
```

Accessor: `ws.sheet_view`. The existing `ws.freeze_panes = "B2"`
shim continues to work — mutations route through
`sheet_view.pane`.

### 2.6 SheetProtection

```python
class SheetProtection:
    sheet: bool = False                 # toggles on the protection
    objects: bool = False
    scenarios: bool = False
    format_cells: bool = True
    format_columns: bool = True
    format_rows: bool = True
    insert_columns: bool = True
    insert_rows: bool = True
    insert_hyperlinks: bool = True
    delete_columns: bool = True
    delete_rows: bool = True
    select_locked_cells: bool = False
    sort: bool = True
    auto_filter: bool = True
    pivot_tables: bool = True
    select_unlocked_cells: bool = False
    password: str | None = None         # hashed via openpyxl algo

    def set_password(self, password: str) -> None: ...
    def check_password(self, password: str) -> bool: ...
```

Accessor: `ws.protection`. Password hashing reuses the helper
from Sprint Ι Pod-γ (`wolfxl.utils.protection.hash_password`).

## 3. OOXML output

### 3.1 Worksheet child order (CT_Worksheet)

```xml
<worksheet>
  <sheetPr/>
  <dimension/>
  <sheetViews/>      <!-- §2.5 -->
  <sheetFormatPr/>
  <cols/>
  <sheetData/>
  ...
  <sheetProtection/> <!-- §2.6 (BEFORE autoFilter) -->
  <autoFilter/>      <!-- existing RFC-024 + Pod 1B -->
  ...
  <pageMargins/>     <!-- §2.2 -->
  <pageSetup/>       <!-- §2.1 -->
  <headerFooter/>    <!-- §2.3 -->
  ...
</worksheet>
```

Wolfxl-merger gains four new `SheetBlock` variants:
`SheetViews`, `SheetProtection`, `PageMargins`, `PageSetup`,
`HeaderFooter`. Each placed at the canonical position above.

### 3.2 print_title_rows

The `_xlnm.Print_Titles` definedName lives in `xl/workbook.xml`
and is sheet-scoped (`localSheetId` set). Routed through
existing RFC-021 defined-names queue.

## 4. Modify mode

XlsxPatcher Phase 2.5n drains queued sheet-setup mutations:

1. For each sheet with mutations, read existing sheet XML.
2. Compose patches via wolfxl-merger using the new SheetBlock
   variants — replace the existing block (idempotent) or insert
   at the canonical position if missing.
3. Route bytes back to file_patches.

Phase 2.5n runs after Phase 2.5m (pivots) and before Phase 3
(cells) so users can mutate setup + cells in the same save.

## 5. Native writer

`crates/wolfxl-writer/src/emit/sheet_setup.rs` (~400 LOC).
Extends `WorksheetData` with the six new fields; emit functions
write the canonical XML in CT_Worksheet order.

## 6. RFC-035 deep-clone

Page setup, page margins, header/footer, sheet views are all
sheet-scoped. Today these aren't deep-cloned because they didn't
have a Python API; with Pod 1A they do. Deep-clone the relevant
XML blocks via the wolfxl-merger primitive.

`print_title_rows` / `print_title_cols` create workbook-scoped
`<definedName>` entries — RFC-035 §3 OQ-c already handles
sheet-scoped definedNames for a cloned sheet.

## 7. Testing

- `tests/test_page_setup.py` (~12 tests).
- `tests/test_page_margins.py` (~6 tests).
- `tests/test_header_footer.py` (~14 tests including format-code grammar).
- `tests/test_print_titles.py` (~8 tests).
- `tests/test_sheet_views.py` (~10 tests).
- `tests/test_sheet_protection.py` (~12 tests including password hash round-trip).
- `tests/diffwriter/test_sheet_setup_*.py` (~5 tests with `WOLFXL_TEST_EPOCH=0`).
- `tests/parity/test_print_settings_parity.py` (~10 tests against openpyxl XML).
- `tests/parity/test_sheet_protection_parity.py` (~6 tests).

## 8. Out of scope (v2.1+)

- `&G` picture in header/footer body — bytes round-trip only;
  picture-add API deferred.
- `<rowBreaks>` / `<colBreaks>` (page breaks) — Pod 2 ships
  the read-side classes; full write-side feature is a v2.1
  follow-up.

## 9. References

- ECMA-376 Part 1 §18.3.1 (CT_Worksheet)
- ECMA-376 Part 1 §18.3.1.51 (CT_PageSetup)
- ECMA-376 Part 1 §18.3.1.49 (CT_PageMargins)
- ECMA-376 Part 1 §18.3.1.36 (CT_HeaderFooter)
- ECMA-376 Part 1 §18.3.1.85 (CT_SheetView)
- ECMA-376 Part 1 §18.3.1.85 (CT_SheetProtection)

## 10. Authoritative dict contract (Rust ↔ Python)

`Worksheet.to_rust_setup_dict()` returns:

```python
{
    "page_setup": {
        "orientation": str | None,
        "paper_size": int | None,
        "fit_to_width": int | None,
        "fit_to_height": int | None,
        "scale": int | None,
        "first_page_number": int | None,
        "horizontal_dpi": int | None,
        "vertical_dpi": int | None,
        "cell_comments": str | None,
        "errors": str | None,
        "use_first_page_number": bool | None,
        "use_printer_defaults": bool | None,
        "black_and_white": bool | None,
        "draft": bool | None,
    } | None,
    "page_margins": {
        "top": float, "bottom": float,
        "left": float, "right": float,
        "header": float, "footer": float,
    } | None,
    "header_footer": {
        "odd_header": {"left": str|None, "center": str|None, "right": str|None} | None,
        "odd_footer":  ...,
        "even_header": ...,
        "even_footer": ...,
        "first_header": ...,
        "first_footer": ...,
        "different_odd_even": bool,
        "different_first": bool,
        "scale_with_doc": bool,
        "align_with_margins": bool,
    } | None,
    "sheet_view": {
        "zoom_scale": int,
        "zoom_scale_normal": int,
        "view": str,                 # "normal" | "pageBreakPreview" | "pageLayout"
        "show_grid_lines": bool,
        "show_row_col_headers": bool,
        "show_outline_symbols": bool,
        "show_zeros": bool,
        "right_to_left": bool,
        "tab_selected": bool,
        "top_left_cell": str | None,
        "pane": {
            "x_split": float, "y_split": float,
            "top_left_cell": str,
            "active_pane": str,
            "state": str,
        } | None,
        "selection": [
            {"active_cell": str, "sqref": str, "pane": str | None},
            ...
        ],
    } | None,
    "sheet_protection": {
        "sheet": bool,
        "objects": bool, "scenarios": bool,
        "format_cells": bool, "format_columns": bool, "format_rows": bool,
        "insert_columns": bool, "insert_rows": bool, "insert_hyperlinks": bool,
        "delete_columns": bool, "delete_rows": bool,
        "select_locked_cells": bool,
        "sort": bool, "auto_filter": bool, "pivot_tables": bool,
        "select_unlocked_cells": bool,
        "password_hash": str | None,        # already hashed; never plaintext
    } | None,
    "print_titles": {
        "rows": str | None,                 # "1:2"
        "cols": str | None,                 # "A:B"
    } | None,
}
```

PyO3 binding `serialize_sheet_setup_dict(d) -> bytes` returns
the bytes for splice into sheet XML; `print_titles` routes
through the existing definedNames PyO3 path.

## 11. Acceptance

- All 6 user idioms from the audit script pass: `ws.page_setup`,
  `ws.page_margins`, `ws.header_footer`, `ws.print_title_rows`,
  `ws.sheet_view`, `ws.protection`.
- 70+ tests green.
- openpyxl `load_workbook` round-trips wolfxl output for each
  feature.
- LibreOffice fixture renders correctly for at least one
  fixture per feature.
- RFC-035 deep-clone preserves page setup / margins / views /
  protection on the clone.
