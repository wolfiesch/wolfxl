# RFC-065 — Workbook properties + internals (Sprint Π Pod Π-δ)

> Closes 5 stubs: `_WorkbookChild`, `CalcProperties`,
> `WorkbookProperties`, `CommentSheet`, `SpreadsheetDrawing`.

## 1. Goal

Land Python construction + state + (where applicable) save() integration
for workbook-level property dataclasses and the two opaque drawing/comment
serializer wrappers.

## 2. Python class layer

### 2.1 `python/wolfxl/workbook/properties.py` (replaces stubs)

```python
from dataclasses import dataclass


@dataclass
class CalcProperties:
    """`<calcPr>` element (CT_CalcPr §18.2.2).

    Backs ``wb.calc_properties``. Stored on Workbook; emitted into
    xl/workbook.xml via patcher Phase 2.5q (workbook-level)."""
    calcId: int = 124519             # noqa: N815
    calcMode: str = "auto"            # noqa: N815  # auto | autoNoTable | manual
    fullCalcOnLoad: bool = False     # noqa: N815
    refMode: str = "A1"              # noqa: N815  # A1 | R1C1
    iterate: bool = False
    iterateCount: int = 100          # noqa: N815
    iterateDelta: float = 0.001      # noqa: N815
    fullPrecision: bool = True       # noqa: N815
    calcCompleted: bool = True       # noqa: N815
    calcOnSave: bool = True          # noqa: N815
    concurrentCalc: bool = True      # noqa: N815
    concurrentManualCount: int | None = None  # noqa: N815
    forceFullCalc: bool = False      # noqa: N815

    def to_rust_dict(self) -> dict[str, Any]: ...


@dataclass
class WorkbookProperties:
    """`<workbookPr>` element (CT_WorkbookPr §18.2.28).

    Backs ``wb.workbook_properties``. Carries date1904, codeName, etc."""
    date1904: bool = False
    dateCompatibility: bool = True       # noqa: N815
    showObjects: str = "all"             # noqa: N815  # all | placeholders | none
    showBorderUnselectedTables: bool = True  # noqa: N815
    filterPrivacy: bool = False          # noqa: N815
    promptedSolutions: bool = False      # noqa: N815
    showInkAnnotation: bool = True       # noqa: N815
    backupFile: bool = False             # noqa: N815
    saveExternalLinkValues: bool = True  # noqa: N815
    updateLinks: str = "userSet"         # noqa: N815
    codeName: str | None = None          # noqa: N815
    hidePivotFieldList: bool = False     # noqa: N815
    showPivotChartFilter: bool = False   # noqa: N815
    allowRefreshQuery: bool = False      # noqa: N815
    publishItems: bool = False           # noqa: N815
    checkCompatibility: bool = False     # noqa: N815
    autoCompressPictures: bool = True    # noqa: N815
    refreshAllConnections: bool = False  # noqa: N815
    defaultThemeVersion: int = 124226    # noqa: N815

    def to_rust_dict(self) -> dict[str, Any]: ...
```

### 2.2 `python/wolfxl/workbook/child.py` (replaces stub)

```python
class _WorkbookChild:
    """openpyxl mixin tracking parent-workbook reference + sheet
    title constraints (no `:`, max 31 chars).

    Wolfxl's Worksheet / ChartSheet already carry the parent-ref via
    ``_workbook`` slot; this mixin is exposed for ``isinstance`` parity
    only — NOT injected into the MRO (would break Worksheet's __slots__).
    """
    parent: Any = None
    title: str = ""

    @classmethod
    def __subclasshook__(cls, C):
        # Worksheet + ChartSheet pass an ``isinstance(ws, _WorkbookChild)`` check
        if hasattr(C, "_workbook") and hasattr(C, "title"):
            return True
        return NotImplemented
```

### 2.3 `python/wolfxl/comments/comments.py` (replaces stub)

```python
@dataclass
class CommentSheet:
    """Internal serializer wrapper over ``ws.comments`` (existing).

    openpyxl users rarely instantiate this directly — they call
    ``ws['A1'].comment = Comment(...)``. Provided for source compat
    with code that walks ``ws._comments`` or constructs a CommentSheet
    explicitly for tests."""
    worksheet: Any = None
    comments: list[Any] = None
    authors: list[str] = None

    def __post_init__(self): ...
    def to_rust_dict(self) -> dict[str, Any]: ...
```

### 2.4 `python/wolfxl/drawing/spreadsheet_drawing.py` (replaces stub)

```python
@dataclass
class SpreadsheetDrawing:
    """Internal wrapper over a sheet's drawings collection.

    Backs ``ws._drawing`` for openpyxl source compatibility; users
    typically interact via ``ws.add_image(...)`` / ``ws.add_chart(...)``."""
    worksheet: Any = None
    images: list[Any] = None
    charts: list[Any] = None

    def __post_init__(self): ...
```

## 3. Workbook integration

```python
class Workbook:
    @property
    def calc_properties(self) -> CalcProperties:
        if self._calc_properties is None:
            self._calc_properties = CalcProperties()
        return self._calc_properties

    @property
    def workbook_properties(self) -> WorkbookProperties:
        if self._workbook_properties is None:
            self._workbook_properties = WorkbookProperties()
        return self._workbook_properties

    def _flush_pending_workbook_props_to_patcher(self) -> None:
        """Drain calc_properties + workbook_properties into xl/workbook.xml.
        Sequenced WITH workbook security in Phase 2.5q."""
```

New `__slots__`: `_calc_properties`, `_workbook_properties`.

## 4. Rust emit + parser

`crates/wolfxl-writer/src/emit/workbook_props.rs` (NEW):
- `emit_calc_pr(spec) -> bytes`
- `emit_workbook_pr(spec) -> bytes`

`crates/wolfxl-writer/src/parse/workbook_props.rs` (NEW): consumes §10
dicts.

## 5. Patcher Phase 2.5q (extend, don't add new phase)

Pod-1D's Phase 2.5q already splices `<workbookProtection>` and
`<fileSharing>` into `xl/workbook.xml`. Pod-Π-δ extends the same drain
to also splice `<workbookPr>` and `<calcPr>` from the new queues:

- `XlsxPatcher.queue_workbook_props_update(calc_dict, workbook_dict)` PyO3 method
- New `queued_workbook_props: Option<WorkbookProps>` field
- Add to do_save no-op guard

## 6. §10 contracts

```python
calc_properties.to_rust_dict() = {
    "calc_id": 124519,
    "calc_mode": "auto",
    "full_calc_on_load": False,
    "ref_mode": "A1",
    "iterate": False,
    "iterate_count": 100,
    "iterate_delta": 0.001,
    "full_precision": True,
    "calc_completed": True,
    "calc_on_save": True,
    "concurrent_calc": True,
    "concurrent_manual_count": null,
    "force_full_calc": False,
}

workbook_properties.to_rust_dict() = {
    "date1904": False,
    "code_name": null,
    "default_theme_version": 124226,
    ...  # 19 fields total
}
```

## 7. Tests

- `tests/test_calc_properties.py` (~15) — construction, attr access, defaults
- `tests/test_workbook_properties.py` (~20)
- `tests/test_workbook_child_isinstance.py` (~5)
- `tests/test_comment_sheet_wrapper.py` (~10)
- `tests/test_spreadsheet_drawing_wrapper.py` (~10)
- `tests/diffwriter/test_workbook_props_bytes.py` (`WOLFXL_TEST_EPOCH=0`)
- `tests/parity/test_workbook_props_parity.py` — round-trip via openpyxl

## 8. Tolerable pre-existing failures

None.
