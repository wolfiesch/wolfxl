# RFC-063 — Merge + tables + worksheet copier (Sprint Π Pod Π-β)

> Closes 7 stubs: `MergeCell`, `MergeCells`, `WorksheetCopy`,
> `TableList`, `TablePartList`, `Related`, `XMLColumnProps`.

## 1. Goal

Wrap existing wolfxl primitives with openpyxl-shaped value types
that preserve attribute access + iteration semantics. NO new
behavior — these are all type-only proxies over state that already
flows to disk.

## 2. Python class layer

### 2.1 `python/wolfxl/worksheet/merge.py` (replaces stubs)

```python
from dataclasses import dataclass
from wolfxl.worksheet.cell_range import CellRange


@dataclass
class MergeCell:
    """Single merged region (CT_MergeCell §18.3.1.55).

    Wolfxl stores merged ranges as plain A1 strings; this class
    provides the openpyxl-shaped type-wrapper so user code
    constructing a MergeCell continues to work."""
    ref: str

    @property
    def coord(self) -> str:
        return self.ref


class MergeCells:
    """Container for MergeCell entries (CT_MergeCells §18.3.1.56).

    Backed by ``ws.merged_cells`` plain set; this wrapper exposes
    iteration, count, and append for openpyxl source compatibility."""
    worksheet: Any
    mergeCell: list[MergeCell]  # noqa: N815

    def __init__(self, worksheet=None, mergeCell=None):
        self.worksheet = worksheet
        self.mergeCell = list(mergeCell) if mergeCell else []

    @property
    def count(self) -> int:
        return len(self.mergeCell)

    def append(self, mc: MergeCell) -> None: ...
    def __iter__(self): return iter(self.mergeCell)
    def __len__(self): return self.count
```

### 2.2 `python/wolfxl/worksheet/copier.py` (replaces stub)

```python
class WorksheetCopy:
    """Thin wrapper over ``Workbook.copy_worksheet`` (RFC-035).

    openpyxl exposes this as a class with `.copy_worksheet()` method;
    wolfxl users typically call ``wb.copy_worksheet(src)`` directly,
    but the class is provided for source compatibility."""

    def __init__(self, source: Any, target: Any) -> None:
        self.source = source
        self.target = target

    def copy_worksheet(self) -> Any:
        # Delegates to the existing RFC-035 deep-clone path.
        wb = self.source._workbook
        return wb.copy_worksheet(self.source, target_title=self.target.title)
```

### 2.3 `python/wolfxl/worksheet/table.py` (replaces stubs)

```python
@dataclass
class Related:
    """rels-pointer dataclass (`r:id="rId1"`)."""
    id: str = ""


@dataclass
class XMLColumnProps:
    """XML-column metadata for table-bound columns."""
    mapId: int                   # noqa: N815
    xpath: str
    denormalized: bool = False
    xmlDataType: str = "string"  # noqa: N815


class TableList:
    """Container over ``ws.tables`` (an existing dict).

    Provides the openpyxl-shape ``__iter__`` / ``__len__`` /
    ``add()`` / ``remove()`` API."""
    worksheet: Any
    def __init__(self, worksheet=None): ...
    def add(self, table): ...
    def remove(self, table_name: str): ...
    def __iter__(self): ...
    def __len__(self): ...
    def __contains__(self, name: str): ...


@dataclass
class TablePartList:
    """`<tableParts>` serialization helper."""
    count: int = 0
    tablePart: list[Related] = None  # noqa: N815

    def __post_init__(self): ...
```

## 3. Integration

- `Worksheet.merge_cells_obj` lazy property returns a `MergeCells`
  instance synced to the underlying set.
- `Worksheet.tables_list` lazy property returns a `TableList` view
  over the existing `ws.tables` dict.
- No save() pipeline changes — these are type proxies only.

## 4. §10 contracts

None new — these wrap existing dict-shape data already plumbed
through patcher / native writer.

## 5. Tests

- `tests/test_merge_value_types.py` (~15)
- `tests/test_worksheet_copier.py` (~10)
- `tests/test_table_value_types.py` (~20)
- `tests/parity/test_merge_iteration_parity.py` — verify iteration order matches openpyxl

## 6. Tolerable pre-existing failures

None.
