# RFC-059 — Public exception types + indexed-list utilities

> **Status**: Approved
> **Phase**: 5 (2.0 — Sprint Ο)
> **Depends-on**: —
> **Unblocks**: 060
> **Pod**: 1E

## 1. Goal

User code that does `try: ...; except IllegalCharacterError`
fails today because wolfxl raises generic `ValueError`. Mirror
openpyxl's exception hierarchy. Also expose the `IndexedList`
utility class and the missing cell-class re-exports.

## 2. Public API

### 2.1 Exceptions

```python
# wolfxl.utils.exceptions
class InvalidFileException(Exception):
    """Raised by load_workbook when the file is not valid xlsx/xlsb/xls."""

class IllegalCharacterError(ValueError):
    """Raised when a cell value contains characters illegal in OOXML
    (control chars 0x00-0x08, 0x0B, 0x0C, 0x0E-0x1F, 0x7F)."""

class CellCoordinatesException(ValueError):
    """Raised when an invalid coordinate string is parsed."""

class ReadOnlyWorkbookException(Exception):
    """Raised when mutating a workbook opened with read_only=True."""

class WorkbookAlreadySaved(Exception):
    """Raised when calling save() twice on a write_only workbook."""
```

`wolfxl.utils.exceptions` becomes a public module. The
existing internal raises get rewrapped:

| Existing raise | New raise |
|---|---|
| `raise ValueError("illegal character ...")` | `raise IllegalCharacterError(...)` |
| `raise ValueError("invalid coordinate ...")` | `raise CellCoordinatesException(...)` |
| `raise ValueError("file is not a valid xlsx ...")` | `raise InvalidFileException(...)` |

Re-raises preserve `__cause__` for backward compat with code
matching `except ValueError`.

### 2.2 IndexedList

```python
# wolfxl.utils.indexed_list
class IndexedList(list):
    """List with O(1) index-of lookup. Used by openpyxl for the
    shared-strings table."""
    clean: bool = True

    def __init__(self, iterable=None): ...
    def add(self, value) -> int: ...
    def index(self, value) -> int: ...
    def __contains__(self, value) -> bool: ...
```

Wolfxl's existing shared-strings table is in Rust; the Python
proxy class doesn't drive it. But `IndexedList` is part of
openpyxl's public surface and user code instantiates it
directly (e.g. for custom tables).

### 2.3 Cell class re-exports

```python
# wolfxl.cell.cell — extended
from .hyperlink import Hyperlink           # already present, re-export at this path
from ._merged import MergedCell            # new — wraps the internal proxy
from ._write_only import WriteOnlyCell     # new — already partially exists
from wolfxl.utils.exceptions import IllegalCharacterError
from wolfxl.cell.rich_text import CellRichText
from wolfxl.cell.cell import ArrayFormula, DataTableFormula  # populated by Pod 1C
```

`MergedCell` is a placeholder representing a cell that lives
inside a merged range but is not the top-left anchor cell. In
openpyxl it has `.value` always None and setting raises. Match
that contract.

`WriteOnlyCell` is a stripped-down cell used by openpyxl's
`Workbook(write_only=True)` mode. Wolfxl doesn't have a
write-only mode (the native writer handles streaming
internally), but expose `WriteOnlyCell` as a thin alias for
construction-time cell creation: `WriteOnlyCell(ws, value=42,
font=Font(bold=True))`.

## 3. Implementation

- `python/wolfxl/utils/exceptions.py` — new module.
- `python/wolfxl/utils/indexed_list.py` — new module.
- `python/wolfxl/cell/_merged.py` — new (wraps existing logic).
- `python/wolfxl/cell/_write_only.py` — new.
- Internal raises rewrapped: ~10 callsites.

## 4. Testing

- `tests/test_exceptions.py` — assert `try: ... except
  IllegalCharacterError` works on real triggers (~10 tests).
- `tests/test_indexed_list.py` — IndexedList behavior (~5
  tests).
- `tests/test_cell_classes.py` — MergedCell / WriteOnlyCell
  contracts (~6 tests).
- `tests/parity/test_exception_parity.py` — verify
  `isinstance(e, openpyxl.IllegalCharacterError)` is False
  but `type(e).__name__` matches (~5 tests; we don't subclass
  openpyxl's hierarchy, just mirror names).

## 5. References

- openpyxl 3.1.x `openpyxl.utils.exceptions` source.
- openpyxl 3.1.x `openpyxl.utils.indexed_list` source.

## 10. Dict contract

This pod is Python-only (no Rust dict shape change). No PyO3
work.

## 11. Acceptance

- `from openpyxl.utils.exceptions import IllegalCharacterError` →
  `from wolfxl.utils.exceptions import IllegalCharacterError`
  works as a drop-in import swap.
- `cell.value = "\x01"` raises `IllegalCharacterError` (not
  generic `ValueError`).
- `IndexedList(["a","b","c"]).index("b") == 1`.
- `MergedCell` `.value` is None and setter raises.
- ~26 tests green.
