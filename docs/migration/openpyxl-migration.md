# Migrate from openpyxl

## Minimal Import Change

```python
# before
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border

# after
from wolfxl import load_workbook, Workbook, Font, PatternFill, Alignment, Border
```

## What usually stays the same

- `wb = load_workbook(path)`
- `wb = Workbook()`
- `ws = wb["Sheet"]`
- `ws["A1"].value = ...`
- `ws["A1"].font = Font(...)`
- `wb.save(path)`

## What to validate during migration

1. Style fidelity in your critical sheets
2. Formula behavior in your downstream consumers
3. Any rarely used openpyxl-only APIs

## Migration Playbook

1. Swap imports in one workflow
2. Run your existing tests
3. Compare output workbook in Excel
4. Measure runtime/memory
5. Roll out gradually to other pipelines

For detailed support coverage, see [Compatibility Matrix](compatibility-matrix.md).
