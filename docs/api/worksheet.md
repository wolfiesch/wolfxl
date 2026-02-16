# Worksheet API

## Access

- `wb["Sheet1"]`
- `ws["A1"]`
- `ws.cell(row=1, column=1, value=...)`

## Core methods

- `iter_rows(min_row=None, max_row=None, min_col=None, max_col=None, values_only=False)`
- `merge_cells(range_string)` (write mode)

## Examples

```python
ws["A1"] = "Hello"
cell = ws.cell(2, 1, "World")

for row in ws.iter_rows(min_row=1, max_row=2, min_col=1, max_col=1, values_only=True):
    print(row)
```
