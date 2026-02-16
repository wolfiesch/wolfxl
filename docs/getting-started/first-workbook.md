# First Workbook Walkthrough

This walkthrough demonstrates the three WolfXL modes.

## Write Mode

Use `Workbook()` to create a new file.

```python
from wolfxl import Workbook

wb = Workbook()
ws = wb.active
ws["A1"] = "Region"
ws["B1"] = "Revenue"
ws["A2"] = "NA"
ws["B2"] = 125000
wb.save("report.xlsx")
wb.close()
```

## Read Mode

Use `load_workbook(path)` for fast reads.

```python
from wolfxl import load_workbook

wb = load_workbook("report.xlsx")
ws = wb["Sheet"]
for row in ws.iter_rows(min_row=1, max_row=2, min_col=1, max_col=2, values_only=True):
    print(row)
wb.close()
```

## Modify Mode

Use `load_workbook(path, modify=True)` for read-modify-write workflows.

```python
from wolfxl import load_workbook

wb = load_workbook("report.xlsx", modify=True)
ws = wb["Sheet"]
ws["B2"] = 130000
wb.save("report_updated.xlsx")
wb.close()
```

## Notes

- `read_only`, `data_only`, and `keep_links` kwargs are accepted for compatibility.
- Save operations flush dirty cells in batch.
