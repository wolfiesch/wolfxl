# Quickstart

## 1) Install

```bash
pip install wolfxl
```

## 2) Write a Workbook

```python
from wolfxl import Workbook, Font

wb = Workbook()
ws = wb.active
ws["A1"] = "Hello WolfXL"
ws["A1"].font = Font(bold=True)
wb.save("quickstart.xlsx")
wb.close()
```

## 3) Read It Back

```python
from wolfxl import load_workbook

wb = load_workbook("quickstart.xlsx")
ws = wb[wb.sheetnames[0]]
print(ws["A1"].value)
wb.close()
```

## 4) Modify Existing File

```python
from wolfxl import load_workbook, PatternFill

wb = load_workbook("quickstart.xlsx", modify=True)
ws = wb["Sheet"]
ws["B1"] = "Edited"
ws["B1"].fill = PatternFill(patternType="solid", fgColor="FFD966")
wb.save("quickstart_modified.xlsx")
wb.close()
```

Next: [First Workbook Walkthrough](first-workbook.md)
