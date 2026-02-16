# Styles API

WolfXL exposes openpyxl-style dataclasses:

- `Font`
- `PatternFill`
- `Alignment`
- `Border`
- `Side`
- `Color`

## Example

```python
from wolfxl import Font, PatternFill, Alignment, Border, Side

ws["A1"].font = Font(bold=True, name="Calibri", size=11)
ws["A1"].fill = PatternFill(patternType="solid", fgColor="FFF2CC")
ws["A1"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
ws["A1"].border = Border(left=Side(style="thin", color="000000"))
ws["A1"].number_format = "#,##0.00"
```

## Notes

- Style writes are flushed on `wb.save()`.
- In modify mode, style changes are queued and applied during patch save.
