# RFC-057 — Dynamic-array formulas (`ArrayFormula`, `DataTableFormula`)

> **Status**: Approved
> **Phase**: 5 (2.0 — Sprint Ο)
> **Depends-on**: 012 (formula reference translator), 013 (patcher)
> **Unblocks**: 060
> **Pod**: 1C

## 1. Goal

Excel 365 spilled-range formulas (`=SEQUENCE(10)` spills to
A1:A10) and pre-365 array formulas (`{=SUM(A1:A10*B1:B10)}`)
have no construction path in wolfxl today. openpyxl exposes
`ArrayFormula(ref, text)` and `DataTableFormula(ref, ca, dt2D,
dtr, r1, r2)`.

## 2. Public API

```python
from wolfxl.cell.cell import ArrayFormula, DataTableFormula

# Pre-365 array formula (CSE form)
cell = ws["A1"]
cell.value = ArrayFormula(ref="A1:A10", text="B1:B10*2")
# Resulting XML: <c r="A1"><f t="array" ref="A1:A10">B1:B10*2</f></c>

# Excel 365 dynamic array (single-cell formula spilling)
cell = ws["A1"]
cell.value = "=SEQUENCE(10)"
ws.formula_attributes["A1"] = {"t": "array", "ref": "A1:A10"}

# 2D data table
cell = ws["B2"]
cell.value = DataTableFormula(
    ref="B2:F11",
    ca=False,             # column-anchored
    dt2D=True,             # two-variable data table
    dtr=False,             # row-anchored
    r1="A1",               # input cell 1
    r2="A2",               # input cell 2
)
```

## 3. Class definitions

```python
class ArrayFormula:
    ref: str          # spill range, e.g. "A1:A10"
    text: str         # formula body without leading "=" or surrounding {}

    def __init__(self, ref: str, text: str): ...
    def __repr__(self) -> str: ...

class DataTableFormula:
    ref: str          # range that the data table fills
    ca: bool = False                  # always-calculate
    dt2D: bool = False                # 2D data table flag
    dtr: bool = False                 # row-input flag
    r1: str | None = None             # first input cell
    r2: str | None = None             # second input cell (only for 2D)
```

## 4. OOXML output

```xml
<!-- ArrayFormula -->
<c r="A1">
  <f t="array" ref="A1:A10">B1:B10*2</f>
</c>

<!-- DataTableFormula -->
<c r="B2">
  <f t="dataTable" ref="B2:F11" dt2D="1" r1="A1" r2="A2"/>
</c>
```

For the spill range, child cells in the array carry just
`<c r="A2"/>` etc (the master cell at A1 holds the formula).

## 5. Cell.value setter contract

`cell.value = ArrayFormula(...)`:
- Sets the cell's formula type to `"array"`.
- Stamps the `ref` attribute on the master cell.
- Cells inside the spill range gain placeholders if not
  already populated. Wolfxl does not pre-evaluate the
  formula; Excel computes spill on open.

`cell.value = DataTableFormula(...)`:
- Sets the formula type to `"dataTable"`.
- Stamps `ref`, `dt2D`, `r1`, `r2` attributes.

`cell.value` getter returns the `ArrayFormula` /
`DataTableFormula` instance for cells whose `<f t="array">` or
`<f t="dataTable">` parsed back. For cells in the spill range
that aren't the master, `cell.value` returns `None` (matches
openpyxl).

## 6. RFC-035 / modify mode

Modify-mode patcher already routes formula text through Phase
3 cell patches. Pod 1C extends the cell-patch shape to carry
the new `formula_type`, `array_ref`, `data_table_*` fields.

RFC-035 deep-clone: array-formula refs are sheet-scoped and
need re-pointing if the formula references cells inside the
copied sheet (RFC-012 translator already handles this for
plain formulas; extend for the `ref` attribute).

## 7. Out of scope (v2.1+)

- **Spill evaluation** — Excel computes the spill output on
  open. wolfxl does not evaluate `=SEQUENCE(10)` etc.
  Documented divergence; users who need pre-evaluated values
  must run Excel/LibreOffice once.
- **Implicit `_xlfn.` prefixing** for newer Excel functions.
  Wolfxl already does this in RFC-012; verify for spill
  functions specifically.

## 8. Testing

- `tests/test_array_formula.py` (~15 tests).
- `tests/test_data_table_formula.py` (~8 tests).
- `tests/diffwriter/test_array_formula_*.py` (~3 byte-stable tests).
- `tests/parity/test_array_formula_parity.py` (~5 openpyxl interop tests).
- LibreOffice fixture render (1 manual test).

## 9. References

- ECMA-376 Part 1 §18.17 (formula-related types)
- Excel 365 spill behavior — Microsoft Learn
- openpyxl 3.1.x `openpyxl.worksheet.formula` source.

## 10. Dict contract

Cell-level patch carries:

```python
{
    "kind": "array" | "data_table" | "normal",
    "formula": {
        "text": str,                # formula body
        "ref": str | None,          # for array/data_table
        "ca": bool | None,
        "dt2D": bool | None,
        "dtr": bool | None,
        "r1": str | None,
        "r2": str | None,
    } | None,
}
```

## 11. Acceptance

- `cell.value = ArrayFormula("A1:A10", "B1:B10*2")` round-trips.
- `cell.value = DataTableFormula(...)` round-trips.
- openpyxl `load_workbook` sees the same `f.t == "array"` /
  `"dataTable"`.
- LibreOffice renders the spill correctly (manual verification).
- ~30 tests green.
