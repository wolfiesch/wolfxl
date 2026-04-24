# Worksheet API

## Access

- `wb["Sheet1"]`
- `ws["A1"]`
- `ws.cell(row=1, column=1, value=...)`

## Core methods

- `iter_rows(min_row=None, max_row=None, min_col=None, max_col=None, values_only=False)`
- `iter_cell_records(min_row=None, max_row=None, min_col=None, max_col=None, data_only=None, include_format=True, include_empty=False, include_formula_blanks=True, include_coordinate=True)`
- `cell_records(...)`
- `calculate_dimension()` — returns the actual used range, including offset ranges like `C4:C4`
- `merge_cells(range_string)` (write mode)

## Examples

```python
ws["A1"] = "Hello"
cell = ws.cell(2, 1, "World")

for row in ws.iter_rows(min_row=1, max_row=2, min_col=1, max_col=1, values_only=True):
    print(row)
```

## Bulk Cell Records

Use `cell_records()` when you need values plus compact formatting metadata
without paying for one Python property access per cell.

```python
records = ws.cell_records(include_format=True)

for record in records:
    print(
        record["coordinate"],
        record["value"],
        record.get("number_format"),
        record.get("bold", False),
    )
```

Each record uses openpyxl-style 1-based coordinates:

- `row`, `column`, `coordinate`
- `value`, `data_type`
- `formula` when the cell contains a formula
- `style_id`, `number_format`, `bold`, `italic`, `font_size`, `h_align`, `indent`
- `bottom_border_style`, `has_bottom_border`, `is_double_underline`

Empty cells are skipped by default. Pass `include_empty=True` to emit a dense
rectangular record stream for the requested range.

Formula cells return formula text by default. Pass `data_only=True` to return
cached formula values instead; uncached formula blanks are skipped unless
`include_empty=True`. For ingestion/dataframe workloads that do not want sparse
template formulas to inflate the record stream, pass
`include_formula_blanks=False`. Pass `include_coordinate=False` when 1-based
`row` / `column` integers are enough and you want to skip A1 coordinate string
allocation. Pass `include_style_id=False` when you need semantic format fields
but not workbook-internal style ids. Pass `include_extended_format=False` to
keep raw font flags and number formats while skipping style-grid fields such as
fill, alignment, and border cues.

`calculate_dimension()` follows openpyxl's used-range shape. A blank sheet
returns `A1:A1`; a sheet with only `C4` populated returns `C4:C4`; a sheet with
values from `A1` through `C7` returns `A1:C7`. `max_row` and `max_column`
continue to expose the bottom/right edge of that range.
