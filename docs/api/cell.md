# Cell API

## Access patterns

- `ws["A1"].value`
- `ws["A1"].font`
- `ws["A1"].fill`
- `ws["A1"].border`
- `ws["A1"].alignment`
- `ws["A1"].number_format`

## Coordinates

- `cell.coordinate`
- `cell.row`
- `cell.column`

## Value mapping

WolfXL maps common Python values to native payload types:

- `None` -> blank
- `bool` -> boolean
- `int/float` -> number
- `date/datetime` -> date/datetime
- `str` starting with `=` -> formula
- other values -> string via `str(value)`
