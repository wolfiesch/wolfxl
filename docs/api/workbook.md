# Workbook API

## Constructors

### `Workbook()`

Creates a new workbook in write mode.

### `load_workbook(filename, read_only=False, data_only=False, keep_links=True, modify=False)`

Opens existing workbook.

- `modify=False`: read mode
- `modify=True`: modify mode (read + patch/save)

Compatibility kwargs (`read_only`, `data_only`, `keep_links`) are currently accepted but ignored.

## Properties

- `sheetnames -> list[str]`
- `active -> Worksheet | None`

## Methods

- `create_sheet(title: str) -> Worksheet` (write mode)
- `save(filename: str) -> None`
- `close() -> None`
- `__getitem__(name: str) -> Worksheet`

## Context manager

`Workbook` supports `with` statements.
