# Workbook API

## Constructors

### `Workbook()`

Creates a new workbook in write mode.

### `load_workbook(filename, read_only=False, data_only=False, keep_links=True, modify=False)`

Opens existing workbook.

- `modify=False`: read mode
- `modify=True`: modify mode (read + patch/save)

- `read_only=True`: uses the streaming reader surface for row iteration workflows.
- `data_only=True`: returns cached formula values where present.
- `keep_links=True`: preserves external-link parts in modify-mode saves.
- `keep_links=False`: hides external links when reading and drops external-link parts on modify-mode save.

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
