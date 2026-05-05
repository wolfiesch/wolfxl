# Standalone Slicer Openpyxl Evidence

Status: reclassified as `out_of_scope` for openpyxl parity.

Openpyxl reference: 3.1.5.

Evidence checked locally:

- `pkgutil.walk_packages(openpyxl.__path__)` returns no module path containing `slicer`.
- `openpyxl.worksheet.table`, `openpyxl.worksheet.worksheet`, `openpyxl.pivot.table`, and `openpyxl.pivot.cache` expose no public names containing `Slicer` or `slicer`.
- A source search under the installed `openpyxl` package finds no `Slicer`, `slicer`, `slicerCache`, or `tableSlicer` symbols.

Decision:

- Pivot-backed slicers remain a supported WolfXL API.
- Standalone table-driven slicer authoring is not an openpyxl compatibility gap because openpyxl has no public authoring surface to match.
- If WolfXL adds table-driven slicer authoring later, track it as a WolfXL-extra with its own OOXML acceptance tests rather than as a parity closure row.
