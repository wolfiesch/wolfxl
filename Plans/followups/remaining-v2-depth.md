# Remaining v2 Depth Follow-Ups

Status: shipped in the follow-up sprint; cache-record regeneration remains a
separate future optimization.

## Pivot Field / Filter / Aggregation Mutation

Shipped surface:

- Construct pivot caches, pivot tables, pivot charts, and pivot-backed slicers.
- Copy pivot-bearing worksheets with deep-cloned pivot table parts.
- Field placement: mutate existing `<rowFields>`, `<colFields>`, `<pageFields>`, and `<dataFields>` blocks.
- Filters: edit page-field and item visibility filters on an existing pivot table.
- Aggregation: change data-field functions such as `sum`, `count`, and `average`.
- Cache regeneration: deferred; layout edits mark refresh-on-open explicitly.

Verification:

- `tests/test_pivot_modify_existing.py` covers field placement, page-filter selection, aggregation changes, and source+layout composition.
- `tests/test_openpyxl_compat_oracle.py` includes `pivots_field_mutation`, `pivots_filter_mutation`, and `pivots_aggregation_mutation`.

## External Workbook Link Authoring

Shipped surface:

- Inspect `_external_links` for target path, sheet names, and cached data.
- Append a new external link from Python.
- Remove an existing external link and prune its rel/content-type wiring.
- Edit an existing external link target.
- Keep cached data explicit: preserve provided payload; do not dereference linked workbooks.

Verification:

- `tests/test_external_links.py` covers append, modify-mode append, remove, and `update_target`.
- `tests/test_openpyxl_compat_oracle.py` includes `external_links_authoring`.
- Tests verify `[Content_Types].xml`, `xl/_rels/workbook.xml.rels`, and each `xl/externalLinks/_rels/*.rels` path after mutation.
