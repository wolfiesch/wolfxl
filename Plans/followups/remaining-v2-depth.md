# Remaining v2 Depth Follow-Ups

Status: follow-up staging after the remaining-limitations closure slice.

## Pivot Field / Filter / Aggregation Mutation

Current v1 surface:

- Construct pivot caches, pivot tables, pivot charts, and pivot-backed slicers.
- Copy pivot-bearing worksheets with deep-cloned pivot table parts.
- Modify an existing pivot table's source range via `ws.pivot_tables[i].source = ...`.

v2 acceptance targets:

- Field placement: mutate existing `<rowFields>`, `<colFields>`, `<pageFields>`, and `<dataFields>` blocks.
- Filters: edit page-field and item visibility filters on an existing pivot table.
- Aggregation: change data-field functions such as `sum`, `count`, and `average`.
- Cache regeneration: optionally refresh cache records when WolfXL can compute the aggregate safely; otherwise mark refresh-on-open explicitly.

Suggested verification:

- Expand `tests/test_pivot_modify_existing.py` with focused fixtures for each mutation family.
- Add oracle probes for `pivots_field_mutation`, `pivots_filter_mutation`, and `pivots_aggregation_mutation`.
- Keep the current `pivots.in_place_edit` row honest by documenting source-range-only support until these probes pass.

## External Workbook Link Authoring

Current v1 surface:

- Inspect `_external_links` for target path, sheet names, and cached data.
- Preserve existing `xl/externalLinks/*` parts and their relationships byte-for-byte in modify mode.

v2 acceptance targets:

- Append a new external link from Python.
- Remove an existing external link and prune its rel/content-type wiring.
- Edit an existing external link target.
- Keep cached data explicit: either update it from a provided payload or preserve/clear it predictably.

Suggested verification:

- Extend `tests/test_external_links.py` from inspection/preservation into append/remove/edit round trips.
- Add authoring coverage to the `external_links_collection` oracle only after public API names are settled.
- Verify `[Content_Types].xml`, `xl/_rels/workbook.xml.rels`, and each `xl/externalLinks/_rels/*.rels` path after every mutation.
