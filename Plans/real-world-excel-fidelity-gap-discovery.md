# Real-World Excel Fidelity Gap Discovery

Date: 2026-05-07
Status: Active hardening plan.

## Objective

Find real-world Excel fidelity gaps that are not covered by openpyxl API parity,
then keep adding evidence until new unfound gaps become rare, explainable, and
hard to reintroduce.

This document treats "fidelity" as package meaning, not just Python API shape.
A workbook can still open while an untouched Excel dependency has been dropped,
renumbered, orphaned, or left pointing at the wrong part.

## Current repo truth

| Signal | Current state | What it proves | What it does not prove |
|---|---|---|---|
| Openpyxl parity ledger | No active tracked openpyxl-supported gaps | WolfXL covers the current openpyxl-shaped surface | Excel-only or external-tool surfaces are exhausted |
| External-oracle fixture pack | 7 pinned workbooks from Excelize, ClosedXML, NPOI, ExcelJS, Apache POI, now checked under no-op, marker-cell, style-cell, tail-row-insert, tail-column-insert, tail-row-delete, tail-column-delete, copy-remove-sheet, and marker-range-move modify-save mutations | Modify-save preserves important authored parts and still opens under safe value/style edits plus first row/column structure, row/column delete, sheet copy/remove, and range-move mutations | Broader structural edits and real Excel-authored long-tail workbooks still need coverage |
| New OOXML audit gate | `scripts/audit_ooxml_fidelity.py` now checks part loss, rel loss, dangling rels, content-type drift, feature part loss, CF dxf bounds, and deeper semantic fingerprints for charts, chart style/color parts, CF/x14 extensions, data validations, external links/cached data/formulas, pivots, slicers, and timelines | The external-oracle pack now catches broken dependency graphs and feature-meaning drift across the named P0 surfaces when those parts are present in the fixture | It is not yet a full Excel-rendered semantic validator or real-Excel corpus proof |

## Risk matrix

| Priority | Surface | Failure modes to mine | Required oracle |
|---|---|---|---|
| P0 | Pivot/slicer preservation across modify saves | Shared pivot caches, slicer cache `extLst`, table slicers, timelines, pivot chart filter buttons, multi-pivot same cache, refresh flags | Package graph diff, Excel/LibreOffice open-save, pivot/slicer metadata probe |
| P0 | Chart/style/color dependency preservation | `chartStyle`/`chartColorStyle` parts, theme/indexed colors, series dPt overrides, combo chart shadow parts, chart sheets, drawing rels | Package graph diff, chart XML semantic probe, rendered PDF/image compare |
| P0 | Conditional formatting extension preservation | x14 data bars/icon sets, pivot-scoped CF, `sqref` expansion after structural edits, formula translation, `priority` and `dxfId` stability | CF semantic probe plus `dxfId`/`dxfs` integrity check |
| P0 | External-link and workbook relationship edge cases | External link cached data, `keep_links`, externalLinkPath rels, workbook rel order, formulas referencing closed workbooks | Relationship graph diff, formula/external-link readback, Excel repair-dialog check |
| P1 | Tables / structured refs / validations | Totals formulas, filter state, table styles, structured-reference formulas, validation formulas moved with rows/cols | Table metadata probe and structural mutation sweep |
| P1 | Drawings / comments / embedded objects | VML comment drawings, threaded comment people, image anchors, hyperlink rels, embedded packages | Drawing rel graph probe and render/open-save checks |
| P1 | Workbook global state | Defined names, calc chain, workbook protection, VBA, custom XML, printer/page setup | Package graph diff plus targeted XML probes |

## Discovery loop

```text
Real Excel / external-tool source workbook
        |
        v
Mutation suite:
  no-op save, one-cell edit, style edit, row/col insert/delete,
  move range, sheet rename/copy/remove, feature add/remove
        |
        v
OOXML package audit:
  parts, rels, targets, content types, feature hotspots
        |
        v
Feature semantic probes:
  pivots, slicers, charts, CF, links, tables, drawings
        |
        v
Excel / LibreOffice / external-reader validation
        |
        v
Gap ledger:
  fixture, mutation, broken invariant, fix, regression test
```

## Concrete next gates

1. Extend `scripts/audit_ooxml_fidelity.py` from first-pass semantic fingerprints
   to deeper feature-level summaries:
   - Done: pivot table rel targets, calculated fields/items, formats, conditional
     formats, pivot-cache rel targets, and field groups.
   - Done: slicer workbook/sheet `extLst` anchors, slicer/slicer-cache rels,
     slicer cache data, and timeline workbook/sheet anchors plus timeline parts.
   - Done: chart rel targets and chart style/color XML part fingerprints.
   - Done: conditional-formatting `sqref`/rule fingerprints, x14 extension
     subtrees, data-validation ranges/formulas, and CF `dxfId` bounds.
   - Done: external-link targets, cached sheet data, sheet names, defined names,
     and worksheet formulas that reference linked workbooks.
   - Done: chart axis IDs, axis metadata, manual layout, series summaries,
     and chart-sheet fingerprints.
   - Still needed: rendered output comparison and formula translation
     semantics after structural edits.
2. Extend the mutation runner beyond safe edits:
   - Current command:
     `uv run --no-sync python scripts/run_ooxml_fidelity_mutations.py tests/fixtures/external_oracle --output-dir /tmp/wolfxl-ooxml-fidelity-sweep`
   - Current safe/default mutations: no-op modify-save, marker-cell modify-save,
     style-cell modify-save, tail-row insert, tail-column insert, tail-marker
     row/column delete, copy-remove-sheet, and marker-range move.
   - Latest pinned-pack sweep: 63 results, 0 failures across 7 fixtures and 9
     default mutations.
   - Latest deeper-fingerprint sweep: 63 results, 0 failures across the same
     pinned pack and default mutation set.
   - Latest feature-add opt-in sweep: add-data-validation and
     add-conditional-formatting pass 14 results across 7 fixtures with only
     declared additive semantic drift.
   - Latest opt-in semantic sweeps: sheet rename passes 7 results with no
     drift; first-row/first-column delete passes 14 results with expected
     conditional-formatting range drift and no unexpected package-fidelity
     failure; first-sheet copy initially exposed a prefixed workbook-root
     splice bug and now passes 7 results as an opt-in expected-drift gate.
   - Latest feature-add bugs found: add-conditional-formatting initially hid a
     ClosedXML x14 CF corruption under broad expected-drift handling; the
     runner now requires the added range marker for feature-add expected drift,
     and CF extraction snaps prefixed raw slices to real tag boundaries.
   - Latest sheet-remove bug found: copy-then-remove originally deleted shared
     image media and left cloned sheet parts behind; the delete cleanup now
     honors in-progress relationship graphs and skips parts still referenced by
     kept workbook parts.
   - Latest bugs found: row insertion exposed a prefixed-XML end-tag corruption
     path in structural rewrites; range move exposed a prefixed `sheetData`
     discovery/re-emission gap. Both are now covered by regression tests.
   - Next mutations: feature remove and richer chart/pivot/slicer structural
     edits where the expected semantic drift can be declared.
3. Expand fixture sources:
   - real Excel-authored workbooks with slicers, timelines, pivot charts, chart
     style/color parts, and external links;
   - generated external-oracle fixtures from Excelize, ClosedXML, Apache POI,
     NPOI, ExcelJS, LibreOffice, and targeted low-level writers.
4. Promote every discovered failure into a minimal fixture and a regression test.
5. Do not call the surface "clear" until each P0 row has at least one real
   Excel-authored fixture, one external-tool fixture, and one structural
   mutation fixture passing the package audit plus feature probe.

## Confidence standard

"No known real-world fidelity gaps" requires all of these:

- no open P0/P1 failures in the gap ledger;
- clean external-oracle preservation with the OOXML audit enabled;
- clean mutation sweep over the real Excel corpus;
- clean Excel and LibreOffice open/save smoke for business-critical fixture
  classes;
- each explicit gap class above has direct evidence, not only a proxy test.

Anything less should be described as "no known gap in the currently covered
surface," not as exhaustive Excel fidelity.
