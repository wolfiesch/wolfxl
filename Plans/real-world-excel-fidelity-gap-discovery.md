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
| External-oracle fixture pack | 14 active pinned workbooks from ClosedXML, Excelize, NPOI, openpyxl, one synthetic external-link OOXML oracle, one Excel-authored chart/CF workbook, one Excel-authored external-link workbook, one Excel-normalized pivot/CF workbook, one Excelize 2.10 pivot-slicer workbook, two public Excel-authored MyExcelOnline slicer workbooks, and one public Excel-authored MyExcelOnline timeline workbook. The Apache POI and ExcelJS image/comment/table sources are preserved under `tests/fixtures/external_oracle/rejected/` because Excel rejects them before WolfXL mutation | Modify-save preserves important authored parts and still opens under safe value/style edits, structural edits, sheet copy/rename/remove, range moves, and external-link relationship mutations. Active fixtures now pass Microsoft Excel open/save smoke, including external-tool slicer, Excel-authored slicer, and Excel-authored timeline fixtures | This is a strong current evidence gate, not exhaustive proof. Rendered comparisons, richer structural edits, and broader real-file corpora still need coverage |
| New OOXML audit gate | `scripts/audit_ooxml_fidelity.py` now checks part loss, rel loss, dangling rels, content-type drift, feature part loss, CF dxf bounds, and deeper semantic fingerprints for charts, chart style/color parts, CF/x14 extensions, data validations, worksheet formulas, external links/cached data/formulas, pivots, slicers, and timelines | The external-oracle pack now catches broken dependency graphs and feature-meaning drift across the named P0 surfaces when those parts are present in the fixture | It is not yet a full Excel-rendered semantic validator or real-Excel corpus proof |
| Coverage evidence audit | `scripts/audit_ooxml_fidelity_coverage.py` maps fixtures plus mutation reports to the P0 evidence standard: external-tool fixture, real Excel fixture, and structural mutation pass. It records concrete feature keys and requires slicer/timeline evidence separately from pivot-table evidence | The current strict P0 evidence gate is green: pivot/slicer, chart/style/color, conditional-formatting, and external-link rows all have external-tool evidence, real Excel evidence, and structural mutation passes. The slicer/timeline group now includes direct timeline evidence from `real-excel-timeline-slicer.xlsx` | It only audits evidence presence; it does not prove rendered visual fidelity or broader corpus coverage |
| App open/save smoke | `scripts/run_ooxml_app_smoke.py` opens and re-saves fixture packs through LibreOffice headless or Microsoft Excel, then validates the saved file as an OOXML ZIP | Microsoft Excel and LibreOffice now open/re-save all 14 active external-oracle fixtures cleanly | The active app-smoke pack is not a rendered comparison and does not prove every real-world workbook class |

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
   - Done: worksheet formula cell-coordinate and formula-text fingerprints,
     plus an opt-in formula move translation oracle.
   - Still needed: rendered output comparison.
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
   - Latest feature-remove opt-in sweep: add-remove-chart now passes 7 results
     across the pinned external-oracle pack by adding/removing a scratch-sheet
     chart subgraph and then removing the scratch sheet in staged saves.
   - Latest formula-translation opt-in sweep: move-formula-range now seeds a
     formula baseline on both sides of the audit and requires observed
     translated-formula semantic drift.
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
   - Latest drawing bugs found: scratch chart add/remove against ClosedXML
     prefixed worksheet XML exposed unprefixed `<worksheet>`, `<drawing>`, and
     `<legacyDrawing>` assumptions in drawing splice/remove helpers; the helpers
     now preserve the worksheet prefix.
   - Latest chart-remove composition bug found: same-save scratch chart removal
     plus scratch sheet deletion could leak queued chart deletions across
     worksheets/workbooks after an intermediate save error; chart deletions now
     live on the worksheet object and an external-oracle regression preserves
     source chart/drawing parts across all 7 pinned fixtures.
   - Latest empty-drawing bug found: adding then removing a chart in an NPOI workbook
     that already has an empty source drawing part deletes that empty drawing
     part and sheet drawing rel. Chart removal now preserves an empty source
     drawing shell when the chart was appended into that shell, while ordinary
     chart-only source removals still delete the drawing.
   - Latest oracle-hardening bug found: first-row/first-column delete and
     first-sheet copy correctly moved or duplicated data-validation ranges in
     Apache POI and ExcelJS fixtures, but the mutation runner treated
     `data_validations_semantic_drift` as unexpected. The runner now declares
     data-validation range movement as expected only for structural mutations,
     while feature-add mutations still require the new range marker.
   - Latest external-link evidence slice: the pinned external-oracle pack now
     includes `synthetic-external-link-basic.xlsx`, a targeted OOXML fixture
     with workbook external-link rels, external-link cached sheet data, and an
     external-reference formula.
   - Latest external-link oracle-hardening bug found: structural row/column
     delete and sheet copy intentionally change the external-reference formula
     set, which also changes the external-link semantic fingerprint because
     the fingerprint includes linked workbook formulas. The mutation runner
     now accepts `external_links_semantic_drift` for structural mutations only
     when the message shows `worksheet_formulas` drift; external-link target
     and cached-data drift remain unexpected.
   - Latest real Excel external-link evidence slice: the pinned
     external-oracle pack now includes
     `real-excel-external-link-basic.xlsx`, an Excel-authored workbook with
     workbook external-link rels, `xlExternalLinkPath/xlPathMissing`, an
     external-reference formula, and a calc chain.
   - Latest real Excel external-link bugs found: the external-link reader did
     not accept Excel's `xlExternalLinkPath/xlPathMissing` relationship type,
     and deleting the only formula removed `xl/calcChain.xml` while leaving
     workbook calc-chain metadata behind. Both are now covered by regression
     tests. The mutation runner classifies only calc-chain part/relationship
     removal as expected volatility for first-row/first-column deletion.
   - Latest real Excel chart/CF evidence slice: the pinned pack now includes
     `real-excel-chart-cf-basic.xlsx`, an Excel-saved workbook with a chart,
     chart style metadata, color-scale/data-bar/icon-set conditional
     formatting, drawing rels, styles, and theme parts.
   - Latest pivot evidence slice: the pinned pack now includes
     `real-excel-normalized-pivot-cf-table.xlsx`, a ClosedXML pivot/CF workbook
     opened and saved by Microsoft Excel. This clears the current strict P0
     evidence gate for pivot preservation, but it is not a substitute for a
     richer workbook authored from scratch in Excel with slicers, timelines,
     and pivot charts.
   - Latest active structural sweep: 70 results, 0 failures across the 14
     active external-oracle fixtures for first-row delete, first-column delete,
     first-sheet copy, first-sheet rename, and formula-range move.
   - Latest coverage audit result: `ready=true`. The hardened audit now
     separates slicer/timeline feature evidence from pivot-table evidence, and
     the current P0 evidence rows all have external-tool fixture evidence, real
     Excel fixture evidence, and structural mutation passes.
   - Latest app-smoke evidence slice: added
     `scripts/run_ooxml_app_smoke.py`. Microsoft Excel and LibreOffice both
     open/save the 14 active fixtures with 14 results and 0 failures.
     Microsoft Excel rejected the previous Apache POI and ExcelJS
     image/comment/table sources even after PNG CRC repair and LibreOffice
     normalization, so they are preserved under
     `tests/fixtures/external_oracle/rejected/` and replaced in the active
     pack with `openpyxl-table-validation-image-comment.xlsx`.
   - Latest Excelize pivot source result: Microsoft Excel rejected the original
     `excelize-sales-pivot-slicer-chart.xlsx` before WolfXL mutation. A
     LibreOffice-normalized copy now passes Excel app smoke and remains active
     for pivot/chart/style/color/table evidence, but the normalization removed
     slicer parts. Treat slicer/timeline coverage as not cleared.
   - Latest slicer fixture hunt: Excel's AppleScript dictionary exposes pivot
     creation APIs and a `slicer` class, but no obvious scripted "add slicer"
     command. A narrow local scan found only the ExcelBench Excelize slicer
     source/diff/modified workbooks, and Microsoft Excel rejects all three
     before WolfXL mutation. The next usable slicer/timeline fixture likely
     needs to be authored manually in Excel, captured from a trusted
     Excel-authored sample, or generated through a more capable Excel
     automation path.
   - Latest real Excel slicer evidence slice: the active pack now includes
     `real-excel-table-slicers.xlsx` from MyExcelOnline's public free practice
     workbook for Excel table slicers and
     `real-excel-pivot-chart-slicers.xlsx` from MyExcelOnline's public free
     practice workbook for pivot charts and slicers. Both workbooks report
     `Application=Microsoft Excel`, contain slicer/slicer-cache parts, pass
     Microsoft Excel and LibreOffice app smoke, and pass the structural
     mutation sweep. The pivot-chart slicer fixture also carries pivot table,
     pivot cache, chart, drawing, and table parts.
   - Latest external-tool slicer evidence slice: generated and pinned
     `excelize-2.10-pivot-slicers.xlsx` with github.com/xuri/excelize/v2
     v2.10.1, deterministic sample data, `AddPivotTable`, and `AddSlicer`.
     The workbook reports `Application=Go Excelize`, contains pivot, slicer,
     slicer-cache, table, drawing, style, and theme parts, passes Microsoft
     Excel app smoke, and clears the external-tool slicer evidence requirement
     in the strict coverage audit.
   - Latest real Excel timeline evidence slice: the active pack now includes
     `real-excel-timeline-slicer.xlsx` from MyExcelOnline's public free
     practice workbook for timeline slicers. The workbook reports
     `Application=Microsoft Excel`, contains `xl/timelines/timeline1.xml` and
     `xl/timelineCaches/timelineCache1.xml`, passes Microsoft Excel and
     LibreOffice app smoke, and passes the structural mutation sweep. The
     mutation runner now classifies timeline fingerprint changes from
     `copy_first_sheet` as expected structural drift, matching the existing
     chart, pivot, slicer, CF, data-validation, formula, and external-link
     sheet-copy handling.
   - Latest external-link oracle-hardening bug found: table structured
     references such as `Table1[REGION]` were being misclassified as external
     workbook formulas because the external-link fingerprint only looked for
     bracket characters. The audit now requires a workbook bracket followed by
     a sheet bang, and a regression test keeps structured references out of
     external-link evidence.
   - Latest bugs found: row insertion exposed a prefixed-XML end-tag corruption
     path in structural rewrites; range move exposed a prefixed `sheetData`
     discovery/re-emission gap. Both are now covered by regression tests.
   - Next mutations/evidence: add richer chart/pivot/slicer/timeline
     structural edits where the expected semantic drift can be declared, then
     move from package-level evidence into rendered comparison and broader
     real-file corpus sweeps.
3. Expand fixture sources:
   - richer native Excel-authored workbooks with slicers, timelines, pivot
     charts, chart style/color parts, and conditional-formatting extensions;
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
