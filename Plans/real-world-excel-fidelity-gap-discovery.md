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
| External-oracle fixture pack | 22 active pinned workbooks: 11 Excel-authored, 3 openpyxl, 2 ClosedXML, 2 Excelize, 1 LibreOffice-normalized Excelize, 1 NPOI, 1 synthetic OOXML external-link oracle, and 1 umya-spreadsheet fixture. The active pack now includes real Excel slicer, timeline, external-link, chart/CF, macro, control-prop, connection, and PowerPivot/data-model fixtures. The Apache POI and ExcelJS image/comment/table sources remain under `tests/fixtures/external_oracle/rejected/` because Excel rejects them before WolfXL mutation | Modify-save preserves important authored parts and still opens under safe value/style edits, structural edits, sheet copy/rename/remove, range moves, external-link relationship mutations, drawing/comment/object payload preservation, and durable workbook-global payload preservation | This is a strong current evidence gate, not exhaustive proof. It is still a curated corpus, not a random or customer-scale Excel corpus |
| New OOXML audit gate | `scripts/audit_ooxml_fidelity.py` now checks part loss, rel loss, dangling rels, content-type drift, feature part loss, CF dxf bounds, and deeper semantic fingerprints for charts, chart style/color parts, CF/x14 extensions, data validations, worksheet formulas, external links/cached data/formulas, workbook connections, PowerPivot data models, pivots, slicers, timelines, drawings/comments/embedded objects, style/theme/color dependencies, Python-in-Excel metadata, sheet metadata, and durable workbook-global payloads such as VBA, custom XML, and printer settings | The external-oracle pack now catches broken dependency graphs and feature-meaning drift across the named P0/P1 surfaces when those parts are present in the fixture. A recursive live SynthGL radar run found Python-in-Excel and sheet-metadata surfaces; those are now classified and semantically fingerprinted | It is not by itself a full Excel-rendered semantic validator, a proof of interactive slicer/timeline behavior, or a complete real-file corpus proof |
| Coverage evidence audit | `scripts/audit_ooxml_fidelity_coverage.py` maps fixtures plus mutation, render, and app reports to the evidence standard: external-tool fixture, real Excel fixture, structural mutation pass, no-op render pass, intentional mutation render pass, source app-open pass, and intentional app-open pass. It records concrete feature keys, requires slicer/timeline evidence separately from pivot-table evidence, rejects `--strict` runs that omit mutation reports, and supports recursive discovery, workbook application provenance fallback, optional emerging-surface gates, and render/app required gates | The current regenerated all-evidence coverage report is `/tmp/wolfxl-coverage-all-evidence-current-code-plus-excel-powerpivot.json`: `ready=true` over 22 fixtures, 13 surfaces, 4 mutation reports, 6 render reports, and 8 app reports. All named surfaces are clear, including PowerPivot/data model with targeted no-op render, intentional render, LibreOffice app-smoke, and Microsoft Excel source plus intentional app-smoke evidence. A recursive live SynthGL coverage run maps 32 nested Excel-authored workbooks to the latest 160-result mutation report and now tracks optional Python-in-Excel and sheet-metadata surfaces when present | It audits evidence presence and no-op/intentional renderability. It still does not prove Excel-rendered pixels, full visual semantic equivalence after every intentional edit, or broader real-file corpus coverage |
| Gap radar | `scripts/audit_ooxml_gap_radar.py` inventories unknown package part families, relationship types, content types, and hidden `ext uri` payloads inside known XML parts. The current pinned pack is clear across all four unknown buckets: 0 unknown part families, 0 unknown rel types, 0 unknown content types, and 0 unknown extension URIs. The recursive SynthGL fixture tree is also clear after classifying Python-in-Excel, sheet metadata, dynamic-array metadata, hidden-fill, and current chart extension URIs | The currently pinned corpus and 32-file live SynthGL fixture tree have no unclassified package-level or extension-URI surface according to the repo's known-surface allowlists | It only proves the currently scanned corpora are classified. A new real-world workbook can still introduce a new known-looking part with novel semantics or a future extension URI that must be triaged |
| Corpus diversity audit | `scripts/audit_ooxml_corpus_buckets.py` inventories workbook provenance and feature buckets across fixture directories or ad hoc workbook drops, including optional observed buckets for Python-in-Excel and sheet metadata | Current external-oracle run is `ready=true` over 22 workbooks with no missing buckets across Excel-authored, external-tool, macro/VBA, PowerPivot, slicer/timeline, embedded/control, external-link, chart, CF, table/validation, drawing/comment/media, and workbook-global coverage. Recursive SynthGL bucket audit sees 32 Excel-authored workbooks, including 1 Python-in-Excel workbook and 2 sheet-metadata workbooks | It proves bucket diversity, not behavioral preservation by itself. It should be paired with mutation, render, app-smoke, and gap-radar gates |
| App open/save smoke | `scripts/run_ooxml_app_smoke.py` opens and re-saves fixture packs through LibreOffice headless or Microsoft Excel, then validates the saved file as an OOXML ZIP. It accepts `--mutation` so intentionally edited workbooks can be app-smoked too | Existing app reports contribute to the all-evidence gate with source and intentional app-open coverage for every current surface | The active app-smoke pack is not a rendered comparison and does not prove every real-world workbook class. The mutation app-smoke evidence currently relies heavily on LibreOffice; Microsoft Excel mutation-smoke remains a stronger but slower optional gate |
| Interactive evidence audit | `scripts/audit_ooxml_interactive_evidence.py` inventories high-risk Excel behaviors that package/render/app-open checks cannot prove: slicer selection state, timeline selection state, pivot refresh/filter state, external-link update prompts, macro project presence, and embedded-control openability | Current strict run over the 22-fixture external-oracle pack is intentionally red: `/tmp/wolfxl-interactive-evidence-external-oracle.json` reports applicable fixtures for all six probe classes and no passing interactive probe reports yet | This is the current highest-confidence blocker before any stronger "no known interactive gaps" claim. It prevents app-open smoke from being mistaken for interactive behavior proof |
| Render comparison smoke | `scripts/run_ooxml_render_compare.py` no-op modify-saves each fixture with WolfXL, exports the original and saved workbook to PDF through LibreOffice, rasterizes pages with `pdftoppm`, and compares page images with ImageMagick RMSE. It also accepts `--mutation` for intentional edit render-smoke runs where the changed workbook is expected to render, but not expected to pixel-match the original. It supports recursive fixture discovery, deterministic page sampling for very large PDFs, and a byte-identical-xlsx short circuit for no-op saves | Existing render reports contribute no-op and intentional render evidence to the all-evidence gate. A broader live SynthGL recursive run passed 31 of 32 workbooks with full render RMSE 0.0; the remaining giant workbook has exact byte-identical `.xlsx` output after no-op save. Guarded mode passes all 32 live SynthGL workbooks | No-op render comparisons are pixel equality checks. Intentional mutation runs are renderability checks, not a proof that every intentional edit's visual result is semantically perfect. It still does not compare Excel-rendered pixels or interactive slicer/timeline state |
| Broader real-file corpus sweep | `scripts/run_ooxml_fidelity_mutations.py --recursive` can now walk nested workbook trees without flattening them first. Latest live SynthGL sweep covered `/Users/wolfgangschoenberger/Projects/SynthGL/tests/app/fixtures` recursively | 32 live SynthGL workbooks pass no-op, marker-cell, style-cell, rename-first-sheet, and move-formula-range audits in one 160-result sweep with 0 failures. Expected drift is limited to intentional style/theme changes from `style_cell` and formula-reference movement from `move_formula_range` | This is broader than the pinned oracle pack, but still not a full real-world Excel corpus; rendered comparison over this corpus and richer feature-aware mutations remain open |

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
   - Done: drawing/comment/embedded-object payload fingerprints, including
     VML comment drawings, drawing XML, media, embeddings, control props,
     ActiveX payloads, threaded comments, people, and related rels.
   - Done: durable workbook-global payload fingerprints for VBA binaries,
     custom XML, and printer settings. Calc-chain remains tracked as volatile
     package/relationship evidence rather than a byte-stable semantic payload.
   - Done: rendered output comparison and intentional render smoke through
     `scripts/run_ooxml_render_compare.py`.
   - Still needed: Excel-rendered pixel comparison, interactive slicer/timeline
     state validation, and broader real-file corpus diversity.
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
   - Latest active structural sweep: 98 results, 0 failures across the 14
     active external-oracle fixtures for first-row delete, first-column delete,
     first-sheet copy, first-sheet rename, scratch chart add/remove,
     conditional-formatting add, and formula-range move. This sweep exposed and
     fixed a real table-copy bug where Excelize-authored table parts with
     `<autoFilter>` children were re-emitted as malformed XML during
     `copy_worksheet`.
   - Latest coverage audit result: `ready=true` when run with the current
     structural mutation report. The render-required coverage audit is also
     `ready=true` when run with the active-pack no-op render report and the
     intentional mutation render report. The hardened audit now rejects
     `--strict` runs that omit mutation reports, supports `--render-report`
     plus `--require-render` and `--require-intentional-render`, separates
     slicer/timeline feature evidence from pivot-table evidence, and confirms
     the current P0 evidence rows all have external-tool fixture evidence, real
     Excel fixture evidence, structural mutation passes, no-op render passes,
     and intentional mutation render passes.
   - Earlier app-smoke evidence slice: added
     `scripts/run_ooxml_app_smoke.py`. Microsoft Excel and LibreOffice both
     open/save the 14 active fixtures with 14 results and 0 failures.
     Microsoft Excel rejected the previous Apache POI and ExcelJS
     image/comment/table sources even after PNG CRC repair and LibreOffice
     normalization, so they are preserved under
     `tests/fixtures/external_oracle/rejected/` and replaced in the active
     pack with `openpyxl-table-validation-image-comment.xlsx`.
   - Earlier intentional app-smoke slice: `scripts/run_ooxml_app_smoke.py`
     now accepts `--mutation` and writes mutation-scoped output artifacts. The
     active 14-fixture pack passes LibreOffice open/save smoke for marker-cell
     edit, first-sheet copy, and formula-range move with 42 results and 0
     failures.
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
   - Latest rendered comparison slice: added
     `scripts/run_ooxml_render_compare.py`, which renders before/after
     no-op modify-save pairs through LibreOffice PDF export, rasterizes pages
     with `pdftoppm`, and compares page images with ImageMagick RMSE. The
     active 14-fixture pack passes with 14 results, 0 failures, and max
     normalized RMSE 0.0 at 96 DPI.
   - Latest intentional-render slice: `scripts/run_ooxml_render_compare.py`
     now accepts `--mutation`. No-op runs remain pixel/RMSE comparisons;
     non-no-op mutation runs render-smoke the mutated workbook because visual
     differences are expected. The active 14-fixture pack passes marker-cell
     edit, first-sheet copy, and formula-range move render-smoke with 42
     results and 0 failures at 96 DPI, sampling large timeline PDFs to 8 pages
     per mutated workbook.
   - Latest broader render slice: the render comparison runner now supports
     recursive discovery, deterministic page sampling, and an explicit
     byte-identical-xlsx pass for no-op saves. The live SynthGL fixture tree
     passes no-op LibreOffice-render comparison for 31 of 32 workbooks, all
     with max RMSE 0.0. The remaining workbook,
     `number_memory/sigman_revenue_support.xlsx`, exports to a 14,926-page
     PDF, so exhaustive full-page rasterization is not practical at 96 DPI.
     It is not a visual drift failure: the before/after `.xlsx` files are
     byte-identical after no-op save, and a targeted 96-DPI sample of pages 1,
     7,463, and 14,926 compares at RMSE 0.0. Guarded recursive render mode now
     passes all 32 live SynthGL workbooks.
   - Latest broader-corpus evidence slice: `scripts/run_ooxml_fidelity_mutations.py`
     now supports `--recursive` discovery for nested fixture trees. The live
     SynthGL fixture tree at
     `/Users/wolfgangschoenberger/Projects/SynthGL/tests/app/fixtures` has
     32 `.xlsx` workbooks and passes a combined no-op, marker-cell,
     style-cell, rename-first-sheet, and move-formula-range sweep with 160
     results and 0 failures. `style_cell` and `move_formula_range` report
     expected drift only for the intentionally introduced style/theme and
     formula-reference changes.
   - Latest broader-corpus coverage slice:
     `scripts/audit_ooxml_fidelity_coverage.py --recursive --report` now maps
     nested live-corpus workbooks to mutation evidence and falls back to
     `docProps/app.xml` application metadata when manifest `tool` provenance is
     absent. Against the same SynthGL fixture tree and 160-result mutation
     report it discovers 32 real-Excel workbooks, confirms structural mutation
     evidence for the live-corpus surfaces that are present, and marks the
     optional Python-in-Excel and sheet-metadata surfaces clear when those parts
     appear. The report remains intentionally not-ready as a full evidence gate
     because that tree has no external-tool-authored workbooks and lacks
     external-link, PowerPivot, connection, macro/VBA, and slicer/timeline
     buckets.
   - Latest broader-corpus oracle-hardening bug found: a SynthGL workbook with
     drawing hyperlink targets like `#'Sheet1'!A1` exposed a false positive in
     the dangling-relationship audit. Fragment targets are in-workbook
     hyperlinks, not package parts; the audit now ignores relationship targets
     beginning with `#`, with a regression test.
   - Latest broader-corpus production bug found: `move_range` anchor rewrites
     parsed hyperlink attributes such as `location="'Sec. 1 &amp; 2 Notes'!A1"`
     and re-emitted unescaped ampersands, producing malformed worksheet XML.
     The range-move rewriter now XML-escapes rewritten anchor attributes, and
     the audit now reports malformed XML parts explicitly instead of masking
     them as secondary semantic drift.
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
     structural edits where the expected semantic drift can be declared, add
     rendered comparison for selected intentional structural edits, and broaden
     corpus size/source diversity beyond the current pinned pack plus SynthGL
     fixture tree.
   - Latest interactive-evidence audit slice:
     `scripts/audit_ooxml_interactive_evidence.py` now turns the remaining
     interactive blind spot into a strict, machine-readable gate. Current
     external-oracle output at
     `/tmp/wolfxl-interactive-evidence-external-oracle.json` is intentionally
     not ready: slicer selection state, timeline selection state, pivot refresh
     state, external-link update prompt behavior, macro project presence, and
     embedded-control openability all have applicable fixtures but no passing
     interactive probe reports yet.
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

## Completion audit - 2026-05-07

Objective restated as concrete deliverables:

1. Detect and preserve pivot/slicer/timeline dependencies across modify saves.
2. Detect and preserve chart/style/color/theme dependencies.
3. Detect and preserve conditional-formatting extension payloads.
4. Detect and preserve external-link and workbook relationship edge cases.
5. Detect and preserve P1 real-world Excel surfaces: tables/structured refs,
   validations, drawings/comments/embedded objects, workbook connections,
   PowerPivot data models, OOXML extension payloads, and workbook globals.
6. Maintain a discovery loop that can find newly introduced real-world Excel
   surface area, not just regressions in known fixtures.
7. Separate "currently covered surface is green" from "there are no possible
   unfound gaps."

Prompt-to-artifact checklist:

| Requirement | Repo artifact or command | Current evidence | Completion judgment |
|---|---|---|---|
| Pivot/slicer preservation across modify saves | `scripts/audit_ooxml_fidelity.py`, `scripts/audit_ooxml_fidelity_coverage.py`, `tests/fixtures/external_oracle/manifest.json` | Coverage report `/tmp/wolfxl-coverage-all-evidence-workbook-globals-payload-final.json`: `pivot_slicer_preservation` clear, 8 fixtures, external-tool and real-Excel sources present | Covered for pinned corpus; not exhaustive for every interactive slicer/timeline state |
| Chart/style/color dependency preservation | Chart, chart-sheet, chart-style, style/theme fingerprints plus render reports | Coverage report: `chart_style_color_preservation` clear, 5 fixtures, no missing evidence | Covered for pinned corpus; Excel-rendered pixels remain stronger future evidence |
| Conditional formatting extension preservation | CF/x14 semantic fingerprints and `dxfId` bounds | Coverage report: `conditional_formatting_extension_preservation` clear, 5 fixtures, no missing evidence | Covered for pinned corpus |
| External links and workbook relationship edge cases | Relationship graph audit, external-link target/cache/formula fingerprints, gap radar rel-type inventory | Coverage report: `external_link_relationship_edges` clear, 3 fixtures. Gap radar clear for unknown rel types | Covered for pinned corpus; new relationship types need radar triage |
| Workbook connections / query metadata | Connection fingerprint and real/openpyxl connection fixtures | Coverage report: `workbook_connections_query_metadata` clear, 3 fixtures | Covered for pinned corpus |
| Python-in-Excel and sheet metadata | Python and sheet-metadata semantic fingerprints plus optional coverage surfaces | Recursive SynthGL coverage report: Python-in-Excel metadata clear on 1 real-Excel fixture; sheet metadata clear on 2 real-Excel fixtures, each with structural mutation evidence | Good emerging-surface gate when those parts appear; broader Python-in-Excel workbook variants still need acquisition |
| PowerPivot / workbook data model | Data-model binary/default/content relationship fingerprint plus targeted render/app-smoke reports | Current all-evidence report: `powerpivot_data_model_preservation` clear on 1 real-Excel fixture, with structural mutation, no-op render, intentional render, LibreOffice source/intentional app-open, and Microsoft Excel source/intentional app-open evidence | Covered narrowly; needs more real-world data-model variants |
| OOXML extension payload preservation | Generic extension payload fingerprint plus extension-URI radar | Coverage report: `ooxml_extension_payload_preservation` clear, 16 fixtures. Gap radar clear for 0 unknown extension URIs | Covered for known extension URIs in pinned corpus |
| Tables / structured refs / validations | Table feature parts, structured-reference and validation fingerprints | Coverage report: `table_structured_refs_validations` clear, 11 fixtures | Covered for pinned corpus; richer table-filter/totals scenarios remain useful |
| Drawings / comments / embedded objects | Drawing-object fingerprint and VML structural fix | Coverage report: `drawings_comments_embedded_objects` clear, 15 fixtures. Drawing mutation sweep: 66 results, 0 failures | Covered for non-destructive object preservation; destructive geometry edits remain classified as expected drift, not unchanged-payload proof |
| Workbook global state | Defined names/protection/page setup plus durable package payload fingerprints | Coverage report: `workbook_global_state` clear, 21 fixtures. Workbook-global sweep: 88 results, 0 failures | Covered for durable payloads in pinned corpus; calc-chain is intentionally treated as volatile |
| Hidden future surface discovery | `scripts/audit_ooxml_gap_radar.py` | Latest radar: 22 fixtures, clear true, 0 unknown part families, rel types, content types, extension URIs | Good tripwire for new package/extension surface, not a semantic proof for every known-looking XML pattern |
| Broader recursive gap radar | `scripts/audit_ooxml_gap_radar.py /Users/wolfgangschoenberger/Projects/SynthGL/tests/app/fixtures --recursive --json --strict` | Latest run: 32 nested workbooks, clear true, 0 unknown part families, rel types, content types, extension URIs. First recursive run exposed Python-in-Excel and sheet-metadata surfaces that are now classified | Good live-corpus tripwire; still not a replacement for richer external workbook acquisition |
| Corpus diversity | `scripts/audit_ooxml_corpus_buckets.py tests/fixtures/external_oracle --json --strict` | Latest run: `ready=True`, 22 workbooks, no missing buckets | Good coverage-shape check; must be paired with behavioral gates |
| Broader recursive coverage audit | `scripts/audit_ooxml_fidelity_coverage.py /Users/wolfgangschoenberger/Projects/SynthGL/tests/app/fixtures --recursive --report /tmp/wolfxl-ooxml-fidelity-mutations-synthgl-python-metadata-final/report.json` plus `scripts/audit_ooxml_corpus_buckets.py ... --recursive --json` | Latest coverage run: 32 fixtures, recursive true, 1 mutation report, all source classes inferred as real Excel, structural mutation evidence attached to present live-corpus surfaces, Python metadata clear on 1 fixture, sheet metadata clear on 2 fixtures. Latest corpus-bucket run: 32 Excel-authored workbooks, missing external-link, external-tool-authored, macro/VBA, PowerPivot, and slicer/timeline buckets | Good check that broad live fixtures are not silently excluded; proves the SynthGL tree is useful side evidence but not a replacement for the pinned external-oracle pack |
| Interactive behavior evidence | `scripts/audit_ooxml_interactive_evidence.py tests/fixtures/external_oracle --strict` | Latest run: `ready=False`; applicable fixture candidates exist for slicer selection, timeline selection, pivot refresh, external-link prompt/cache behavior, macro project presence, and embedded-control openability, with no passing interactive probe reports yet | Not covered. This is the clearest remaining evidence gap after package/render/app-open gates |
| Whole-pack preservation under common edits | `tests/test_external_oracle_preservation.py` | `198 passed` after the latest changes | Strong pinned-pack regression gate |
| Combined all-evidence gate | `scripts/audit_ooxml_fidelity_coverage.py --strict --require-render --require-intentional-render --require-app --require-intentional-app` | Latest regenerated report `/tmp/wolfxl-coverage-all-evidence-current-code-plus-excel-powerpivot.json`: `ready=True`, 22 fixtures, 13 surfaces, 4 mutation reports, 6 render reports, 8 app reports | Strong current-state gate |

Current conclusion:

- The repo can honestly claim: **no known fidelity gap in the currently pinned
  and classified real-world OOXML surface.**
- The repo should not claim: **no real-world Excel fidelity gaps exist.** That
  would require broader corpus diversity, Excel-rendered pixel comparison,
  interactive slicer/timeline validation, and more adversarial feature-specific
  mutations than the current pack can provide.

Next evidence slices before declaring a higher-confidence "no known gaps":

1. Add a larger external workbook corpus sweep with provenance buckets:
   Excel-authored, Microsoft-template, finance-model, BI/reporting, macro,
   PowerPivot, slicer/timeline, embedded-object/control, and external-link
   workbooks.
2. Add a Microsoft Excel render/app mutation gate for a small representative
   subset, especially slicer/timeline, PowerPivot, macro, external-link, and
   embedded-object workbooks.
3. Add interactive state probes where package/render evidence is weak:
   slicer selections, timeline selections, pivot refresh state, external-link
   update prompts, macro/project presence, and embedded-control openability.
4. Keep the gap radar strict: every newly seen part family, relationship type,
   content type, or extension URI must become either an allowlisted known
   surface with a semantic fingerprint or an explicit gap.
