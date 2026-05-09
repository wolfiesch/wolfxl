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
| New OOXML audit gate | `scripts/audit_ooxml_fidelity.py` now checks part loss, rel loss, dangling rels, content-type drift, feature part loss, CF dxf bounds, worksheet formulas, data validation formulas, conditional formatting formulas, workbook defined names, chart formulas, pivot-chart sources, and pivot-cache `worksheetSource` attributes pointing at missing local sheets, and deeper semantic fingerprints for charts, chart style/color parts, CF/x14 extensions, data validations, worksheet formulas, external links/cached data/formulas, workbook connections, PowerPivot data models, pivots, slicers, timelines, drawings/comments/embedded objects, style/theme/color dependencies, Python-in-Excel metadata, sheet metadata, and durable workbook-global payloads such as VBA, custom XML, and printer settings | The external-oracle pack now catches broken dependency graphs, internal worksheet/DV/CF formulas, workbook defined-name, chart formula, pivot-source, and pivot-cache source references to missing sheets, and feature-meaning drift across the named P0/P1 surfaces when those parts are present in the fixture. The latest `rename_first_sheet` audit run `/tmp/wolfxl-ooxml-fidelity-mutations-rename-sheet-defined-name-audit-20260508/report.json` has 22 results and 0 failures under the worksheet/workbook/chart/pivot source reference checks. A recursive live SynthGL radar run found Python-in-Excel and sheet-metadata surfaces; those are now classified and semantically fingerprinted | It is not by itself a full Excel-rendered semantic validator, a proof of interactive slicer/timeline behavior, or a complete real-file corpus proof |
| Coverage evidence audit | `scripts/audit_ooxml_fidelity_coverage.py` maps fixtures plus mutation, render, and app reports to the evidence standard: external-tool fixture, real Excel fixture, structural mutation pass, no-op render pass, intentional mutation render pass, source app-open pass, and intentional app-open pass. It records concrete feature keys, requires slicer/timeline evidence separately from pivot-table evidence, rejects `--strict` runs that omit mutation reports, and supports recursive discovery, workbook application provenance fallback, optional emerging-surface gates, render/app required gates, and engine-specific render requirements | The current regenerated all-evidence coverage report is `/tmp/wolfxl-coverage-all-evidence-current-code-plus-excel-powerpivot.json`: `ready=true` over 22 fixtures, 13 surfaces, 5 mutation reports, 6 render reports, and 9 app reports, with `retarget_external_links` accepted as direct structural evidence for external-link relationship edges. The Excel-render full-pack coverage report with marker-cell, copy-first-sheet, rename-first-sheet, move-formula-range, add-data-validation, add-conditional-formatting render-equivalence, add-remove-chart render-equivalence, retarget-external-links, and delete-first-row/column intentional smoke `/tmp/wolfxl-coverage-excel-render-full-pack-with-rename-sheet-defined-name-20260508.json` is also `ready=true`: 22 fixtures, `render_engine_required=excel`, 10 Excel-render reports, 0 unclear surfaces. The refreshed external-link retarget mutation report `/tmp/wolfxl-ooxml-fidelity-mutations-external-link-retarget.json` has 22 results and 0 failures, with required retarget drift observed on the three external-link-bearing fixtures and matching formula-token rewrites where formulas name the linked workbook. The matching Microsoft Excel app-open smoke `/tmp/wolfxl-app-smoke-excel-retarget-external-links.json` has 22 results and 0 failures, and clears intentional app-open coverage for those three external-link fixtures. A recursive live SynthGL coverage run maps 32 nested Excel-authored workbooks to the latest 160-result mutation report and now tracks optional Python-in-Excel and sheet-metadata surfaces when present | It audits evidence presence and no-op/intentional renderability. It now has full-pack Microsoft Excel-rendered no-op, marker-cell, style-cell, insert-tail-row/column, and delete-marker-tail-row intentional edits with expected visual-delta evidence, copy-first-sheet structural intentional, rename-first-sheet structural intentional, move-formula-range structural intentional, delete-first-row/column structural intentional, add-data-validation feature-add, add-conditional-formatting feature-add, add-remove-chart feature add/remove with render-equivalence, external-link retarget package-audit evidence, external-link retarget Excel app-open evidence, and external-link retarget Excel render-equivalence evidence, but still does not prove click-level structural visual equivalence after every intentional edit, click-level interactive behavior, or broader real-file corpus coverage |
| Gap radar | `scripts/audit_ooxml_gap_radar.py` inventories unknown package part families, relationship types, content types, hidden `ext uri` payloads inside known XML parts, and expected app-unsupported features such as PowerView. The current pinned pack is clear across all package unknown buckets. The recursive SynthGL fixture tree is also clear after classifying Python-in-Excel, sheet metadata, dynamic-array metadata, hidden-fill, and current chart extension URIs. A 61-workbook umya-spreadsheet issue corpus is now also clear after classifying sensitivity-label metadata, JavaScript API project payloads, WMF thumbnails, non-numbered theme parts, x14 data-validation extensions, drawing hidden-effect extensions, and chart pivot-option extensions. The 71-workbook SynthGL Codex Spark archive is clear at `/tmp/wolfxl-gap-radar-synthgl-codex-spark-archive-71-classified-20260508.json` after classifying named sheet views, chartEx relationships, and modern worksheet/pivot extension URIs. The 28-readable-workbook SynthGL real-world-ingestion corpus is clear at `/tmp/wolfxl-gap-radar-synthgl-real-world-ingestion-20260509.json` after classifying Excel `[trash]/0000.dat` package debris. The 16-workbook SynthGL CDS case-study corpus is clear at `/tmp/wolfxl-gap-radar-synthgl-cds-case-study-20260509.json`. The 52-readable-workbook Calamine reader corpus is clear at `/tmp/wolfxl-gap-radar-calamine-tests-20260508.json` with 3 explicit invalid or unsupported source skips | The currently pinned corpus, 32-file live SynthGL fixture tree, 61-file umya issue corpus, 71-file SynthGL archive, 28-readable-file SynthGL real-world-ingestion corpus, 16-file SynthGL CDS case-study corpus, and 52-readable-file Calamine reader corpus have no unclassified package-level or extension-URI surface according to the repo's known-surface allowlists | It only proves the currently scanned corpora are classified. A new real-world workbook can still introduce a new known-looking part with novel semantics or a future extension URI that must be triaged |
| Corpus diversity audit | `scripts/audit_ooxml_corpus_buckets.py` inventories workbook provenance and feature buckets across fixture directories or ad hoc workbook drops, including optional observed buckets for Python-in-Excel and sheet metadata. `scripts/audit_ooxml_corpus_portfolio.py` now aggregates those bucket reports into a portfolio-level diversity gate | Current external-oracle run is `ready=true` over 22 workbooks with no missing buckets across Excel-authored, external-tool, macro/VBA, PowerPivot, slicer/timeline, embedded/control, external-link, chart, CF, table/validation, drawing/comment/media, and workbook-global coverage. Recursive SynthGL bucket audit sees 32 Excel-authored workbooks, including 1 Python-in-Excel workbook and 2 sheet-metadata workbooks. The umya-spreadsheet issue corpus adds 63 workbooks with 53 Excel-authored, 9 external-tool-authored, 2 macro/VBA, 2 external-link, 6 chart, 21 drawing/comment/media, 21 embedded/control, and 2 sheet-metadata examples, but no PowerPivot or slicer/timeline examples. The SynthGL real-world-ingestion corpus adds 28 readable Excel-authored workbooks with 28 CF-extension/workbook-global examples, 12 drawing/media and embedded/control examples, 10 table/validation examples, 2 external-link examples, and 1 sheet-metadata example. The SynthGL CDS case-study corpus adds 16 Excel-authored workbooks with 1 chart workbook, 16 CF-extension/workbook-global examples, 2 drawing/media and embedded/control examples, 2 sheet-metadata examples, and 2 table/validation examples. The Calamine reader corpus adds 52 readable workbooks with external-tool, macro/VBA, external-link, chart, table/validation, drawing/media, embedded/control, and workbook-global examples, but no PowerPivot or slicer/timeline examples. Latest portfolio report `/tmp/wolfxl-corpus-portfolio-buckets-20260509.json` is `ready=true` across 15 source reports, 236 unique readable workbooks, and all required diversity buckets: 210 Excel-authored, 26 external-tool-authored, 5 macro/VBA, 2 PowerPivot/data-model, 7 slicer/timeline, 11 external-link, 16 chart, and 39 table/validation examples | It proves bucket diversity, not behavioral preservation by itself. The portfolio is broader than any single corpus, but it is still curated and not random/customer-scale. It should be paired with mutation, render, app-smoke, and gap-radar gates |
| App open/save smoke | `scripts/run_ooxml_app_smoke.py` opens fixture packs through LibreOffice headless or Microsoft Excel, verifies the generated file is an OOXML ZIP, and now rejects false positives where Excel leaves an unrelated active workbook such as `Book1` or `missing value`. It accepts `--mutation` so intentionally edited workbooks can be app-smoked too | Existing app reports contribute to the all-evidence gate with source and intentional app-open coverage for every current surface. The hardened Excel app-open coverage gate `/tmp/wolfxl-coverage-excel-app-open-full-pack-with-cf-verified.json` is also `ready=true` over 22 fixtures using source, marker-cell, and add-conditional-formatting full-pack Microsoft Excel reports with 0 failures. The external-link retarget Excel app-open report `/tmp/wolfxl-app-smoke-excel-retarget-external-links.json` is green over all 22 fixtures, including the three external-link-bearing retarget cases | The active app-smoke pack is not a rendered comparison and does not prove every real-world workbook class. The hardened gate prevented false positives while exposing the now-fixed conditional-formatting styles-order/prefix bug |
| Interactive evidence audit | `scripts/audit_ooxml_interactive_evidence.py` inventories high-risk Excel behaviors that package/render/app-open checks cannot fully prove: slicer state, timeline state, pivot refresh/filter state, external-link update prompts, macro project presence, and embedded-control openability. `scripts/run_ooxml_interactive_probe.py` implements Microsoft Excel probes for all six classes as `ooxml_state_presence` checks: Excel opens/saves the workbook and the relevant OOXML state remains present. It also has a separate `--probe-kind excel_ui_interaction` mode for targeted UI actions where the local Excel UI can expose them; external-link UI probes can either temporarily force Excel's update-link prompt setting or use the current setting unchanged. | Current strict state-presence run over the 22-fixture external-oracle pack with all six probe reports is green: `/tmp/wolfxl-interactive-evidence-external-oracle-all-20260508.json` reports `ready=true`. UI-interaction artifacts are green for macro security prompt handling (`/tmp/wolfxl-ui-interaction-macro-20260508/interactive-probe-report.json`, clicked `Disable Macros`), pivot refresh (`/tmp/wolfxl-ui-interaction-pivot-20260508/interactive-probe-report.json`, executed Excel `refresh all`), embedded list-box click persistence (`/tmp/wolfxl-ui-interaction-control-click-20260508/interactive-probe-report.json`, clicked `List Box 1`, saved, and verified `ctrlProp1.xml` changed from `sel=0` to `sel=2`), adjacent issue-corpus list-box click persistence (`/tmp/wolfxl-ui-interaction-umya-listbox-control-20260508/interactive-probe-report.json`, clicked `List Box 1` in `issue_208.xlsx`, saved, and verified persisted control state changed), stateless button-control click persistence (`/tmp/wolfxl-ui-interaction-button-control-20260508/interactive-probe-report.json`, clicked `Button 1`, saved, and verified the button control parts remained present), table-slicer all-item-click persistence (`/tmp/wolfxl-ui-interaction-slicer-all-items-20260508/interactive-probe-report.json`, selected `REGION`, clicked the first visible item in both `REGION` and `YEAR`, saved, and verified `Table1` persisted `EAST` and `2014` filters), pivot-chart slicer click persistence (`/tmp/wolfxl-ui-interaction-pivot-chart-slicer-extra-20260508/interactive-probe-report.json`, selected `REGION`, clicked the first visible item in both `REGION` and `YEAR`, saved, and verified slicer state changed), external-tool pivot-slicer click persistence (`/tmp/wolfxl-ui-interaction-excelize-pivot-slicer-20260508/interactive-probe-report.json`, selected `Month`, clicked the first visible item in both `Month` and `Region`, saved, and verified slicer state changed), shared pivot-slicer cache persistence (`/tmp/wolfxl-ui-interaction-shared-slicer-cache-20260508/interactive-probe-report.json`, selected `SALES MONTH`, clicked one item, saved, and verified the slicer-cache item state changed), timeline month-click persistence (`/tmp/wolfxl-ui-interaction-timeline-click-20260508/interactive-probe-report.json`, selected `ORDER DATE`, clicked `May`, saved, and verified the persisted timeline selection changed to May 2012), forced external-link prompt handling (`/tmp/wolfxl-ui-interaction-external-link-review-fix-20260508/interactive-probe-report.json`, temporarily enabled Excel's update-link prompt setting, clicked `Don't Update`, and restored the prior setting), current-setting external-link prompt handling (`/tmp/wolfxl-ui-interaction-external-link-current-prompt-20260508/interactive-probe-report.json`, left Excel's prompt setting unchanged, clicked `Don't Update`, and verified the external-link parts remained present), and adjacent issue-corpus current-setting external-link prompt handling (`/tmp/wolfxl-ui-interaction-umya-external-link-current-prompt-20260508/interactive-probe-report.json`, opened `issue_297.xlsx`, left Excel's prompt setting unchanged, clicked `Don't Update`, and verified external-link parts remained present) | This prevents app-open smoke or package-presence checks from being mistaken for interactive behavior proof. The UI-interaction mode is partial: two embedded list-box clicks, one stateless button-control click, table-slicer item clicks, Microsoft-authored pivot-chart slicer item clicks, external-tool-authored pivot-slicer item clicks, one shared pivot-slicer cache click, one timeline date-range click, and forced plus current-setting external-link prompt paths across two fixtures now have real-Excel proof; broader control/slicer/timeline variants remain open |
| Render comparison smoke | `scripts/run_ooxml_render_compare.py` no-op modify-saves each fixture with WolfXL, exports the original and saved workbook to PDF through LibreOffice or Microsoft Excel, rasterizes pages with `pdftoppm`, and compares page images with ImageMagick RMSE. It records `render_engine`; `scripts/audit_ooxml_fidelity_coverage.py --require-render-engine excel` can reject LibreOffice reports when the desired proof is specifically Microsoft Excel-rendered output. Intentional mutations are render-smoked instead of pixel-compared against the original workbook, because their visual output is expected to change | Existing LibreOffice render reports contribute no-op and intentional render evidence to the all-evidence gate. The Excel renderer now stages files inside Excel's sandbox container to avoid macOS `Grant File Access` prompts. Latest Excel no-op render report `/tmp/wolfxl-render-excel-full-pack.json`: `render_engine=excel`, 22 workbooks, 0 failures, max RMSE 0.0. Latest Excel intentional marker-cell render report `/tmp/wolfxl-render-excel-intentional-marker-full-pack.json`: 22 workbooks, 0 failures. Latest marker-cell render-delta audit `/tmp/wolfxl-marker-cell-render-delta-20260509.json`: `ready=true`, 19 changed, 3 sampled pages unchanged, 0 failures, 0 inconclusive. Latest Excel style-cell render report `/tmp/wolfxl-render-excel-intentional-style-cell-full-pack-20260509.json`: 22 workbooks, 0 failures. Latest style-cell render-delta audit `/tmp/wolfxl-style-cell-render-delta-20260509.json`: `ready=true`, 19 changed, 3 sampled pages unchanged, 0 failures, 0 inconclusive. Latest Excel insert-tail-row render report `/tmp/wolfxl-render-excel-intentional-insert-tail-row-full-pack-20260509.json`: 22 workbooks, 0 failures. Latest insert-tail-row render-delta audit `/tmp/wolfxl-insert-tail-row-render-delta-20260509.json`: `ready=true`, 21 changed, 1 sampled page unchanged, 0 failures, 0 inconclusive. Latest Excel insert-tail-column render report `/tmp/wolfxl-render-excel-intentional-insert-tail-col-full-pack-20260509.json`: 22 workbooks, 0 failures. Latest insert-tail-column render-delta audit `/tmp/wolfxl-insert-tail-col-render-delta-20260509.json`: `ready=true`, 19 changed, 3 sampled pages unchanged, 0 failures, 0 inconclusive. Latest Excel delete-marker-tail-row render report `/tmp/wolfxl-render-excel-intentional-delete-marker-tail-row-full-pack-20260509.json`: 22 workbooks, 0 failures. Latest delete-marker-tail-row render-delta audit `/tmp/wolfxl-delete-marker-tail-row-render-delta-20260509.json`: `ready=true`, 1 changed, 21 sampled pages unchanged, 0 failures, 0 inconclusive; the changed case is the chart-heavy fixture where deleting a tail row intentionally adjusts a drawing anchor. Latest Excel structural intentional copy-first-sheet render report `/tmp/wolfxl-render-excel-intentional-copy-sheet-full-pack.json`: 22 workbooks, 0 failures. Latest Excel structural intentional rename-first-sheet render report `/tmp/wolfxl-render-excel-intentional-rename-first-sheet-defined-name-fixed-20260508.json`: 22 workbooks, 0 failures after fixing existing worksheet formulas, chart `<c:f>`, chart `<c:pivotSource><c:name>`, workbook `<definedName>`, and pivot-cache `<worksheetSource sheet>` rewrites. Latest rename-sheet render-equivalence audit `/tmp/wolfxl-rename-sheet-render-equivalence-first-last-pages-20260509.json`: `ready=true`, 22 passed, 0 failures, 0 inconclusive, using first-and-last page sampling where needed. Latest Excel structural intentional move-formula-range render report `/tmp/wolfxl-render-excel-intentional-move-formula-range-full-pack.json`: 22 workbooks, 0 failures. Latest move-formula-range render-delta audit `/tmp/wolfxl-move-formula-range-render-delta-20260509.json`: `ready=true`, 19 changed, 3 sampled pages unchanged, 0 failures, 0 inconclusive, proving this mutation is not a visually neutral equivalence case. Latest Excel add-data-validation render report `/tmp/wolfxl-render-excel-intentional-add-data-validation-full-pack.json`: 22 workbooks, 0 failures. Latest add-data-validation render-equivalence audit `/tmp/wolfxl-add-data-validation-render-equivalence-20260509.json`: `ready=true`, 22 passed, 0 failures, 0 inconclusive, proving this visually neutral edit does not change sampled Excel-rendered pages. Latest Excel add-conditional-formatting render report `/tmp/wolfxl-render-excel-intentional-add-conditional-formatting-full-pack-fixed.json`: 22 workbooks, 0 failures. Latest add-conditional-formatting render-equivalence audit `/tmp/wolfxl-add-conditional-formatting-render-equivalence-20260509.json`: `ready=true`, 22 passed, 0 failures, 0 inconclusive, max RMSE 0.0. Latest Excel add-remove-chart render report `/tmp/wolfxl-render-excel-intentional-add-remove-chart-20260508.json`: 22 workbooks, 0 failures. Latest add-remove-chart render-equivalence audit `/tmp/wolfxl-add-remove-chart-render-equivalence-20260509.json`: `ready=true`, 22 passed, 0 failures, 0 inconclusive. Latest Excel external-link retarget render report `/tmp/wolfxl-render-excel-retarget-external-links-first-last-pages-20260509.json`: 22 workbooks, 0 failures. Latest external-link retarget render-equivalence audit `/tmp/wolfxl-retarget-external-links-render-equivalence-first-last-pages-20260509.json`: `ready=true`, 22 passed, 0 failures, 0 inconclusive, after fixing linked-workbook formula-token rewrites so the synthetic cached-value case renders `99` instead of `#REF!`. Latest Excel delete-first-row/column render report `/tmp/wolfxl-render-excel-intentional-delete-first-axis-20260508.json`: 44 results, 0 failures. Latest delete-first-row render-delta audit `/tmp/wolfxl-delete-first-row-render-delta-20260509.json`: `ready=true`, 21 changed, 1 sampled page unchanged, 0 failures, 0 inconclusive. Latest delete-first-column render-delta audit `/tmp/wolfxl-delete-first-col-render-delta-20260509.json`: `ready=true`, 21 changed, 1 sampled page unchanged, 0 failures, 0 inconclusive. Latest recursive SynthGL no-op report `/tmp/wolfxl-render-excel-synthgl-recursive-byte-identical-20260508/render-compare-report.json`: `render_engine=excel`, 32 workbooks, 0 failures, with every no-op output byte-identical to source. Latest engine-specific coverage report `/tmp/wolfxl-coverage-excel-render-full-pack-with-rename-sheet-defined-name-20260508.json`: `ready=true`, 10 Excel render reports, 0 unclear surfaces | No-op render comparisons are pixel equality checks when bytes differ; byte-identical no-op outputs are stronger package-level evidence and avoid renderer nondeterminism. The rename-sheet slice has feature-specific before/after pixel-equivalence proof, copy-first-sheet has copied-page equivalence proof, external-link retargeting has before/after Excel pixel-equivalence proof for the sampled pages, add-data-validation, add-conditional-formatting, and add-remove-chart have before/after Excel pixel-equivalence proof for the full pinned pack, and marker-cell, style-cell, insert-tail-row/column, delete-marker-tail-row, move-formula-range, plus first row/column deletion now have expected visual-delta proof instead of being misclassified as visually neutral. Other intentional mutation runs remain renderability checks, not proof that every intentional edit's visual result is semantically perfect. The engine-specific gate prevents LibreOffice pixels from being mistaken for Excel pixels. This is full-pack no-op plus marker-cell/style-cell/insert-tail-row-column/delete-marker-tail-row visual-delta, copy-first-sheet, rename-first-sheet, move-formula-range, delete-first-row/column structural, add-data-validation, add-conditional-formatting render-equivalence, add-remove-chart render-equivalence, and retarget-external-links Excel-render evidence, plus recursive SynthGL no-op byte-identical evidence; it is not exhaustive intentional-edit visual equivalence |
| Broader real-file corpus sweep | `scripts/run_ooxml_fidelity_mutations.py --recursive` can now walk nested workbook trees without flattening them first. Latest live SynthGL sweep covered `/Users/wolfgangschoenberger/Projects/SynthGL/tests/app/fixtures` recursively; latest archive sweep covered `/Users/wolfgangschoenberger/Projects/SynthGL/data/archives/codex_spark_excel_copies_dedup_20260214_030712/files` | 32 live SynthGL workbooks pass no-op, marker-cell, style-cell, rename-first-sheet, and move-formula-range audits in one 160-result sweep with 0 failures. A newer richer SynthGL sweep at `/tmp/wolfxl-ooxml-fidelity-mutations-synthgl-richer-fixed-20260508.json` covers add-data-validation, add-conditional-formatting, copy-first-sheet, delete-first-row, delete-first-col, and add-remove-chart across the same 32 workbooks: 192 results, 0 failures. The focused worksheet-formula rename fix was rechecked against `/Users/wolfgangschoenberger/Projects/SynthGL/tests/app/fixtures` for 32 rename-first-sheet results and 0 failures, and the adjacent umya-spreadsheet issue-corpus rename sweep now passes 63 results with 0 failures at `/tmp/wolfxl-ooxml-fidelity-mutations-umya-rename-after-macro-sidecar-20260508.json` after the macro-sidecar preservation fix. The richer sweep exposed and then verified the fix for copied-sheet hyperlink attribute escaping on `ingestion/time_series/omb_historical_budget_2026.xlsx`. The matching recursive no-op render-equivalence run over the same 32 workbooks is green because each WolfXL no-op output is byte-for-byte identical to its source. The 71-workbook SynthGL archive quick sweep originally reported 16 failures in `/tmp/wolfxl-ooxml-fidelity-mutations-synthgl-codex-spark-archive-71-quick-20260508/report.json`; re-auditing those before/after pairs after suppressing exact source-preexisting XML/reference defects yields `/tmp/wolfxl-ooxml-fidelity-mutations-synthgl-codex-spark-archive-71-quick-reaudit-20260508.json` with 142 passed results and 0 failures. The suppressed defects were stale workbook defined names and malformed pivot style XML already present in the source files. The SynthGL real-world-ingestion mixed-validity corpus adds `/tmp/wolfxl-ooxml-fidelity-mutations-synthgl-real-world-ingestion-quick-20260509.json` with 135 rows, 0 failures, 84 source-valid passes or expected-drift passes, and 51 explicit invalid-source skips, plus `/tmp/wolfxl-ooxml-fidelity-mutations-synthgl-real-world-ingestion-structural-20260509.json` with 225 rows, 0 failures, 140 source-valid passes or expected-drift passes, and 85 explicit invalid-source skips. The SynthGL CDS case-study corpus adds `/tmp/wolfxl-ooxml-fidelity-mutations-synthgl-cds-case-study-quick-20260509.json` with 48 rows, 0 failures, 46 passed, and 2 expected-drift passes, plus `/tmp/wolfxl-ooxml-fidelity-mutations-synthgl-cds-case-study-structural-20260509.json` with 80 rows, 0 failures, 13 passed, and 67 expected-drift passes. The Calamine reader corpus adds `/tmp/wolfxl-ooxml-fidelity-mutations-calamine-tests-fixed-20260508/report.json` with 495 default mutation rows, 0 failures, 416 passed, 52 expected style/theme drift passes, and 27 explicit invalid-source skips | This is broader than the pinned oracle pack and now includes richer feature-aware mutations plus archived, ingestion-real-file, CDS case-study, and Calamine reader-corpus pressure tests, but still not a full real-world Excel corpus; additional non-SynthGL real-file sources remain open |
| Current evidence bundle | `Plans/ooxml-current-evidence-bundle.json` plus `scripts/audit_ooxml_evidence_bundle.py --strict` pins the current generated report artifacts, records each producer command, and verifies expected readiness/counts across all-evidence coverage, interactive evidence, aggregated Excel UI-interaction evidence, Excel-render full-pack no-op and intentional evidence, feature-specific render-equivalence and render-delta evidence, hardened Excel app-open evidence, corpus diversity, gap radar, recursive SynthGL side evidence, umya issue-corpus side evidence, umya `tests/result_files` macro-sidecar mutation evidence, the umya issue-corpus rename rerun, public Microsoft sample evidence, the Contoso PowerPivot plus shared-slicer sidecar evidence, sidecar corpus-bucket reports, shared-slicer UI-interaction side evidence, SynthGL real-world-ingestion side evidence, SynthGL CDS case-study side evidence, Calamine reader-corpus side evidence, ILPA reporting-template side evidence, WBD wDesk side evidence, BF30 public-download side evidence, blind-holdout side evidence, rescue-download side evidence, SEC/EDGAR side evidence, Iran OSINT side evidence, and the aggregate corpus-portfolio diversity gate | Latest bundle audit `/tmp/wolfxl-current-evidence-bundle-audit-delete-marker-tail-row-delta-20260509.json`: `ready=true`, 138 reports, 138 producer commands, 0 issues | This prevents stale or origin-unclear `/tmp` report references from silently supporting completion claims. It still verifies generated artifacts, not a permanent corpus expansion by itself |

## Latest sidecar corpus update

On 2026-05-08, the mutation runner's manifest discovery was widened beyond
WolfXL's native `fixtures` schema. It now also accepts ExcelBench-style
`files` manifests and curated `workbooks` manifests, which prevents sidecar
corpora from silently scanning as empty when their manifest schema differs.
For exploratory corpora, `--skip-invalid-source` now treats manifest entries
whose files are absent as explicit `skipped_source_missing` rows instead of
counting placeholder paths as fidelity failures.

| Corpus | Scan result | Mutation result | Remaining corpus holes |
|---|---:|---:|---|
| ExcelBench core fixtures, `/Users/wolfgangschoenberger/Projects/ExcelBench/fixtures/excel` | 19 workbooks, gap radar clear | 171/171 default mutation results passed | Missing macro/VBA, PowerPivot/data model, slicer/timeline, external-link, external-tool-authored, and chart coverage in this corpus |
| ExcelBench external validated fixtures, `/Users/wolfgangschoenberger/Projects/ExcelBench/results_dev_external/fixtures/validated` | 7 workbooks, gap radar clear | 63/63 default mutation results passed | Missing macro/VBA, PowerPivot/data model, slicer/timeline, and external-link coverage in this corpus |
| ExcelBench curated real-world manifest, `/Users/wolfgangschoenberger/Projects/ExcelBench/fixtures/real_world` | 8 existing workbooks, gap radar clear | 72/72 existing-workbook default mutation results passed; 2 manifest placeholders skipped as missing | This is a first-pass curated/fixture-backed set, not broad customer real-world evidence; missing macro/VBA, PowerPivot/data model, slicer/timeline, external-link, and chart coverage |
| Public OSINT data workbooks, `/Users/wolfgangschoenberger/Projects/iran/osint/data` | 16 Microsoft Excel-authored workbooks, gap radar clear at `/tmp/wolfxl-gap-radar-iran-osint-20260509.json`; bucket report pinned at `/tmp/wolfxl-corpus-buckets-iran-osint-20260509.json` | 144/144 default mutation results passed at `/tmp/wolfxl-ooxml-fidelity-mutations-iran-osint-20260508` | Adds non-SynthGL public workbook diversity for conditional-formatting extensions, extension payloads, workbook-global state, and one drawing/embedded-object workbook; still missing external-tool-authored, macro/VBA, PowerPivot/data model, slicer/timeline, external-link, chart, and table/validation buckets |
| Calamine reader test corpus, `/Users/wolfgangschoenberger/Projects/calamine/tests` | 52 readable OOXML workbooks; gap radar clear with 3 invalid or unsupported source skips (`issue252.xlsx`, `issue_530.xlsx`, `pass_protected.xlsx`) | 495 default mutation rows at `/tmp/wolfxl-ooxml-fidelity-mutations-calamine-tests-fixed-20260508`: 416 passed, 52 passed with expected style/theme drift, 27 explicit invalid-source skips, 0 failures after ISO date-cell and UTF-8-BOM workbook-splice fixes | Adds non-SynthGL reader-corpus diversity for external-tool-authored files, macro/VBA, external links, charts, tables/validations, drawings/controls, and workbook-global state; still missing PowerPivot/data model and slicer/timeline buckets |
| BF30 remaining real-world downloads, `/Users/wolfgangschoenberger/Projects/SynthGL/data/real-world-ingestion/downloads/bf30_remaining` | 6 readable Microsoft Excel-authored workbooks; gap radar clear at `/tmp/wolfxl-gap-radar-bf30-remaining-fixed-20260508.json` with 5 unreadable source skips; bucket report pinned at `/tmp/wolfxl-corpus-buckets-bf30-remaining-20260509.json` | 99 default mutation rows at `/tmp/wolfxl-ooxml-fidelity-mutations-bf30-remaining-fixed-20260508`: 48 passed, 6 passed with expected style/theme drift, 45 explicit invalid-source skips, 0 failures after Mac Excel view-extension classification, native invalid-ZIP skip handling, and full-sheet hidden defined-name preservation | Adds public real-file diversity for external links, conditional-formatting extensions, drawing/media, embedded/control, table/validation, and workbook-global state; still missing external-tool-authored, macro/VBA, PowerPivot/data model, slicer/timeline, and chart buckets |
| Blind-holdout downloads, `/Users/wolfgangschoenberger/Projects/SynthGL/data/real-world-ingestion/downloads/blind_holdout` | 1 readable Microsoft Excel-authored workbook plus 5 explicit non-ZIP source skips; gap radar clear at `/tmp/wolfxl-gap-radar-blind-holdout-20260509.json`; bucket report pinned at `/tmp/wolfxl-corpus-buckets-blind-holdout-20260509.json` | 54 default mutation rows at `/tmp/wolfxl-ooxml-fidelity-mutations-blind-holdout-20260509`: 8 passed, 1 passed with expected style/theme drift, 45 explicit invalid-source skips, 0 failures | Adds a tiny held-out public-file pressure test for conditional-formatting extensions, table/validation, extension payloads, and workbook-global state. It remains side evidence only because most sources are blocked/non-ZIP and the corpus lacks chart, drawing/media, embedded/control, external-link, external-tool-authored, macro/VBA, PowerPivot/data model, and slicer/timeline buckets |
| Rescue downloads, `/Users/wolfgangschoenberger/Projects/SynthGL/data/real-world-ingestion/downloads/rescue` | 2 readable Microsoft Excel-authored budget workbooks plus 1 explicit non-ZIP source skip; gap radar clear at `/tmp/wolfxl-gap-radar-rescue-downloads-20260509.json`; bucket report pinned at `/tmp/wolfxl-corpus-buckets-rescue-downloads-20260509.json` | 27 default mutation rows at `/tmp/wolfxl-ooxml-fidelity-mutations-rescue-downloads-20260509`: 16 passed, 2 passed with expected style/theme drift, 9 explicit invalid-source skips, 0 failures | Adds held-out budget-workbook diversity for conditional-formatting extensions, drawing/media, embedded/control, table/validation, extension payloads, and workbook-global state. It remains side evidence because it lacks chart, external-link, external-tool-authored, macro/VBA, PowerPivot/data model, and slicer/timeline buckets |
| SEC/EDGAR downloads, `/Users/wolfgangschoenberger/Projects/SynthGL/data/real-world-ingestion/downloads/sec_edgar` | 6 readable Microsoft Excel-authored workbooks; gap radar clear at `/tmp/wolfxl-gap-radar-sec-edgar-20260509.json`; bucket report pinned at `/tmp/wolfxl-corpus-buckets-sec-edgar-20260509.json` | 54 default mutation rows at `/tmp/wolfxl-ooxml-fidelity-mutations-sec-edgar-20260509`: 48 passed, 6 passed with expected style/theme drift, 0 failures | Adds public-company/SEC workbook diversity for conditional-formatting extensions and workbook-global state, including Microsoft's FY2024 financial report plus five SEC aging workbooks. It remains side evidence because it lacks chart, drawing/media, embedded/control, external-link, external-tool-authored, macro/VBA, PowerPivot/data model, slicer/timeline, and table/validation buckets |
| SynthGL real-world ingestion tree, `/Users/wolfgangschoenberger/Projects/SynthGL/data/real-world-ingestion` | 28 readable Microsoft Excel-authored workbooks plus 17 explicit invalid-source skips; gap radar clear at `/tmp/wolfxl-gap-radar-synthgl-real-world-ingestion-20260509.json` after classifying Excel `[trash]/0000.dat` package debris | Quick mutation report `/tmp/wolfxl-ooxml-fidelity-mutations-synthgl-real-world-ingestion-quick-20260509.json`: 77 passed, 7 passed with expected style/theme drift, 51 skipped invalid source, 0 failures. Structural mutation report `/tmp/wolfxl-ooxml-fidelity-mutations-synthgl-real-world-ingestion-structural-20260509.json`: 21 passed, 119 passed with expected intentional/style drift, 85 skipped invalid source, 0 failures | Adds a wider SynthGL ingestion-side pressure test for CF extensions, external links, drawing/media, embedded/control, sheet metadata, table/validation, and workbook-global state. It still lacks external-tool-authored, macro/VBA, PowerPivot/data model, slicer/timeline, and chart buckets, so it remains side evidence rather than a completion gate |
| SynthGL CDS case-study tree, `/Users/wolfgangschoenberger/Projects/SynthGL/data/cds-case-study` | 16 readable Microsoft Excel-authored workbooks; gap radar clear at `/tmp/wolfxl-gap-radar-synthgl-cds-case-study-20260509.json` | Quick mutation report `/tmp/wolfxl-ooxml-fidelity-mutations-synthgl-cds-case-study-quick-20260509.json`: 46 passed, 2 passed with expected style/theme drift, 0 failures. Structural mutation report `/tmp/wolfxl-ooxml-fidelity-mutations-synthgl-cds-case-study-structural-20260509.json`: 13 passed, 67 passed with expected intentional/style drift, 0 failures | Adds a tracked case-study pressure test with one chart workbook, sheet metadata, table/validation, drawing/media, embedded/control, CF extensions, and workbook-global state. It still lacks external-link, external-tool-authored, macro/VBA, PowerPivot/data model, and slicer/timeline buckets, so it remains side evidence |
| WBD wDesk supplemental workbook, `/Users/wolfgangschoenberger/Projects/SynthGL/output/ui_runs/ingest_confidence/real_files/archive/dev/public-company-ir-supplemental-workbook/warner-bros-discovery` | 1 readable wDesk-authored workbook; gap radar clear at `/tmp/wolfxl-gap-radar-wdesk-wbd-directory-20260509.json`; bucket report pinned at `/tmp/wolfxl-corpus-buckets-wdesk-wbd-20260509.json` | 9 default mutation rows at `/tmp/wolfxl-ooxml-fidelity-mutations-wdesk-wbd-quoted-sheet-fixed-20260508`: 8 passed, 1 passed with expected style/theme drift, 0 failures after preserving source quotes in unchanged sheet-scoped print-area defined names | Adds external-tool-authored public-company workbook diversity for conditional-formatting extensions, drawing/media, embedded/control, sheet metadata, table/validation, workbook globals, and extension payloads; still missing macro/VBA, PowerPivot/data model, slicer/timeline, external-link, and chart buckets in this corpus |
| ILPA reporting-template workbooks, `/Users/wolfgangschoenberger/Projects/SynthGL/data/domain-ground-truth/11-lending-labor-banking/ilpa` | 4 readable Microsoft Excel-authored finance workbooks; gap radar clear at `/tmp/wolfxl-gap-radar-ilpa-20260508.json`; bucket report pinned at `/tmp/wolfxl-corpus-buckets-ilpa-20260509.json`; 3 of 4 contain external-link parts | 36 default mutation rows at `/tmp/wolfxl-ooxml-fidelity-mutations-ilpa-20260508`: 32 passed, 4 passed with expected style/theme drift, 0 failures; 24 richer structural/feature rows at `/tmp/wolfxl-ooxml-fidelity-mutations-ilpa-structural-20260508`: 8 passed, 16 passed with expected intentional drift, 0 failures for copy-first-sheet, rename-first-sheet, move-formula-range, add-data-validation, add-conditional-formatting, and retarget-external-links | Adds public finance-model diversity for external links, conditional-formatting extensions, drawing/media, embedded/control, workbook globals, and extension payloads; still missing external-tool-authored, macro/VBA, PowerPivot/data model, slicer/timeline, chart, and table/validation buckets in this corpus. Regenerate app-smoke before using app-open as current evidence for this side corpus |
| umya-spreadsheet result files, `/Users/wolfgangschoenberger/Projects/umya-spreadsheet/tests/result_files` | 78 readable OOXML workbooks plus 2 skipped by the bucket scanner; gap radar clear at `/tmp/wolfxl-gap-radar-umya-result-files-package-rel-fixed-20260508.json` after classifying embedded workbook `package` relationships as known embedded-object relationships | 240 conservative mutation rows at `/tmp/wolfxl-ooxml-fidelity-mutations-umya-result-files-auto-normalize-off-20260508`: 157 passed, 77 passed with expected style/theme drift, 6 skipped invalid source, 0 failures after preserving macro/comment/VML/shared-string sidecars by limiting automatic openpyxl VBA package normalization to explicit `keep_vba=True`, and after treating existing-sheet mutations as no-ops for zero-sheet workbooks such as `empty_worksheet.xlsx` | Adds macro/VBA, chart, conditional-formatting, drawing/comment/media, embedded/control, table/validation, workbook-global, and one external-link sidecar pressure test. This corpus is now clean for the conservative mutation slice, but still misses external-tool-authored, PowerPivot/data model, and slicer/timeline coverage |

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
   - Done: structural intentional-edit Excel-rendered smoke for
     `copy_first_sheet` and `move_formula_range` across the pinned
     22-workbook external-oracle pack.
   - Still needed: additional structural/adversarial intentional-edit
     Excel-rendered smoke, broader slicer/timeline UI variants, and broader
     real-file corpus diversity.
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
     A later Excel app-open/render probe found that newly created `<dxfs>`
     sections were inserted after `<colors>` and without the source namespace
     prefix in some workbooks; the CF writer now inserts `<dxfs>` before
     `tableStyles`/`colors`/`extLst` and prefixes generated CF/dxf XML to
     match prefixed source parts.
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
   - Latest shared-slicer breadth slice: added sidecar fixture
     `tests/fixtures/slicer_timeline_variants/real-excel-shared-slicer-two-pivots.xlsx`
     from MyExcelOnline's public free practice workbook for using one slicer
     across two pivot tables. The workbook reports `Application=Microsoft
     Excel`, contains two pivot tables on one pivot cache and a slicer cache
     connected to both pivots, and now has pinned mutation, Microsoft Excel
     source/intentional app-open, no-op/intentional Microsoft Excel render,
     destructive row/column deletion app-open and Microsoft Excel render,
     strict gap-radar, and interactive pivot/slicer presence evidence.
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
     differences are expected. The full 22-workbook external-oracle pack now
     passes Microsoft Excel render-smoke for marker-cell, copy-first-sheet,
     move-formula-range, add-data-validation, and add-conditional-formatting
     mutations with 110 combined results and 0 failures, using page sampling
     for large exports.
   - Latest broader render slice: the render comparison runner now supports
     recursive discovery, deterministic page sampling, and an explicit
     byte-identical-xlsx pass for no-op saves. A Microsoft Excel recursive
     no-op run over the live SynthGL fixture tree passes all 32 workbooks
     because every WolfXL no-op output is byte-for-byte identical to its source.
     An unguarded Excel-render attempt rendered 31 of 32 workbooks at RMSE 0.0
     before `number_memory/sigman_revenue_support.xlsx` hit Excel PDF-export
     automation fragility; a targeted package diff and Excel app-smoke showed
     that workbook is unchanged and still opens before and after no-op save.
   - Latest broader-corpus evidence slice: `scripts/run_ooxml_fidelity_mutations.py`
     now supports `--recursive` discovery for nested fixture trees. The live
     SynthGL fixture tree at
     `/Users/wolfgangschoenberger/Projects/SynthGL/tests/app/fixtures` has
     32 `.xlsx` workbooks and passes a combined no-op, marker-cell,
     style-cell, rename-first-sheet, and move-formula-range sweep with 160
     results and 0 failures. `style_cell` and `move_formula_range` report
     expected drift only for the intentionally introduced style/theme and
     formula-reference changes.
   - Latest adjacent issue-corpus mutation slice: the 61-workbook
     umya-spreadsheet issue corpus now passes a quick no-op, marker-cell, and
     add-conditional-formatting sweep with 183 results and 0 failures, plus a
     broader style-cell, copy-first-sheet, rename-first-sheet,
     move-formula-range, and add-data-validation sweep with 305 results and 0
     failures. The add-conditional-formatting mutation intentionally adds one
     `<dxf>` style record, so the runner now treats style/theme drift as
     expected only when `styles.xml` is otherwise identical after removing
     `<dxfs>` and the `<dxf>` count increases by exactly one.
   - Latest adjacent issue-corpus production bug found: `issue_178.xlsx`
     carries only an `x14:dataValidations` extension payload. The modify-mode
     data-validation patcher used to capture that extension payload as if it
     were the main worksheet `<dataValidations>` block, then reparent
     `x14:dataValidation` under an unprefixed wrapper and emit unbound `x14` /
     `xm` prefixes. The extractor now ignores prefixed extension
     `dataValidations` blocks, preserving the extension payload in place while
     adding a normal worksheet data-validation block.
   - Latest adjacent issue-corpus app-open slice: a representative
     source-valid umya subset (`issue_178.xlsx`, `pr_204.xlsx`,
     `wps_comment.xlsx`) passes Microsoft Excel app-open smoke for both source
     and add-data-validation mutation: 6 results, 0 failures. The full umya
     source-open triage finds 55 source-valid workbooks and 6 source-invalid or
     source-misdirected workbooks (`issue_216.xlsx`, `issue_217.xlsx`,
     `issue_219.xlsx`, `issue_220.xlsx`, `issue_222.xlsx`, `issue_225.xlsx`).
     The 55 source-valid workbooks then pass Microsoft Excel app-open after
     add-data-validation mutation with 55 results and 0 failures.
   - Latest adjacent issue-corpus render slice: the same 55 source-valid umya
     workbooks pass Microsoft Excel render-smoke after add-data-validation
     mutation with 55 results and 0 failures, sampling at most one page per
     workbook. This exposed and fixed a render-runner sampler bug where a
     one-page sample on a 3,000-page workbook was accidentally rasterizing the
     entire PDF.
   - Latest PowerPivot breadth slice: added a sidecar fixture
     `tests/fixtures/powerpivot_variants/real-excel-powerpivot-contoso-pnl.xlsx`
     from Microsoft's official ContosoPnL_Excel2013.zip sample for Profit and
     Loss Data Modeling and Analysis with Microsoft PowerPivot in Excel. It
     contains `xl/model/item.data`, workbook connections, four pivot cache
     definitions, three pivot tables, and slicer caches. The sidecar passes
     marker/copy-sheet/move-formula mutation sweep, records expected
     PowerView app-unsupported status for Microsoft Excel source and marker
     app-smoke with zero clean app-open passes, passes sampled Microsoft Excel
     render-smoke, and passes strict package gap radar with `power_view`
     recorded as an expected app-unsupported feature after classifying
     PowerPivot custom-property payloads, theme media relationships, and x15
     pivot/slicer extension URIs.
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
     external-oracle output with all six probe reports at
     `/tmp/wolfxl-interactive-evidence-external-oracle-all-20260508.json` is
     ready.
     `scripts/run_ooxml_interactive_probe.py` clears
     macro project presence for `real-excel-macro-basic.xlsm` by opening it in
     Microsoft Excel and verifying `xl/vbaProject.bin` remains present. It also
     clears embedded-control openability for `umya-control-props-basic.xlsx` and
     `real-excel-control-props-basic.xlsx` by opening them in Microsoft Excel
     and verifying embedded/control OOXML parts remain present. External-link
     prompt/cache behavior is clear for `synthetic-external-link-basic.xlsx`,
     `real-excel-timeline-slicer.xlsx`, and `real-excel-external-link-basic.xlsx`
     by opening them in Microsoft Excel, dismissing link update prompts, and
     verifying external-link OOXML parts remain present. Pivot refresh-state
     presence is clear for 7 pivot-bearing fixtures by opening them in
     Microsoft Excel and verifying pivot cache/table OOXML parts remain
     present. Slicer OOXML state-presence is clear for 4 slicer-bearing
     fixtures, and timeline OOXML state-presence is clear for
     `real-excel-timeline-slicer.xlsx`, by opening them in Microsoft Excel and
     verifying slicer/timeline OOXML state remains present. This is a current
     openability/state-presence gate, not full click-level selection
     automation.
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
| Pivot/slicer preservation across modify saves | `scripts/audit_ooxml_fidelity.py`, `scripts/audit_ooxml_fidelity_coverage.py`, `tests/fixtures/external_oracle/manifest.json`, `tests/fixtures/slicer_timeline_variants/manifest.json` | Coverage report `/tmp/wolfxl-coverage-all-evidence-workbook-globals-payload-final.json`: `pivot_slicer_preservation` clear, 8 fixtures, external-tool and real-Excel sources present. Sidecar shared-slicer variant adds a real Excel workbook with one slicer cache connected to two pivot tables, plus structural mutation, Excel app-open, no-op/intentional/destructive Excel render, strict gap-radar, and interactive pivot/slicer evidence | Covered more strongly for pinned and sidecar corpora; still not exhaustive for every interactive slicer/timeline state |
| Chart/style/color dependency preservation | Chart, chart-sheet, chart-style, style/theme fingerprints plus render reports | Coverage report: `chart_style_color_preservation` clear, 5 fixtures, no missing evidence. Engine-specific Excel render coverage `/tmp/wolfxl-coverage-excel-render-full-pack-with-rename-sheet-defined-name-20260508.json` is also clear | Covered for pinned corpus with full-pack Excel-rendered no-op, marker-cell and style-cell intentional edits with expected visual-delta evidence, copy-first-sheet structural intentional, rename-first-sheet structural intentional, move-formula-range structural intentional with expected visual-delta evidence, delete-first-row/column structural intentional with expected visual-delta evidence, add-data-validation, add-conditional-formatting render-equivalence, add-remove-chart render-equivalence, and retarget-external-links evidence; additional feature-specific intentional Excel rendering remains stronger future evidence |
| Conditional formatting extension preservation | CF/x14 semantic fingerprints and `dxfId` bounds | Coverage report: `conditional_formatting_extension_preservation` clear, 5 fixtures, no missing evidence | Covered for pinned corpus |
| External links and workbook relationship edge cases | Relationship graph audit, external-link target/cache/formula fingerprints, gap radar rel-type inventory, `retarget_external_links` mutation, Microsoft Excel app-open smoke, and Microsoft Excel render-equivalence for the retargeted outputs | Coverage report: `external_link_relationship_edges` clear, 3 fixtures. Gap radar clear for unknown rel types. Latest external-link retarget run: 22 fixtures, 0 failures; `synthetic-external-link-basic.xlsx`, `real-excel-timeline-slicer.xlsx`, and `real-excel-external-link-basic.xlsx` each show required external-link semantic drift plus the expected replaced external-link relationship, and named workbook formula tokens are rewritten when present. Latest matching Microsoft Excel app-open smoke: 22 fixtures, 0 failures, including those three retargeted external-link cases. Latest render-equivalence audit: `/tmp/wolfxl-retarget-external-links-render-equivalence-first-last-pages-20260509.json`, `ready=True`, 22 passed, 0 failures, 0 inconclusive | Covered more strongly for pinned target-retarget behavior; new relationship types need radar triage |
| Workbook connections / query metadata | Connection fingerprint and real/openpyxl connection fixtures | Coverage report: `workbook_connections_query_metadata` clear, 3 fixtures | Covered for pinned corpus |
| Python-in-Excel and sheet metadata | Python and sheet-metadata semantic fingerprints plus optional coverage surfaces | Recursive SynthGL coverage report: Python-in-Excel metadata clear on 1 real-Excel fixture; sheet metadata clear on 2 real-Excel fixtures, each with structural mutation evidence | Good emerging-surface gate when those parts appear; broader Python-in-Excel workbook variants still need acquisition |
| PowerPivot / workbook data model | Data-model binary/default/content relationship fingerprint plus targeted render/app-smoke reports | Current all-evidence report: `powerpivot_data_model_preservation` clear on 1 real-Excel fixture, with structural mutation, no-op render, intentional render, LibreOffice source/intentional app-open, and Microsoft Excel source/intentional app-open evidence. Sidecar Contoso PowerPivot variant adds a second real Microsoft sample with 3 mutation results, 2 expected PowerView app-unsupported Excel smoke results, 1 sampled Excel render result, strict gap-radar clearance with expected `power_view`, and sidecar coverage where PowerPivot is clear | Covered more strongly than before, but still narrow versus the full universe of data-model designs. The Contoso sidecar is not clean editable Excel app-open evidence on this Mac build because PowerView is blocked behind an unsupported-content prompt |
| OOXML extension payload preservation | Generic extension payload fingerprint plus extension-URI radar | Coverage report: `ooxml_extension_payload_preservation` clear, 16 fixtures. Gap radar clear for 0 unknown extension URIs | Covered for known extension URIs in pinned corpus |
| Tables / structured refs / validations | Table feature parts, structured-reference and validation fingerprints | Coverage report: `table_structured_refs_validations` clear, 11 fixtures | Covered for pinned corpus; richer table-filter/totals scenarios remain useful |
| Drawings / comments / embedded objects | Drawing-object fingerprint and VML structural fix | Coverage report: `drawings_comments_embedded_objects` clear, 15 fixtures. Drawing mutation sweep: 66 results, 0 failures | Covered for non-destructive object preservation; destructive geometry edits remain classified as expected drift, not unchanged-payload proof |
| Workbook global state | Defined names/protection/page setup plus durable package payload fingerprints | Coverage report: `workbook_global_state` clear, 21 fixtures. Workbook-global sweep: 88 results, 0 failures | Covered for durable payloads in pinned corpus; calc-chain is intentionally treated as volatile |
| Hidden future surface discovery | `scripts/audit_ooxml_gap_radar.py` | Latest radar: 22 fixtures, clear true, 0 unknown part families, rel types, content types, extension URIs | Good tripwire for new package/extension surface, not a semantic proof for every known-looking XML pattern |
| Broader recursive gap radar | `scripts/audit_ooxml_gap_radar.py /Users/wolfgangschoenberger/Projects/SynthGL/tests/app/fixtures --recursive --json --strict` | Latest run: 32 nested workbooks, clear true, 0 unknown part families, rel types, content types, extension URIs. First recursive run exposed Python-in-Excel and sheet-metadata surfaces that are now classified | Good live-corpus tripwire; still not a replacement for richer external workbook acquisition |
| Umya issue-corpus gap radar | `scripts/audit_ooxml_gap_radar.py /Users/wolfgangschoenberger/Projects/umya-spreadsheet/tests/test_files --json --strict` | Latest run: 61 workbooks, clear true, 0 unknown part families, rel types, content types, extension URIs after classifying sensitivity-label metadata, JavaScript API project payloads, WMF thumbnails, non-numbered theme parts, x14 data-validation extensions, drawing hidden-effect extensions, and chart pivot-option extensions | Useful adjacent issue corpus; now paired with quick mutation side evidence, but still not a full evidence gate |
| Corpus diversity | `scripts/audit_ooxml_corpus_buckets.py tests/fixtures/external_oracle --json --strict` | Latest run: `ready=True`, 22 workbooks, no missing buckets | Good coverage-shape check; must be paired with behavioral gates |
| Umya issue-corpus diversity, mutation smoke, and source-valid app/render smoke | `scripts/audit_ooxml_corpus_buckets.py /Users/wolfgangschoenberger/Projects/umya-spreadsheet/tests/test_files --json`; `scripts/run_ooxml_fidelity_mutations.py ... --mutation no_op --mutation marker_cell --mutation add_conditional_formatting`; `scripts/run_ooxml_fidelity_mutations.py ... --mutation style_cell --mutation copy_first_sheet --mutation rename_first_sheet --mutation move_formula_range --mutation add_data_validation`; `scripts/run_ooxml_app_smoke.py ... --app excel --mutation source`; `scripts/run_ooxml_app_smoke.py /tmp/wolfxl-umya-excel-source-valid-55 --app excel --mutation add_data_validation`; `scripts/run_ooxml_render_compare.py /tmp/wolfxl-umya-excel-source-valid-55 --render-engine excel --max-pages-per-fixture 1 --mutation add_data_validation`; `scripts/audit_ooxml_fidelity_coverage.py ... --report /tmp/wolfxl-ooxml-fidelity-mutations-umya-quick.json --report /tmp/wolfxl-ooxml-fidelity-mutations-umya-structural.json --app-report /tmp/wolfxl-app-smoke-excel-umya-source-triage.json --app-report /tmp/wolfxl-app-smoke-excel-umya-source-valid-55-add-dv.json --render-report /tmp/wolfxl-render-excel-umya-source-valid-55-add-dv.json` | Latest bucket run: 63 workbooks with 53 Excel-authored, 9 external-tool-authored, 2 macro/VBA, 2 external-link, 6 chart, 21 drawing/comment/media, 21 embedded/control, and 2 sheet-metadata examples; missing PowerPivot and slicer/timeline buckets. Latest quick mutation run: 183 results, 0 failures. Latest structural mutation run: 305 results, 0 failures after fixing the `x14:dataValidations` reparenting bug found by `issue_178.xlsx`. Latest full source triage: 61 source workbooks, 55 open cleanly in Microsoft Excel, 6 fail source-open. Latest source-valid intentional Microsoft Excel app-open smoke: 55 add-data-validation mutation results, 0 failures. Latest source-valid intentional Microsoft Excel render smoke: 55 add-data-validation mutation results, 0 failures. Combined coverage side report: 61 fixtures, 2 mutation reports, 2 app reports, 1 render report, structural mutation, source/intentional app-open, and intentional render evidence attached to present chart, CF, external-link, table/validation, workbook-global, sheet-metadata, style/theme, and extension-payload surfaces; intentionally `ready=false` as a full gate | Adds breadth, mutation confidence, broad real-Excel openability, and bounded renderability over source-valid workbooks in an adjacent issue corpus; still needs PowerPivot, slicer/timeline examples, and stronger external-tool examples for some feature groups before it can behave like a complete corpus |
| Calamine issue-corpus exploratory sweep | `scripts/audit_ooxml_corpus_buckets.py /Users/wolfgangschoenberger/Projects/calamine/tests --recursive --json`; `scripts/audit_ooxml_gap_radar.py /Users/wolfgangschoenberger/Projects/calamine/tests --recursive --json`; `scripts/run_ooxml_fidelity_mutations.py /Users/wolfgangschoenberger/Projects/calamine/tests --recursive --output-dir /tmp/wolfxl-ooxml-fidelity-mutations-calamine-quick-skip-invalid-20260508 --mutation no_op --mutation marker_cell --skip-invalid-source`; `scripts/run_ooxml_fidelity_mutations.py /Users/wolfgangschoenberger/Projects/calamine/tests --recursive --output-dir /tmp/wolfxl-ooxml-fidelity-mutations-calamine-structural-skip-invalid-20260508 --mutation style_cell --mutation copy_first_sheet --mutation rename_first_sheet --mutation move_formula_range --mutation add_data_validation --skip-invalid-source`; `scripts/run_ooxml_fidelity_mutations.py /Users/wolfgangschoenberger/Projects/calamine/tests --recursive --output-dir /tmp/wolfxl-ooxml-fidelity-mutations-calamine-tests-fixed-20260508 --skip-invalid-source` | Latest bucket run audits 52 readable workbooks and explicitly skips 3 invalid or unsupported inputs: one bad CRC workbook, one backslash-path unsafe package, and one encrypted/non-ZIP workbook. Bucket coverage is intentionally not ready because PowerPivot and slicer/timeline are absent. Latest gap radar audits 52 readable workbooks, skips the same 3 invalid inputs, and is clear for unknown part families, relationship types, content types, and extension URIs after classifying macOS `.DS_Store` package noise and case-variant `xl/SharedStrings.xml` as known package surface. The quick mutation sweep now produces 110 results, 0 failures, 104 passed mutations, and 6 explicit `skipped_source_invalid` results for the invalid/encrypted/unsafe source inputs. The richer structural/feature sweep now produces 275 results, 0 failures, 102 passed mutations, 158 expected-drift passes, and 15 explicit `skipped_source_invalid` results for the same invalid/encrypted/unsafe source inputs, after closing the real reader gaps for OPC case-variant part names (`xl/SharedStrings.xml`) and namespace-prefix variants on relationship IDs (`relationships:id`). The pinned default mutation sweep produces 495 rows, 0 failures, 416 passed mutations, 52 expected style/theme drift passes, and 27 explicit invalid-source skips | Useful broader-corpus pressure test and now pinned in the evidence bundle. It exposed audit-tool robustness gaps plus two real reader fidelity gaps, now fixed so exploratory corpora produce machine-readable skipped/unknown reports and valid case-variant or namespace-prefix workbooks load correctly. It is a clean mutation pressure test over source-valid files, but not a complete gate because the corpus still includes intentionally malformed fixtures and lacks key BI/slicer buckets |
| Broader recursive coverage audit | `scripts/audit_ooxml_fidelity_coverage.py /Users/wolfgangschoenberger/Projects/SynthGL/tests/app/fixtures --recursive --report /tmp/wolfxl-ooxml-fidelity-mutations-synthgl-python-metadata-final/report.json --report /tmp/wolfxl-ooxml-fidelity-mutations-synthgl-richer-fixed-20260508.json` plus `scripts/audit_ooxml_corpus_buckets.py ... --recursive --json` | Latest coverage run `/tmp/wolfxl-coverage-synthgl-recursive-richer-20260508.json`: 32 fixtures, recursive true, 2 mutation reports, all source classes inferred as real Excel, richer structural/feature mutation evidence attached to present live-corpus surfaces, Python metadata clear on 1 fixture, sheet metadata clear on 2 fixtures. Latest corpus-bucket run: 32 Excel-authored workbooks, missing external-link, external-tool-authored, macro/VBA, PowerPivot, and slicer/timeline buckets | Good check that broad live fixtures are not silently excluded; proves the SynthGL tree is useful side evidence but not a replacement for the pinned external-oracle pack |
| Excel-rendered output evidence | `scripts/run_ooxml_render_compare.py ... --render-engine excel` plus `scripts/audit_ooxml_fidelity_coverage.py ... --require-render --require-render-engine excel` | Latest raw Excel-render reports: `/tmp/wolfxl-render-excel-full-pack.json` has 22 no-op results and 0 failures; `/tmp/wolfxl-render-excel-intentional-marker-full-pack.json` has 22 marker-cell mutation results and 0 failures; `/tmp/wolfxl-marker-cell-render-delta-20260509.json` adds before/after Excel visual-delta evidence for marker-cell with `ready=True`, 19 changed, 3 sampled pages unchanged, 0 failures, and 0 inconclusive; `/tmp/wolfxl-render-excel-intentional-style-cell-full-pack-20260509.json` has 22 style-cell mutation results and 0 failures; `/tmp/wolfxl-style-cell-render-delta-20260509.json` adds before/after Excel visual-delta evidence for style-cell with `ready=True`, 19 changed, 3 sampled pages unchanged, 0 failures, and 0 inconclusive; `/tmp/wolfxl-render-excel-intentional-insert-tail-row-full-pack-20260509.json` has 22 insert-tail-row mutation results and 0 failures; `/tmp/wolfxl-insert-tail-row-render-delta-20260509.json` adds before/after Excel visual-delta evidence for insert-tail-row with `ready=True`, 21 changed, 1 sampled page unchanged, 0 failures, and 0 inconclusive; `/tmp/wolfxl-render-excel-intentional-insert-tail-col-full-pack-20260509.json` has 22 insert-tail-column mutation results and 0 failures; `/tmp/wolfxl-insert-tail-col-render-delta-20260509.json` adds before/after Excel visual-delta evidence for insert-tail-column with `ready=True`, 19 changed, 3 sampled pages unchanged, 0 failures, and 0 inconclusive; `/tmp/wolfxl-render-excel-intentional-delete-marker-tail-row-full-pack-20260509.json` has 22 delete-marker-tail-row mutation results and 0 failures; `/tmp/wolfxl-delete-marker-tail-row-render-delta-20260509.json` adds before/after Excel visual-delta evidence for delete-marker-tail-row with `ready=True`, 1 changed, 21 sampled pages unchanged, 0 failures, and 0 inconclusive; `/tmp/wolfxl-render-excel-intentional-copy-sheet-full-pack.json` has 22 copy-first-sheet structural mutation results and 0 failures; `/tmp/wolfxl-render-excel-intentional-rename-first-sheet-defined-name-fixed-20260508.json` has 22 rename-first-sheet structural mutation results and 0 failures; `/tmp/wolfxl-rename-sheet-render-equivalence-first-last-pages-20260509.json` adds before/after Excel pixel equivalence for rename-first-sheet with `ready=True`, 22 passed, 0 failures, and 0 inconclusive; `/tmp/wolfxl-render-excel-intentional-move-formula-range-full-pack.json` has 22 move-formula-range structural mutation results and 0 failures; `/tmp/wolfxl-move-formula-range-render-delta-20260509.json` adds before/after Excel pixel delta evidence for move-formula-range with `ready=True`, 19 changed, 3 sampled pages unchanged, 0 failures, and 0 inconclusive; `/tmp/wolfxl-render-excel-intentional-delete-first-axis-20260508.json` has 44 delete-first-row/column structural mutation results and 0 failures; `/tmp/wolfxl-delete-first-row-render-delta-20260509.json` and `/tmp/wolfxl-delete-first-col-render-delta-20260509.json` add before/after Excel pixel delta evidence for first row/column deletion with `ready=True`, 21 changed, 1 sampled page unchanged, 0 failures, and 0 inconclusive each; `/tmp/wolfxl-render-excel-intentional-add-data-validation-full-pack.json` has 22 add-data-validation feature-add results and 0 failures; `/tmp/wolfxl-render-excel-intentional-add-conditional-formatting-full-pack-fixed.json` has 22 add-conditional-formatting feature-add results and 0 failures; `/tmp/wolfxl-add-conditional-formatting-render-equivalence-20260509.json` adds before/after Excel pixel equivalence for add-conditional-formatting with `ready=True`, 22 passed, 0 failures, and 0 inconclusive; `/tmp/wolfxl-render-excel-intentional-add-remove-chart-20260508.json` has 22 add-remove-chart feature-add/remove results and 0 failures; `/tmp/wolfxl-add-remove-chart-render-equivalence-20260509.json` adds before/after Excel pixel equivalence for add-remove-chart with `ready=True`, 22 passed, 0 failures, and 0 inconclusive; `/tmp/wolfxl-render-excel-retarget-external-links-first-last-pages-20260509.json` has 22 external-link retarget results and 0 failures; `/tmp/wolfxl-retarget-external-links-render-equivalence-first-last-pages-20260509.json` adds before/after Excel pixel equivalence for external-link retarget with `ready=True`, 22 passed, 0 failures, and 0 inconclusive. Latest coverage report `/tmp/wolfxl-coverage-excel-render-full-pack-with-rename-sheet-defined-name-20260508.json`: `ready=True`, 22 fixtures audited, 10 Excel render reports, 0 unclear surfaces. The producer stages files in Excel's sandbox container to avoid macOS `Grant File Access` prompts | Covered as full-pack no-op plus marker-cell, style-cell, insert-tail-row/column, and delete-marker-tail-row intentional edits with expected visual-delta evidence, copy-first-sheet structural intentional, rename-first-sheet structural intentional and render-equivalent, move-formula-range structural intentional with expected visual-delta evidence, delete-first-row/column structural intentional with expected visual-delta evidence, add-data-validation, add-conditional-formatting render-equivalence, add-remove-chart render-equivalence, and retarget-external-links Excel-rendered and render-equivalent evidence across the current surface matrix; still not exhaustive feature-specific intentional-edit visual equivalence |
| Interactive behavior evidence | `scripts/audit_ooxml_interactive_evidence.py tests/fixtures/external_oracle --recursive --strict --report /tmp/wolfxl-interactive-probes-pivot-one-clean-20260508/interactive-probe-report.json --report /tmp/wolfxl-interactive-probes-slicer-one-clean-20260508/interactive-probe-report.json --report /tmp/wolfxl-interactive-probes-timeline-one-clean-20260508/interactive-probe-report.json --report /tmp/wolfxl-interactive-probes-external-link-one-clean-20260508/interactive-probe-report.json --report /tmp/wolfxl-interactive-probes-embedded-one-clean-20260508/interactive-probe-report.json --report /tmp/wolfxl-interactive-probes-macro-clean-rerun-20260508/interactive-probe-report.json`; `scripts/audit_ooxml_interactive_evidence.py tests/fixtures/external_oracle --probe-kind excel_ui_interaction ...`; plus `scripts/run_ooxml_interactive_probe.py --probe-kind excel_ui_interaction` for targeted UI action reports | Latest state-presence run `/tmp/wolfxl-interactive-evidence-external-oracle-all-20260508.json`: `ready=True`; all six current `ooxml_state_presence` probe classes clear with `incomplete_report_count=0`. Latest aggregated UI-interaction run `/tmp/wolfxl-ui-interaction-evidence-external-oracle-20260509.json`: `ready=True`, 22 fixtures, 10 reports, all six `excel_ui_interaction` probe classes clear. Latest sidecar UI-interaction run `/tmp/wolfxl-ui-interaction-evidence-shared-slicer-cache-sidecar-20260509.json`: `ready=True`, scoped to `slicer_selection_state` on the shared-slicer/two-pivots workbook. Latest UI-interaction reports: macro `Disable Macros` button click clear at `/tmp/wolfxl-ui-interaction-macro-20260508/interactive-probe-report.json`; pivot `refresh all` command clear at `/tmp/wolfxl-ui-interaction-pivot-20260508/interactive-probe-report.json`; embedded list-box click persistence clear at `/tmp/wolfxl-ui-interaction-control-click-20260508/interactive-probe-report.json` with `List Box 1` clicked, workbook saved, and `ctrlProp1.xml` rewritten from `sel=0` to `sel=2`; adjacent issue-corpus list-box click persistence clear at `/tmp/wolfxl-ui-interaction-umya-listbox-control-20260508/interactive-probe-report.json` with `List Box 1` clicked, workbook saved, and control state changed; stateless button-control click persistence clear at `/tmp/wolfxl-ui-interaction-button-control-20260508/interactive-probe-report.json` with `Button 1` clicked, workbook saved, and control parts preserved; table-slicer all-item-click persistence clear at `/tmp/wolfxl-ui-interaction-slicer-all-items-20260508/interactive-probe-report.json` with `REGION` selected, both `REGION` and `YEAR` clicked, workbook saved, and `table1.xml` rewritten with `EAST` and `2014` filters; pivot-chart slicer click persistence clear at `/tmp/wolfxl-ui-interaction-pivot-chart-slicer-extra-20260508/interactive-probe-report.json` with `REGION` selected, both `REGION` and `YEAR` clicked, workbook saved, and slicer state changed; external-tool pivot-slicer click persistence clear at `/tmp/wolfxl-ui-interaction-excelize-pivot-slicer-20260508/interactive-probe-report.json` with `Month` selected, both `Month` and `Region` clicked, workbook saved, and slicer state changed; shared pivot-slicer cache persistence clear at `/tmp/wolfxl-ui-interaction-shared-slicer-cache-20260508/interactive-probe-report.json` with `SALES MONTH` selected, one item clicked, workbook saved, and the slicer-cache item state rewritten; timeline month-click persistence clear at `/tmp/wolfxl-ui-interaction-timeline-click-20260508/interactive-probe-report.json` with `ORDER DATE` selected, May clicked, and `timelineCache1.xml` rewritten from Q1 2012 to May 2012; forced external-link prompt handling clear at `/tmp/wolfxl-ui-interaction-external-link-review-fix-20260508/interactive-probe-report.json` with Excel's update-link prompt temporarily forced, `Don't Update` clicked, and the prior setting restored; current-setting external-link prompt handling clear at `/tmp/wolfxl-ui-interaction-external-link-current-prompt-20260508/interactive-probe-report.json` with the existing prompt setting left unchanged and `Don't Update` clicked; adjacent issue-corpus current-setting external-link prompt handling clear at `/tmp/wolfxl-ui-interaction-umya-external-link-current-prompt-20260508/interactive-probe-report.json` with `issue_297.xlsx`, existing prompt setting left unchanged, and `Don't Update` clicked | Covered for current Excel openability/state-presence probes plus thirteen targeted UI-action paths. Two embedded list-box clicks, one stateless button-control click, table-slicer item clicks, Microsoft-authored pivot-chart slicer item clicks, external-tool-authored pivot-slicer item clicks, one shared pivot-slicer cache click, one timeline date-range interaction, and forced plus current-setting external-link prompt paths across two fixtures now have real-Excel proof. Broader control/slicer/timeline variants remain open |
| Whole-pack preservation under common edits | `tests/test_external_oracle_preservation.py` | `198 passed` after the latest changes | Strong pinned-pack regression gate |
| Combined all-evidence gate | `scripts/audit_ooxml_fidelity_coverage.py --strict --require-render --require-intentional-render --require-app --require-intentional-app` | Latest regenerated report `/tmp/wolfxl-coverage-all-evidence-current-code-plus-excel-powerpivot.json`: `ready=True`, 22 fixtures, 13 surfaces, 5 mutation reports, 6 render reports, 9 app reports; external-link relationship edges now accept `retarget_external_links` as a structural mutation and clear on three retargeted external-link fixtures with intentional Microsoft Excel app-open evidence | Strong current-state gate |
| Evidence bundle freshness | `scripts/audit_ooxml_evidence_bundle.py Plans/ooxml-current-evidence-bundle.json --strict` | Latest run `/tmp/wolfxl-current-evidence-bundle-audit-delete-marker-tail-row-delta-20260509.json`: `ready=True`, 138 report artifacts verified, 138 producer commands recorded, 0 issues | Stronger provenance over the current generated evidence set; still dependent on generated reports being refreshed when fixtures or gates change |
| Completion claim guard | `scripts/audit_ooxml_completion_claim.py Plans/ooxml-current-evidence-bundle.json --strict-current-evidence` | Latest run `/tmp/wolfxl-completion-claim-audit-delete-marker-tail-row-delta-20260509.json`: `current_supported_claim_ready=True`, `exhaustive_claim_ready=False`, `missing_requirement_count=4`, and current bundle ready with 138 pinned reports | Prevents the supported current claim from being silently upgraded into the exhaustive "no real-world Excel fidelity gaps" claim |

Current conclusion:

- The repo can honestly claim: **no known fidelity gap in the currently pinned
  and classified real-world OOXML surface.**
- The repo should not claim: **no real-world Excel fidelity gaps exist.** That
  would require broader corpus diversity, additional feature-specific
  structural intentional-edit Excel-rendered smoke, broader click-level
  slicer/date-range variants, and more adversarial mutations than
  the current pack can provide.
- The machine-readable completion claim guard now encodes that boundary:
  `scripts/audit_ooxml_completion_claim.py --strict-current-evidence` passes
  for the current supported claim, while `--strict-claim` intentionally fails
  until the open requirements above are closed.

Next evidence slices before declaring a higher-confidence "no known gaps":

1. Keep widening the external workbook corpus sweep beyond the current
   15-report, 236-readable-workbook portfolio. The portfolio now covers all
   required provenance/feature buckets, including macro/VBA, PowerPivot,
   slicer/timeline, embedded-object/control, external-link, external-tool, and
   chart/table examples, but it is still curated evidence rather than a random
   or customer-scale workbook sample.
2. Extend structural intentional mutation Excel render-smoke beyond the current
   `copy_first_sheet`, `rename_first_sheet`, `move_formula_range`,
   `delete_first_row`, `delete_first_col`, `add_remove_chart`, and
   `retarget_external_links` slices into higher-risk feature-specific edits
   such as pivot/slicer structural edits and other external-link
   relationship-preserving edits.
3. Continue expanding click-level interactive UI automation beyond the current
   `ooxml_state_presence` probes. Macro prompt handling, pivot refresh,
   two embedded list-box clicks, one stateless button-control click,
   table-slicer item clicks, one shared
   pivot-slicer cache click, Microsoft-authored pivot-chart slicer item clicks,
   external-tool-authored pivot-slicer item clicks, one timeline date-range
   click, and external-link prompt handling under both a prompt-forcing setup
   and the current Excel prompt setting now have pinned UI-interaction
   artifacts; broader embedded-control/slicer/timeline variants remain open.
4. Keep the gap radar strict: every newly seen part family, relationship type,
   content type, or extension URI must become either an allowlisted known
   surface with a semantic fingerprint or an explicit gap.
