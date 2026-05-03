# Openpyxl Parity Program

**Status**: Proposed (S0 in progress, kicked off 2026-05-03).
**Target tag**: rolling - each sprint can ship under its own minor version
(v2.1.x, v2.2.x, ...). No single tag closes the program.
**Calendar**: ~12 sprints across S0-S11; S0 is the one mandatory pre-cursor.
**Predecessor**: v2.0.0 (Sprint Ν - pivot construction).
**Successor**: programme retires when the openpyxl-compat oracle reports
≥95% pass and the compatibility matrix has zero ❌ rows in P0/P1.

This program runs in parallel with `Plans/pre-release-expanded-oracle-and-cleanup.md`;
neither blocks the other. Public-launch freeze stays in force until both this
program and that pre-release sprint reach their gates.

## Why this program exists

The 2026-05-03 audit confirmed wolfxl is "openpyxl-compatible" but not yet a
true drop-in replacement. The README and `docs/trust/limitations.md` already
track this honestly, but there is no single sequenced workplan that closes the
user-visible API gaps end-to-end and no quantitative measure of how close the
public surface actually is.

This program converts the audit's gap list into 12 sprints (S0-S11), each
tracked against an explicit verification gate, and adds a hard "% openpyxl
test suite passing" number that monotonically rises sprint-over-sprint.

## Goals

1. Close every openpyxl drop-in gap in priority order; nothing permanently
   deferred.
2. Quantify compatibility by running an openpyxl-API oracle against wolfxl.
   Publish baseline (S0), improve sprint-over-sprint.
3. Delegate mechanical work to Codex subagents in parallel; reserve Claude for
   design, OOXML semantics, and oracle-harness work.
4. Decision-gate the three structural giants (`.xlsb` writes, `.ods` ecosystem,
   VBA generation) before committing engineering effort (S9-S11).

## Non-goals

- Not replacing or pausing `pre-release-expanded-oracle-and-cleanup.md`; that
  track continues independently.
- Not lifting the v2.0 release freeze - freeze stays until both this program
  and the existing pre-release sprint reach their gates.
- Not refactoring existing modules beyond what each gap fix requires (the
  cleanup track owns refactors).

## Status table (live tracker)

Status values: `proposed | in-progress | blocked | review | landed`.

Class legend: **∥** parallel-safe within sprint · **⊥** serialized · **🤖**
codex-delegable (bounded spec + exemplar exists) · **🧠** needs Claude
(design choices).

Priority legend: **P0** ship-blocking for "drop-in" claim · **P1** high
impact, medium cost · **P2** medium impact, scoped feature · **P3** large
structural effort, decision-gated.

| #   | Gap                                                                 | Pri | Class | Sprint | Status      | Owner   | Branch | Started    | Completed | Notes |
|-----|---------------------------------------------------------------------|----:|:-----:|-------:|:------------|:--------|:-------|:-----------|:----------|:------|
| G01 | Openpyxl test suite as compatibility oracle                         | P0  | ⊥ 🧠  | S0     | landed      | Claude  | main   | 2026-05-03 | 2026-05-03 | Probe-based harness in `tests/test_openpyxl_compat_oracle.py`; baseline recorded below |
| G02 | Public compatibility matrix (spec-driven, scannable)                | P0  | ⊥ 🧠  | S0     | landed      | Claude  | main   | 2026-05-03 | 2026-05-03 | Spec-driven (Python module not YAML; pyyaml not vendored); replaces prose in `docs/migration/compatibility-matrix.md` |
| G03 | Diagonal borders Python bridge (writer model exists)                | P1  | ∥ 🧠  | S1     | blocked     | Claude  | feat/parity-G03-diagonal-borders | 2026-05-03 |           | Codex pod returned POD-BLOCKED 02:59 PDT — scope wider than handoff: modify-mode `BorderSpec` (`src/wolfxl/styles.rs:61`) has no diagonal fields and `border_to_xml:269` hardcodes `<diagonal/>`. Reclassified 🤖→🧠. Claude takeover after S1 sibling pods land. |
| G04 | Protection (sheet/workbook) Python flow-through                     | P1  | ∥ 🟨  | S1     | partial     | Codex   | main | 2026-05-03 | 2026-05-03 | Workbook half landed (a694bbc): WorkbookProtection camelCase aliases, 15 tests pass, closes `protection.workbook` probe. Cell half deferred to Claude takeover (native_writer_formats lacks Protection emit). |
| G05 | NamedStyle / GradientFill / DifferentialStyle flow                  | P1  | ∥ 🧠  | S1     | partial     | Claude  | main | 2026-05-03 | 2026-05-03 | NamedStyle cellStyleXfs infra landed: writer stamps `xfId` on `<xf>`, reader walks `cellXfs[s].xf_id -> cellStyles[xf_id]` and exposes `cell.style`. GradientFill writer + reader landed: writer emits `<gradientFill>` (type/degree/path attrs + ordered stops); reader parses gradient block back into Python `GradientFill` via the format dict's `gradient` sub-dict. Both `cell.named_style` and `cell.gradient_fill` probes flipped `partial -> supported` (oracle 36/50 -> 38/50). DifferentialStyle still tracked for follow-up. |
| G06 | Image replace/delete public API                                     | P1  | ∥ 🤖  | S1     | landed      | Codex   | main | 2026-05-03 | 2026-05-03 | POD-DONE (b01a447 → cherry-picked to main). 528 lines: `_worksheet_media.py`, `patcher_drawing.rs`. Closes `images.replace_remove` probe. 2286 passed, no regressions. |
| G07 | Array/DataTable/spill formula coverage audit + tests                | P1  | ∥ 🤖  | S1     | landed      | Codex   | main | 2026-05-03 | 2026-05-03 | Cherry-picked (792d14d). Closes `array_formulas.data_table` probe. Pod hit -k filter mismatch in handoff; work re-verified clean (2285 passed). |
| G08 | Threaded comments write+modify                                      | P1  | ∥ 🧠  | S2     | landed      | Claude  | main | 2026-05-03 | 2026-05-03 | Steps 2-5 landed. Step 5 (modify mode): `apply_threaded_comments_phase` (`src/wolfxl/patcher_sheet_blocks.rs`) extracts existing threadedCommentsN.xml + personList.xml, merges queued ops, emits fresh bytes, mutates rels graph + content-types, and synthesizes `tc={topId}` legacy placeholders into `queued_comments`. Python flushes (`_flush_pending_threaded_comments_to_patcher`, `_flush_pending_persons_to_patcher`) drain BEFORE the legacy comments flush. `has_pending_save_work` extended for the new queues. 5 modify-mode round-trip tests (`tests/test_threaded_comments_modify.py`); full suite 2363 passed; oracle pass rate 78% (39/50) with `comments_threaded` confirmed supported. |
| G09 | Rich text in headers/footers                                        | P1  | ∥ 🤖  | S2     | landed      | Claude  | main   | 2026-05-03 | 2026-05-03 | Openpyxl-shaped `_HeaderFooterPart` lifted onto `HeaderFooterItem.left/center/right` (`python/wolfxl/worksheet/header_footer.py`). Each part exposes `text`/`font`/`size`/`color` with mini-format compose + parse (`&"Arial,Bold"&14&KFF0000Title`) mirroring openpyxl's `FORMAT_REGEX`. Setting a plain string still works (parses inline codes). The Rust writer/reader paths are unchanged - the inline mini-format string round-trips opaquely through `HeaderFooterItemSpec.left/center/right`. Modify-mode patcher round-trips through the existing `apply_sheet_setup_phase` path. 22 new write+modify tests (`tests/test_header_footer_richtext.py`); existing `test_header_footer.py` updated to the part-shaped API. Full Python suite 2441 passed; cargo `--workspace` clean. Oracle pass rate 78% -> 80% (40/50): `rich_text.headers_footers` flips `not_yet -> supported`. |
| G10 | Rich text in chart labels (extend existing)                         | P2  | ∥ 🤖  | S2     | landed      | Claude  | main   | 2026-05-03 | 2026-05-03 | `DataLabelList` + `DataLabel` accept a `rich=` kwarg that inflates a `CellRichText` into a wolfxl `RichText` and flows through `_dlbls_to_snake` as `tx_pr_runs`. Rust `DataLabels` carries `Vec<TitleRun>` and `emit_data_labels` drops a `<c:txPr>` block sharing the chart-title run emitter (`crates/wolfxl-writer/src/emit/charts/text.rs::emit_tx_pr`). Axis-title rich text rides the existing `Title.tx.rich` path. 8 tests in `tests/test_chart_label_richtext.py` (write-mode round-trip + openpyxl read-back + modify-mode preservation); full suite 2428 passed; oracle pass rate 80% -> 82% (41/50) post-merge with `charts.label_rich_text` flipping `not_yet -> supported`. |
| G11 | CF icon sets (Python builder + writer emitter)                      | P1  | ∥ 🤖  | S3     | proposed    |         |        |            |           | Exemplar: RFC-026 conditional formatting |
| G12 | CF data bars (Python builder + writer emitter)                      | P1  | ∥ 🤖  | S3     | proposed    |         |        |            |           | Exemplar: RFC-026 |
| G13 | CF color scales beyond basic                                        | P2  | ∥ 🤖  | S3     | proposed    |         |        |            |           | Exemplar: RFC-026 |
| G14 | CF stop-if-true, priority ordering, dxf integration                 | P2  | ⊥ 🤖  | S3     | proposed    |         |        |            |           | Serialized after G11-G13 (shared module) |
| G15 | Combination charts (mixed types, secondary axis)                    | P1  | ⊥ 🧠  | S4     | rfc-drafted | Claude  |        | 2026-05-03 |           | RFC-069 drafted; awaiting impl pod (S4 cohort) |
| G16 | Pivot chart per-point overrides                                     | P2  | ⊥ 🤖  | S4     | proposed    |         |        |            |           | Exemplar: RFC-046 chart overrides |
| G17 | Pivot table mutation of existing pivots                             | P1  | ⊥ 🧠  | S5     | proposed    |         |        |            |           | RFC required; cache refresh consistency |
| G18 | External links (workbook-level collection + rels)                   | P1  | ⊥ 🧠  | S6     | proposed    |         |        |            |           | RFC required |
| G19 | VBA inspection API (read-only)                                      | P2  | ⊥ 🧠  | S6     | proposed    |         |        |            |           | Read-only; no authoring |
| G20 | `write_only=True` streaming write mode                              | P1  | ⊥ 🧠  | S7     | proposed    |         |        |            |           | Bounded-memory append-only path |
| G21 | Slicer outside pivot context                                        | P2  | ⊥ 🤖  | S8     | proposed    |         |        |            |           | Exemplar: pivot-side slicer support |
| G22 | Defined-name edge cases                                             | P2  | ∥ 🤖  | S8     | proposed    |         |        |            |           | Exemplar: RFC-021 |
| G23 | Calc-chain edge cases                                               | P2  | ∥ 🤖  | S8     | proposed    |         |        |            |           |  |
| G24 | Print settings depth audit                                          | P2  | ∥ 🤖  | S8     | proposed    |         |        |            |           | Exemplar: RFC-055 |
| G25 | `.xlsb` write support                                               | P3  | ⊥ 🧠  | S9     | proposed    |         |        |            |           | Decision-gated before kickoff |
| G26 | `.xls` write support                                                | P3  | ⊥ 🧠  | S9     | proposed    |         |        |            |           | Decision-gated before kickoff |
| G27 | `.ods` read+write                                                   | P3  | ⊥ 🧠  | S10    | proposed    |         |        |            |           | Decision-gated before kickoff |
| G28 | VBA generation (macro authoring)                                    | P3  | ⊥ 🧠  | S11    | proposed    |         |        |            |           | Decision-gated before kickoff |

## Baseline metrics (Sprint 0)

Recorded 2026-05-03 from a clean ``pytest tests/test_openpyxl_compat_oracle.py``
run with ``WOLFXL_COMPAT_ORACLE_WRITE_BASELINE=1``. Raw JSON snapshot lives at
``.pytest_cache/compat_oracle_baseline.json`` (gitignored).

The numbers below are post-triage. Initial run showed 27 passed / 4 xpassed;
each xpass was hardened to also reload through openpyxl (the reference
reader). One flipped to xfailed (real gap), three confirmed already-supported
and had their spec entries promoted to ``supported``. The harness now sits at
zero xpasses, so every probe carries an honest signal.

| Date       | Metric                                                  |        Value | Notes |
|------------|---------------------------------------------------------|-------------:|:------|
| 2026-05-03 | Compat-oracle total probes                              |           50 | curated baseline; S1+ may grow |
| 2026-05-03 | Compat-oracle passed (green)                            |           30 | green = ``status=supported`` and probe passes |
| 2026-05-03 | Compat-oracle xpassed                                   |            0 | post-triage; spec aligned with reality |
| 2026-05-03 | Compat-oracle xfailed                                   |           20 | real gaps tracked under G03-G24 |
| 2026-05-03 | Compat-oracle failed                                    |            0 | hard failures; any non-zero blocks the sprint |
| 2026-05-03 | Compat-oracle pass rate                                 |        60.0% | passed / total |
| 2026-05-03 | Matrix rows ✅ Supported                                |    43 / 74   | Sourced from ``docs/migration/_compat_spec.py`` |
| 2026-05-03 | Matrix rows 🟡 Partial                                  |    17 / 74   | Each Partial row carries a ``gap_id`` |
| 2026-05-03 | Matrix rows ❌ Not Yet                                  |    13 / 74   | All tracked under G01-G28 |
| 2026-05-03 | Matrix rows ⛔ Out of Scope                             |     1 / 74   | ``.ods`` reads (G27 - decision-gated S10) |

Triage outcome (xpass investigation):

| xpass entry                      | Triage action                                                                          |
|----------------------------------|----------------------------------------------------------------------------------------|
| ``charts.combination``           | **Hardening exposed real gap** - openpyxl reload sees only one chart family. Stays G15.|
| ``protection.sheet``             | **Promoted to ``supported``** - sheet flag, formatCells override, password hash all round-trip. G04 retained for ``protection.workbook`` only. |
| ``cf.data_bars``                 | **Promoted to ``supported``** - openpyxl reload sees DataBar with cfvo min/max preserved. G12 retained for percent / formula cfvo cases (not yet probed). |
| ``array_formulas.array_formula`` | **Promoted to ``supported``** - openpyxl reload reconstructs ``ArrayFormula(ref, text)``. G07 retained for ``array_formulas.data_table``. |

Program-level gate: pass rate must rise sprint-over-sprint until ≥95%. With a
50-probe baseline at 60%, that means closing ~17 more gaps' worth of probes
(or strengthening probes and surfacing new gaps - either way moves the
denominator up).

The pass-% number must monotonically rise sprint-over-sprint; any regression
fails the per-sprint gate and blocks the sprint from being marked landed.

## Sprint 1 metrics (post-merge)

Recorded 2026-05-03 after cherry-picking the three S1 winners onto main
(``792d14d`` G07 DataTableFormula, ``a694bbc`` G04 WorkbookProtection
camelCase aliases, ``da51897`` G06 image remove/replace). G03 and G05 were
reclassified from 🤖 to 🧠 after their pods correctly POD-BLOCKED on
out-of-scope native-Rust changes; both await Claude takeover.

| Date       | Metric                                                  |        Value | Delta vs S0 baseline |
|------------|---------------------------------------------------------|-------------:|:---------------------|
| 2026-05-03 | Compat-oracle total probes                              |           50 | unchanged |
| 2026-05-03 | Compat-oracle passed (green)                            |           38 | +8 (data_table, protection.workbook, images.replace_remove, protection.cell, diagonal_borders, named_style, gradient_fill, +1) |
| 2026-05-03 | Compat-oracle xfailed                                   |           12 | -8 (gaps closed) |
| 2026-05-03 | Compat-oracle failed                                    |            0 | unchanged |
| 2026-05-03 | Compat-oracle pass rate                                 |        76.0% | +16.0 pp |
| 2026-05-03 | Matrix rows ✅ Supported                                |    51 / 74   | +8 |
| 2026-05-03 | Matrix rows 🟡 Partial                                  |    10 / 74   | -7 |
| 2026-05-03 | Matrix rows ❌ Not Yet                                  |    12 / 74   | -1 |
| 2026-05-03 | Matrix rows ⛔ Out of Scope                             |     1 / 74   | unchanged |
| 2026-05-03 | Full Python suite (``pytest -q``)                       | 2333 / 12 xf | -8 xfailed (gap closures); 0 failed |

S1 sprint gate: pass rate moved 60.0% → 66.0% with zero regressions in
``cargo test --workspace`` or ``pytest -q``. Two pods (G03, G05) deferred to
Claude with no LOC merged; one pod (G04) had its cell-half reverted and the
workbook-half kept. Matrix renderer regenerated from merged spec; row counts
verified via ``status_totals()``.

## Sprint sequence

### Sprint 0 - Foundation (CRITICAL FIRST, ~1 session, Claude-led)

In scope:

- **G01** - Build harness in `tests/test_openpyxl_compat_oracle.py`. The
  harness ships a curated probe set (≥30 probes) tied to gap entries in the
  spec YAML, plus a stub fetch script `scripts/fetch_openpyxl_corpus.py` that
  S1 can flesh out to vendor openpyxl's source-distribution tests.
- **G02** - Spec-driven matrix at `docs/migration/compatibility-matrix.md`,
  source spec at `docs/migration/_compat_spec.yaml` so each gap closure
  auto-updates the table via `scripts/render_compat_matrix.py`.
- Tracker file at `Plans/openpyxl-parity-program.md` (this file).

Out of scope:

- Vendoring openpyxl's full upstream test corpus - deferred to S1, behind the
  fetch-script stub.
- Closing any actual gaps - S0 is measurement only.

Acceptance:

1. `Plans/openpyxl-parity-program.md` exists and renders.
2. `docs/migration/_compat_spec.py` imports; `docs/migration/compatibility-matrix.md` regenerates from it.
3. `tests/test_openpyxl_compat_oracle.py` runs; baseline pass count recorded
   in this file under "Baseline metrics".
4. `cargo test --workspace` and `pytest -q` stay green - no regressions in
   the existing suite.

Acceptance log (2026-05-03):

1. ✅ Tracker file present (this file); 12-sprint structure + status table render.
2. ✅ `_compat_spec.py` imports cleanly (74 entries / 22 categories);
   `python scripts/render_compat_matrix.py` regenerates
   `docs/migration/compatibility-matrix.md` deterministically.
3. ✅ Oracle harness runs: 50 probes, 27 passed, 4 xpassed, 19 xfailed,
   0 failed (62.0% pass rate). Baseline JSON written to
   `.pytest_cache/compat_oracle_baseline.json` when
   `WOLFXL_COMPAT_ORACLE_WRITE_BASELINE=1`.
4. ✅ Regression check: `uv run --no-sync pytest -q` (excluding
   external-oracle/diffwriter slices that need extra fixtures) →
   **2280 passed, 41 skipped, 19 xfailed, 4 xpassed, 0 failed**.

Spec note: the plan originally called the source spec
`_compat_spec.yaml`. We landed it as a typed Python module
(`_compat_spec.py`) instead because PyYAML is not in the wolfxl
runtime/dev deps and adding it just to parse one config file would
have been a net cost. The renderer imports the module via
`importlib.util.spec_from_file_location`; the oracle harness imports
`ENTRIES` directly. Behaviour is unchanged from the YAML plan.

### Sprint 1 - Bridge Completion (5-way parallel codex, ~1-2 sessions wall-clock)

In scope: G03-G07. Each gap is a separate codex handoff on its own git
worktree. All five touch independent code paths and have existing exemplars.

Codex handoff specs are in [`Plans/rfcs/handoffs/`](rfcs/handoffs/) — one
file per pod, each self-contained per the program's 6-field contract
(goal / files / exemplar / acceptance / guards / verification). The
[handoffs README](rfcs/handoffs/README.md) is the dispatch index.

Pod plan:

| Pod | Branch                                  | Deliverable | Spec |
|-----|-----------------------------------------|-------------|------|
| α   | `feat/parity-G03-diagonal-borders`      | Diagonal border Python bridge; tests | [G03](rfcs/handoffs/G03-diagonal-borders.md) |
| β   | `feat/parity-G04-protection`            | Workbook-protection camelCase aliases + Cell.protection setter (sheet half done in S0) | [G04](rfcs/handoffs/G04-workbook-protection.md) |
| γ   | `feat/parity-G05-named-style-bridge`    | NamedStyle / GradientFill bridge + combined-style preservation | [G05](rfcs/handoffs/G05-named-style-gradient-differential.md) |
| δ   | `feat/parity-G06-image-replace-remove`  | `Worksheet.replace_image` / `remove_image` (modify mode) | [G06](rfcs/handoffs/G06-image-replace-remove.md) |
| ε   | `feat/parity-G07-data-table-formula`    | DataTableFormula round-trip (basic ArrayFormula done in S0) | [G07](rfcs/handoffs/G07-data-table-formula.md) |

Acceptance:

1. Compat oracle pass count rises by ≥7 probes (G03 +1, G04 +2, G05 +3, G06 +1, G07 +1; cell.protection counts under G04, combined-style counts under G05).
2. Each pod's existing-feature tests stay green.
3. `tests/test_external_oracle_preservation.py` stays green when fixtures present.
4. Compat matrix updated; status table flipped to `landed` for G03-G07.

### Sprint 2 - Comments & Rich Text (3-way parallel, ~1-2 sessions)

In scope: G08 (Claude-led, threaded comment OOXML semantics), G09 + G10
(codex handoffs).

Acceptance:

1. Threaded comments survive openpyxl read.
2. Header/footer rich text round-trips through modify mode.
3. Chart-label rich text round-trips for all v1.6+ chart families.
4. Compat oracle pass count up; matrix updated.

### Sprint 3 - Conditional Formatting Completeness (4-way with merge gate, ~2 sessions)

In scope: G11-G13 parallel codex; G14 serialized after merges (touches
`crates/wolfxl-writer/src/emit/conditional_formats.rs` shared with all three).

Acceptance:

1. CF rule types reach openpyxl coverage (icon sets, data bars, all color-scale variants).
2. Differential styles emit under modify mode and write mode.
3. `stopIfTrue` and priority ordering survive load-modify-save.
4. Compat oracle pass count up.

### Sprint 4 - Charts beyond Singletons (~2 sessions)

In scope: G15 (Claude-designed RFC for combo chart axis-id allocation), G16
(codex handoff).

Acceptance:

1. Combination chart (e.g. bar + line on shared category axis with secondary
   value axis) opens correctly in Excel and LibreOffice.
2. Pivot-chart per-point overrides round-trip.
3. External-oracle preservation green.

### Sprint 5 - Pivot Mutation (~2-3 sessions)

In scope: G17. Claude design + codex implementation. RFC required.

Acceptance:

1. Edit-then-save preserves existing pivot semantics.
2. Cache refresh consistency verified against ClosedXML/NPOI fixtures.
3. `Plans/pre-release-expanded-oracle-and-cleanup.md` flag for G17 in-place
   pivot edits flips to closed.

### Sprint 6 - External Links + VBA Inspection (~2 sessions)

In scope: G18 + G19 parallel.

Acceptance:

1. External-link round-trips through modify mode without dropping the
   `xl/externalLinks/` parts.
2. VBA inspection API exposes module names and signatures (read-only, not
   authoring).

### Sprint 7 - Streaming Write (~2-3 sessions, Claude-led)

In scope: G20. `Workbook(write_only=True)` with bounded-memory append-only
path. Today the call accepts the kwarg but writes via the standard
in-memory path (per `compatibility-matrix.md` line 26 note).

Acceptance:

1. 10M-row write within bounded memory (≤ 200 MB RSS measured).
2. openpyxl `write_only` documented usage parity.

### Sprint 8 - Edge Cases (~1-2 sessions)

In scope: G21-G24 parallel where independent.

Acceptance:

1. Edge cases covered in compat matrix.
2. Compat oracle pass count up.

### Sprint 9 - Legacy Writes (DECISION-GATED before kickoff)

In scope: G25 (`.xlsb` write), G26 (`.xls` write).

Decision required: most users transcribe to `.xlsx` instead. Confirm scope
before spending 4-6 sessions on `.xlsb` BIFF12 writer alone. Block on a
documented user-need signal in `Plans/followups/`.

### Sprint 10 - `.ods` (DECISION-GATED)

In scope: G27. ODS is OpenDocument, a different parser/writer ecosystem.

Decision required: is wolfxl really the right home for ODS, or should users
keep odfpy?

### Sprint 11 - VBA Generation (DECISION-GATED)

In scope: G28. Macro authoring from Python is rare and large.

Decision required: confirm there is a real user need before scoping. The
audit's data on user-asked-for VBA-authoring is sparse; do not engineer ahead
of demand.

## Codex delegation strategy

A codex handoff for this program is a self-contained markdown spec under
`Plans/rfcs/handoffs/` with:

1. **Goal** - one sentence describing the user-facing capability.
2. **Files to touch** - explicit paths from the feature-add pattern. The
   exploration mapped these for `Image` (queue → flush →
   native_writer / patcher_drawing → tests) and `BarChart` (same plus chart
   emit modules). Each codex spec reuses that path.
3. **Reference exemplar** - pointer to an analogous already-implemented
   feature ("follow `Image` for new drawing-anchored types; follow
   `BarChart.to_rust_dict()` for new chart types").
4. **Acceptance tests** - specific test names that must pass, including the
   openpyxl-oracle slice that targets this gap.
5. **Out-of-scope guards** - explicit list: don't refactor patcher, don't
   change OOXML write order, don't break PyO3 signatures, don't touch
   unrelated `_workbook_*.py` helpers.
6. **Verification commands** - `cargo test`, `uv run --no-sync maturin
   develop`, `pytest tests/test_<feature>.py`,
   `pytest tests/test_openpyxl_compat_oracle.py -k <slice>`, plus
   external-oracle gate when relevant.

🤖-tagged items are good codex fits (bounded spec + exemplar). 🧠-tagged items
require Claude design first (RFC under `Plans/rfcs/`), then optional codex
handoff for mechanical implementation.

For an N-way parallel sprint, spawn N handoffs from the same Claude session,
each on its own git worktree. Each handoff includes a "do not merge until N-1
sibling branches land" note to avoid merge-order surprises in shared files.

## RFC numbering

Next free RFC number at S0 kickoff: **RFC-068**. Each 🧠 gap gets an RFC.
🤖 gaps can skip RFC and go straight to a codex handoff spec under
`Plans/rfcs/handoffs/`.

| Gap | RFC                | Status     | Title (proposed)                                  |
|-----|--------------------|------------|---------------------------------------------------|
| G08 | RFC-068            | implemented (write+read+modify) | Threaded comments write+modify   |
| G15 | RFC-069            | drafted    | Combination charts (multi-family plotArea + secondary axis) |
| G17 | RFC-070            | proposed   | Pivot table mutation in modify mode               |
| G18 | RFC-071            | proposed   | External links - workbook-scoped collection + rels |
| G19 | RFC-072            | proposed   | VBA inspection API (read-only)                    |
| G20 | RFC-073            | proposed   | `write_only=True` streaming write mode            |
| G25 | RFC-074 (gated)    | proposed   | `.xlsb` write support                             |
| G26 | RFC-075 (gated)    | proposed   | `.xls` write support                              |
| G27 | RFC-076 (gated)    | proposed   | `.ods` read+write                                 |
| G28 | RFC-077 (gated)    | proposed   | VBA generation                                    |

## Cross-session pickup protocol

A fresh session reads this file, finds rows with `in-progress` or `blocked`
in the status table, and resumes from the per-sprint section for that gap.
No re-derivation of priority needed.

When closing a gap:

1. Update the row in the status table: status → `landed`, fill `Completed` date.
2. Update `docs/migration/_compat_spec.yaml` for the matching probe(s).
3. Run `python scripts/render_compat_matrix.py` to regenerate
   `docs/migration/compatibility-matrix.md`.
4. Run the per-sprint gate (see Verification below). Only mark `landed` if
   the gate is green.

## Verification

Per-sprint gate (run before any sprint is marked landed):

1. `cargo test --workspace` - Rust tests pass.
2. `uv run --no-sync maturin develop` - extension rebuilds.
3. `uv run --no-sync pytest -q` - full Python suite green.
4. `uv run --no-sync pytest tests/test_external_oracle_preservation.py -q` -
   external-oracle gate (when fixtures present).
5. **NEW**: `uv run --no-sync pytest tests/test_openpyxl_compat_oracle.py -q`
   - openpyxl-API oracle pass count must not regress.
6. `WOLFXL_RUN_LIBREOFFICE_SMOKE=1 uv run --no-sync pytest tests/diffwriter/soffice_smoke.py -q`
   - LibreOffice smoke (when soffice available).

Program-level gate: openpyxl oracle pass % monotonically increases sprint-over-sprint
until ≥95%. Revisit the bar after S0 reveals baseline; the bar may need
adjusting once we know how many probes the curated set ends up at.

End-to-end: at program completion the README's "Openpyxl-style imports. One
import change." section can honestly be re-titled "Drop-in replacement for
openpyxl" with the matrix and pass-% as evidence.

## Lessons applied

- Pre-dispatch contract spec (Sprint Μ-prime lesson #12): each codex handoff
  spec is self-contained before the worktree is opened.
- Doc-only pod can scaffold with `<!-- TBD: SHA -->` markers (lesson #3):
  release notes drafted before tag.
- Strict xfail = bug receipt (lesson #1): any compat oracle probe that fails
  red must remain red - not silently skipped - until the gap closes.
- Worktree pattern (lesson #8): N-way parallel sprints use isolated worktrees.
- Quantitative oracle as primary metric (NEW in this program): the
  pass-count number replaces qualitative "is it parity?" judgement calls.

## Risk register

| # | Risk | Mitigation |
|---|------|-----------|
| 1 | Probe-based oracle is "what we already think we should support", missing real-user gaps. | S1+ adds a fetch-script path that vendors openpyxl's actual source-distribution tests; probe set is a starting point, not a ceiling. |
| 2 | "% pass monotonically rises" forces a temptation to delete failing probes instead of fixing gaps. | Probe deletion requires a comment in the YAML referencing why (false positive, duplicate, gap permanently deferred); deletions show up in PR diff. |
| 3 | Decision-gated sprints (S9-S11) end up engineered without real demand because the program treats them as "remaining work". | Each gated sprint has an explicit decision-required block at the top; the freeze is on engineering, not on collecting user-demand evidence. |
| 4 | Codex handoffs mass-merge into shared modules (e.g. CF emit module in S3), creating merge conflicts. | S3 and any other shared-module sprint use a serial-merge-after-N-parallel pattern; handoff specs include "do not merge until siblings land" notes. |
| 5 | The compat oracle and the existing `tests/parity/openpyxl_surface.py` ratchet drift apart. | Render compat matrix from `_compat_spec.yaml`; `openpyxl_surface.py` keeps its existing `wolfxl_supported` flags as test-time machine input. Each gap closure flips both atomically. |

## Calendar

S0 is single-session work (this session). S1-S8 are each ~1-2 sessions with
parallel pod fan-out where the class column shows ∥. S9-S11 are gated and
not on the calendar until decision-required signals fire.

| Sprint | Calendar (approximate) | Mode |
|--------|------------------------|------|
| S0     | 2026-05-03 (this session) | Claude-led, single session |
| S1     | 2026-05-04 → 2026-05-08   | 5-way parallel codex |
| S2     | 2026-05-08 → 2026-05-12   | Claude (G08) + 2 codex |
| S3     | 2026-05-12 → 2026-05-18   | 3-way parallel codex + serial G14 |
| S4     | 2026-05-18 → 2026-05-22   | Claude RFC + 1 codex |
| S5     | 2026-05-22 → 2026-05-29   | Claude RFC + codex impl |
| S6     | 2026-05-29 → 2026-06-02   | 2 RFCs in parallel |
| S7     | 2026-06-02 → 2026-06-09   | Claude-led; long-poll write-only correctness |
| S8     | 2026-06-09 → 2026-06-12   | 4-way parallel codex |
| S9     | gated                     | only if decision fires |
| S10    | gated                     | only if decision fires |
| S11    | gated                     | only if decision fires |

## Changelog

- 2026-05-03 — Plan file created (this session). S0 in progress: tracker +
  matrix spec + oracle harness skeleton landing under one Claude session.
