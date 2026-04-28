# RFC Index — openpyxl-Replacement Gap Closure

> **Source plan**: kept in conversation transcript; see `Plans/launch-posts.md` for the marketing artifacts that depend on this work landing.
> **Scope**: Modify-mode parity (T1.5) + structural worksheet ops + W5 rip-out. Excludes Phase 2-5 read gaps and construction-side stubs (chart construction, NamedStyle, etc.).
> **Goal**: Defensible "full openpyxl replacement" claim by end of Phase 4.

This file is regenerated from each RFC's frontmatter. Edit a source RFC, not this index.

---

## Status Table (23 RFCs)

| ID | Title | Status | Phase | Estimate | Depends-on | Unblocks |
|---|---|---|---|---|---|---|
| 001 | W5 — rust_xlsxwriter rip-out | Shipped | 2 | S | — | (clean baseline) |
| 010 | Infra — rels graph parser/serializer | Shipped | 2 | M | 001 | 022, 023, 024, 035 |
| 011 | Infra — XML-block-merger primitive | Shipped | 2 | M | 001 | 022, 023, 024, 025, 026 |
| 012 | Infra — formula reference translator | Shipped | 2 | L | 001 | 021, 030, 031, 034, 035 |
| 013 | Infra — patcher extensions (ZIP-add, ancillary parts, two-phase flush) | Shipped | 2 | M | 001 | 022, 023, 024, 035 |
| 020 | T1.5 — Document properties | Shipped | 3 | S | 013 | — |
| 021 | T1.5 — Defined names mutation | Shipped | 3 | M | 011, 012 | 030, 031, 034, 035, 036 |
| 022 | T1.5 — Hyperlinks | Shipped | 3 | M | 010, 011, 013 | 030, 031, 035 |
| 023 | T1.5 — Comments + VML drawings | Shipped | 3 | L | 010, 011, 013 | 030, 031, 035 |
| 024 | T1.5 — Tables | Shipped | 3 | M | 010, 011, 013 | 030, 031, 035 |
| 025 | T1.5 — Data validations | Shipped | 3 | M | 011 | 030, 031, 035 |
| 026 | T1.5 — Conditional formatting | Shipped | 3 | M | 011 | 030, 031, 035 |
| 030 | Structural — insert_rows / delete_rows | Shipped | 4 | L | 012, 021, 022, 024, 025, 026 | 034 |
| 031 | Structural — insert_cols / delete_cols | Shipped | 4 | L | 012, 021, 022, 024, 025, 026 | 034 |
| 034 | Structural — move_range | Shipped | 4 | L | 012, 030, 031 | — |
| 035 | Structural — copy_worksheet | Shipped | 4 | XL | 010, 012, 013, 021, 022, 023, 024, 025, 026 | — |
| 036 | Structural — move_sheet | Shipped | 4 | M | 021 | — |
| 040 | Read-side parity — rich text (`Cell.rich_text`) | Shipped | 5 (1.3) | M | — | (T2 rich-text writes, post-1.3) |
| 041 | Read-side parity — streaming reads (`read_only=True`) | Shipped | 5 (1.3) | L | — | (large-fixture LRBench ingest) |
| 042 | Read-side parity — password-protected reads (`password=`) | Shipped | 5 (1.3) | M | (msoffcrypto-tool optional dep) | (post-1.3 encrypted writes) |
| 043 | Read-side parity — `.xlsb` / `.xls` reads (runtime-dispatched calamine backends) | Shipped | 5 (1.4) | L | calamine-styles (xlsx); upstream calamine (xlsb / xls) | (closes Phase 5 — KNOWN_GAPS empty post-1.4) |
| 044 | Encryption — write-side OOXML encryption (`Workbook.save(password=...)`) | Shipped | 5 (1.5) | M | 042 (msoffcrypto-tool optional dep), 013 | (T3 closure for encrypted xlsx writes) |
| 045 | Image construction — `wolfxl.drawing.image.Image` (replace stub) | Shipped | 5 (1.5) | L | 010, 013 | (T3 closure for image construction; chart-construction prerequisites) |
| 046 | Chart construction — `wolfxl.chart.*` (replace `_make_stub`) | Shipped | 5 (1.6 + 1.6.1 + 1.7) | XL | 010, 013, 035, 045 | (v2.0.0 pivot charts) |
| 050 | `Worksheet.remove_chart` / `replace_chart` + RichText title (Sprint Ξ) | Shipped | 5 (1.7) | M | 046 | — |
| 051 | docs/migration overhaul for v1.7 (Sprint Ξ) | Shipped | 5 (1.7) | M | 046 | — |
| 052 | docs/performance refresh for v1.7 (Sprint Ξ) | Shipped | 5 (1.7) | M | — | — |
| 053 | Public launch posts + materialise `Plans/launch-posts.md` (Sprint Ξ) | Shipped | 5 (1.7) | S | — | — |
| 047 | Pivot caches — `wolfxl.pivot.PivotCache` + cache definition + records emit (Sprint Ν) | Approved | 5 (2.0) | XL | 010, 013, 035, 046 | 048, 049, 054 |
| 048 | Pivot tables — `wolfxl.pivot.PivotTable` + layout + RFC-035 deep-clone (Sprint Ν) | Approved | 5 (2.0) | XL | 010, 013, 035, 046, 047 | 049, 054 |
| 049 | Pivot-chart linkage — `chart.pivot_source = pt` (Sprint Ν) | Approved | 5 (2.0) | M | 046, 048 | 054 |
| 054 | v2.0.0 launch hardening + docs + README rewrite (Sprint Ν) | Approved | 5 (2.0) | M | 047, 048, 049 | (PyPI publish + launch) |
| 055 | Print / view / sheet protection (Sprint Ο Pod 1A) | Approved | 5 (2.0) | L | 010, 011, 013, 035 | 060 |
| 056 | AutoFilter conditions + eval engine (Sprint Ο Pod 1B) | Approved | 5 (2.0) | XL | 011, 013, 026 | 060 |
| 057 | Array / DataTable formulas (Sprint Ο Pod 1C) | Approved | 5 (2.0) | L | 012, 013 | 060 |
| 058 | Workbook-level security (Sprint Ο Pod 1D) | Approved | 5 (2.0) | M | 044, 011 | 060 |
| 059 | Public exceptions + IndexedList (Sprint Ο Pod 1E) | Approved | 5 (2.0) | S | — | 060 |
| 060 | openpyxl-shaped class re-export shims (Sprint Ο Pod 2) | Approved | 5 (2.0) | M | 055, 056, 057, 058, 059 | (v2.0 launch) |
| 061 | Advanced pivot construction — slicers, calc fields, calc items, GroupItems, styling (Sprint Ο Pod 3) | Approved | 5 (2.0) | XL | 047, 048, 049, 026 | (v2.0 launch) |
| 062 | Page breaks + dimensions (Sprint Π Pod Π-α) | Approved | 5 (2.0) | M | 011, 013, 035 | (v2.0 launch) |
| 063 | Merge + tables + worksheet copier (Sprint Π Pod Π-β) | Approved | 5 (2.0) | M | 024, 035 | (v2.0 launch) |
| 064 | Styles — NamedStyle + GradientFill + Protection + DifferentialStyle + Fill (Sprint Π Pod Π-γ) | Approved | 5 (2.0) | XL | 026, 061 | (v2.0 launch) |
| 065 | Workbook properties + internals (Sprint Π Pod Π-δ) | Approved | 5 (2.0) | M | 058 | (v2.0 launch) |
| 066 | Re-routes + RFC-060 doc cleanup (Sprint Π Pod Π-ε) | Approved | 5 (2.0) | XS | 055, 060 | (v2.0 launch) |

Estimate buckets: S = ≤2 days, M = 3-5 days, L = 1-2 weeks, XL = 2+ weeks (calendar, with parallel subagent dispatch + review overhead).

**Estimate roll-up** (calendar days, sequenced by dependency, Claude Code with parallel subagents):
- Phase 2: 1 S (½d) + 4 M (≈3d each) + 1 L (≈10d) ≈ **2 weeks** (W5 first, then 010/011/012/013 in three parallel pods)
- Phase 3: 1 S + 5 M + 1 L ≈ **2-3 weeks** (parallel pods, bottleneck is review)
- Phase 4: 3 L + 1 M + 1 XL ≈ **3-4 weeks** (some sequencing — 030/031 → 034; 035 is critical path)
- Hardening + docs: **1 week**
- **Total**: 6-7 weeks. Hard cuts at end of Phase 3 (1.0 release) and end of Phase 4 (1.1).

---

## Dependency DAG

```
                    001 W5 rip-out  (independent, ship first)
                         │
                         ▼ (clean baseline — no rust_xlsxwriter)
                    ┌────────────────────────────────────────┐
                    │ Phase 2 — Foundation (parallel pod)    │
                    │  010 rels graph                        │
                    │  011 xml-block-merger                  │
                    │  012 formula xlator                    │
                    │  013 patcher infra extensions (NEW)    │
                    └────────────────────────────────────────┘
                         │
            ┌────────────┴───────────────┬────────────────┐
            ▼                            ▼                ▼
       020 properties              021 defined names    022 hyperlinks (010+011+013)
       (no infra deps)             (011+012)            023 comments   (010+011+013)
                                   │                    024 tables     (010+011+013)
                                   │                    025 data validations (011)
                                   │                    026 cond. formatting (011)
                                   │                            │
                                   └────────────────────────────┤
                                                                ▼
                                          ┌─────────────────────────────────┐
                                          │ Phase 4 — Structural ops        │
                                          │                                 │
                                          │  030 insert/delete rows         │
                                          │  031 insert/delete cols         │
                                          │  (both: 012+021+022+024+025+026)│
                                          │                                 │
                                          │  034 move_range  (012+030+031)  │
                                          │  035 copy_worksheet             │
                                          │      (010+012+013+021+022+023   │
                                          │       +024+025+026)             │
                                          │  036 move_sheet  (021 only)     │
                                          └─────────────────────────────────┘
                                                          │
                                          ┌───────────────┴─────────────────┐
                                          │ Phase 5 — Read + construction   │
                                          │ parity (1.3 → 1.6)              │
                                          │                                 │
                                          │  040 rich-text round-trip       │
                                          │  041 streaming reads            │
                                          │  042 password-protected reads   │
                                          │  043 .xlsb / .xls reads         │
                                          │  044 encryption writes  (042)   │
                                          │  045 image construction (010,   │
                                          │      013)                       │
                                          │  046 chart construction         │
                                          │      (010, 013, 035, 045)       │
                                          └─────────────────────────────────┘
```

**Critical path** (longest dep chain): 001 → 012 → 021 → 030 → 034. Anything blocking 012 (formula translator) blocks five downstream RFCs; it is the single most important infra item.

---

## Wave 1 + Wave 2 Findings (cross-cutting)

Three findings emerged from the research that shape the corpus but live outside any single RFC:

### Finding 1 — RFC-013 was added to absorb cross-cutting patcher gaps

The original plan listed 15 RFCs. Both pod-P1 (RFC-010) and pod-P6 (RFC-035) independently identified the same cross-cutting plumbing gaps:
- ZIP rewriter is ADD-blind (`src/wolfxl/mod.rs:343-356` only patches existing entries; can't add new comments / tables / sheet parts)
- Patcher reads no ancillary parts at `open()` (tables, comments, CF, DV are read via the `CalamineStyledBook` reader, not loaded into mutable patcher state)
- No two-phase flush for cross-sheet aggregation (`[Content_Types].xml` overrides + `xl/styles.xml` `<dxfs>` from multiple sheets)

These are bundled as **RFC-013 — Patcher Infrastructure Extensions** (Phase 2, M, ~5.5 days). Without RFC-013, four downstream RFCs (022, 023, 024, 035) would re-litigate the same plumbing.

### Finding 2 — P0 fix-it: structural ops raise `AttributeError`, not `NotImplementedError`

Pod-P5 confirmed: `ws.insert_rows(2)` today raises `AttributeError: 'Worksheet' object has no attribute 'insert_rows'` rather than `NotImplementedError` pointing at RFC-030. This is worse than a stubbed feature — users hit it at runtime instead of construction.

**Recommended P0 commit before Phase 2 kickoff** (~30 LOC, no research needed): add `insert_rows` / `delete_rows` / `insert_cols` / `delete_cols` / `move_range` / `copy_worksheet` / `move_sheet` stubs that raise `NotImplementedError("structural op X is RFC-NNN, scheduled for 1.1 / Phase 4")`. Removes the worst surprise; signals roadmap.

### Finding 3 — RFC-020 dropped its RFC-011 dependency (correct call)

Pod-P2 re-evaluated and dropped RFC-011 from RFC-020. `docProps/core.xml` is a 600-byte standalone file; full rewrite is correct; XML-block-merger is over-engineering for this case. The judgment call is preserved in RFC-020 §4.2.

---

## Open Questions Requiring Decision Before Phase 2 Kickoff

Each is flagged with the RFC that surfaced it. Listed by reverse impact. Decisions locked 2026-04-25 ahead of Phase-2 RFC-010 kickoff; question #3 is the only one still pending and gates RFC-012 only.

| # | Question | Surfaced by | Decision |
|---|---|---|---|
| 1 | **Crate boundary**: extract `crates/wolfxl-rels/` and `crates/wolfxl-formula/` as new workspace crates? Adds 2 crates. Alternative violates the "wolfxl-core carries no PyO3 dependency" rule from CLAUDE.md. | RFC-010, RFC-012 | **Decision: Approve new-crate path** (2026-04-25) — matches existing `wolfxl-core` precedent; keeps PyO3 boundary clean. RFC-010 / RFC-012 implementers create the crates. |
| 2 | **Tokenizer port** (RFC-012): port openpyxl's Bachtal tokenizer (~450 LOC) to Rust, OR wrap existing `python/wolfxl/calc/_parser.py`? | RFC-012 | **Decision: Approve Rust port** (2026-04-25) — perf budget (<1s for 100k formulas) requires it; existing parser doesn't support re-emission. Lives in new `crates/wolfxl-formula/`. |
| 3 | **`respect_dollar` semantics** (RFC-012 §5.5): default `false` (insert/delete coordinate-remap — `$A$1` shifts) for RFC-030/031, opt-in `true` (paste-style — `$` short-circuits) for RFC-034. P1 subagent claimed this matches Excel but didn't verify in Excel. | RFC-012, RFC-034 | **Pending — verify in Excel before RFC-012 starts** (5-min test, see "Next Steps" §0). Only RFC-012 / RFC-030 / RFC-031 / RFC-034 are blocked on this; everything else can proceed. |
| 4 | **`<conditionalFormatting>` replace-all semantics** (RFC-011 §5.5): caller-supplied CF block triggers full replacement of all existing CF? RFC-026 must read+merge first if it wants preservation. | RFC-011, RFC-026 | **Decision: Approve replace-all** (2026-04-25) — keeps merger contract clean; RFC-026 owns the read-merge-write loop. |
| 5 | **Hyperlink deletion sentinel** (RFC-022): change `_pending_hyperlinks` to use `None` value as explicit-delete sentinel rather than `pop()`? Behavior change at the Python layer. | RFC-022 | **Decision: Approve None-sentinel** (2026-04-25) — `pop()` loses deletion intent before flush. RFC-022 §5 spec is authoritative; never use `pop()` on `_pending_hyperlinks`. |
| 6 | **`WOLFXL_STRUCTURAL_PARITY=openpyxl` env flag** (RFC-030/031): opt-in flag to suppress wolfxl's better-than-openpyxl behavior (formulas/CF/DV shifting) for users who copy openpyxl docs? | RFC-030, RFC-031 | **Decision: Defer to post-1.1** (2026-04-25) — adds feature-flag complexity for marginal value. Loud divergence note in 1.1 release notes instead. |
| 7 | **App.xml Company/Manager** (RFC-020): wolfxl's Python `DocumentProperties` doesn't expose these; first dirty save silently drops them. | RFC-020 | **Decision: Accept as known regression** for 1.0 (2026-04-25). Follow-up issue: "Add Company/Manager fields to Python `DocumentProperties` (post-1.0)". |
| 8 | **VML write-mode bug** (RFC-023): native writer's `compute_margin` assumes default col widths; modify-mode patcher fixes only the patcher path, not write mode. | RFC-023 | **Decision: Accept** (2026-04-25) — matches openpyxl's preexisting behavior; RFC-023 §7 documents the seam. Follow-up issue: "Fix VML margin in native write mode (post-1.0)". |

---

## Phase Sequencing

| Phase | RFCs | Calendar | Notes |
|---|---|---|---|
| **P0 fix-it** | (no RFC) | ½ day | NotImplementedError stubs for structural ops. Ship before Phase 2. |
| **Phase 2 — Foundation** | 001, 010, 011, 012, 013 | ~2 weeks | W5 first (½d). Then 4 infra RFCs in parallel — bottleneck is 012 (L). |
| **Phase 3 — T1.5 modify-mode** | 020, 021, 022, 023, 024, 025, 026 | ~2-3 weeks | 7 RFCs in parallel pods. Review is the bottleneck, not coding. |
| **Phase 4a — Row/col structural** | 030, 031 | ~2 weeks | Mechanical once 012 + RFCs 022/024/025/026 land. |
| **Phase 4b — Range/sheet structural** | 034, 035, 036 | ~2-3 weeks | 035 is critical path (XL); 034/036 land in parallel. |
| **Phase 4c — Hardening** | (no RFCs) | 1 week | Fuzz, golden expansion, KNOWN_GAPS.md cleanup, doc updates, release notes. |
| **Phase 5 — Read + construction parity** | 040, 041, 042, 043, 044, 045, 046 | rolling 1.3 → 1.6 | Sprint Ι (1.3) read-side parity, Sprint Κ (1.4) `.xlsb`/`.xls`, Sprint Λ (1.5) encryption + image construction, Sprint Μ (1.6) chart construction. |

---

## Hard Cuts

- **End of Phase 3 (~3-week mark from research kickoff)**: ship as **WolfXL 1.0 — full modify-mode parity with openpyxl**. Closes the headline T1.5 gap. Structural ops still raise `NotImplementedError` pointing at 1.1.
- **End of Phase 4 (~6-7 week mark)**: ship as **WolfXL 1.1 — full openpyxl replacement**. Marketing claim defensible.

---

## Conventions

- Every RFC follows `000-template.md`. Section headings are fixed so dispatch agents can navigate mechanically.
- Frontmatter `Depends-on` and `Unblocks` are the source of truth for the DAG above. Update them when scope changes; this index will be regenerated.
- Verification: every RFC's §6 must address the 6-layer matrix (Rust unit, golden round-trip, openpyxl parity, LibreOffice, cross-mode, regression fixture). Absence is justified in §10.
- Implementation handoff: `python scripts/verify_rfc.py --rfc NNN` is the standardized "done" check. Author this script as part of Phase 2's first execution sub-task.

---

## Next Steps (after research → execution handoff)

### Sprint Δ — Phase 3 close + Phase 4 critical-path unblock (2026-04-26)

Wave A + Wave B of the next-slice sprint shipped in a single dispatch
of four parallel pods. The merge happened on `feat/native-writer`
(commits 6a50873 → 1c51233):

| RFC | Status | Branch | Merge commit |
|---|---|---|---|
| 021 — defined names | Shipped | `feat/rfc-021-defined-names` | `6a50873` |
| 024 — tables | Shipped | `feat/rfc-024-tables` | `f99b169` |
| 023 — comments + VML | Shipped | `feat/rfc-023-comments` | `de0f7c0` |
| 012 — formula xlator | Shipped | `feat/rfc-012-formula-xlator` | `876fae0` |

Verification: 914 pytest passed (1 deselected env-dep, 11 pre-existing
skips); 5 workspace cargo crates green;
`pytest tests/parity/ -q -x` → 97 passed + 1 skipped.

**Phase 3 is closed.** The codebase is ready to ship as **WolfXL 1.0
— full modify-mode parity with openpyxl**.

### Outstanding (carried into Phase 4 kickoff)

0. **Pre-step (gates RFC-030/031/034)**: Manual 5-minute Excel
   verification of `$A$1` shift behavior. See
   `Plans/rfcs/notes/excel-respect-dollar-check.md`. RFC-012 shipped
   with `ShiftPlan::respect_dollar` as a required field with no
   `Default`; the verification result determines what default RFC-030
   / RFC-031 wire up at the patcher boundary. RFC-012 itself does not
   need it.
1. **Phase 4a dispatch**: 030 + 031 in parallel (now unblocked by 012,
   021, 022, 024, 025, 026 all shipped). Estimated 2 weeks calendar.
2. **Phase 4b dispatch**: 034, 035, 036 — sequenced after 030/031;
   035 is XL critical path.
3. **Phase 4c hardening**: fuzz, golden expansion,
   `KNOWN_GAPS.md` cleanup, release notes.

The dependency DAG above is the source of truth for scheduling.

### Sprint Ξ ("Xi") — v1.7 launch slice (2026-04-27)

Public-launch slice (no pivot tables). Burns down v1.6.1 chart-stack
debt + refreshes migration / performance docs + materialises launch
posts. See `Plans/sprint-xi.md`.

Four pods landed inline on `feat/native-writer` (small-LOC slice;
no worktree fanout needed):

| RFC | Pod | Outcome |
|---|---|---|
| 050 — `Worksheet.remove_chart` / `replace_chart` + RichText title | α | Shipped |
| 051 — docs/migration overhaul | β | Shipped |
| 052 — docs/performance refresh | γ | Shipped |
| 053 — `Plans/launch-posts.md` drafts | δ | Shipped |

Sprint Ξ is closed; tag `v1.7.0` cut. Next: **Sprint Ν** (v2.0.0
pivot tables + pivot charts + public-launch-with-pivots).

### Sprint Ν ("Nu") — pivot tables + pivot charts → v2.0.0 (in progress, 2026-04-27)

User picked **Option A — full pivot construction** (~3-4 wk) on
2026-04-27 over the 80/20 refresh-on-open variant. Reasoning: the
"full openpyxl replacement" marketing claim requires pivots that
work without an Excel-side refresh round-trip — i.e. open the
wolfxl-emitted workbook in any OOXML-compliant reader and the
pivot's data is already populated. That requires authoring
`pivotCacheRecords{N}.xml`.

5 parallel pods + integrator. Pre-dispatch contract specs in
RFC-047 §10 / RFC-048 §10 / RFC-049 §10 are AUTHORITATIVE
(lesson #12 from Sprint Μ-prime: write the contract BEFORE pod
dispatch).

| RFC | Pod | Status | Branch (when dispatched) |
|---|---|---|---|
| 047 — pivot caches | α + β + γ | Approved | `feat/sprint-nu-pod-{alpha,beta,gamma}` |
| 048 — pivot tables | α + β + γ | Approved | `feat/sprint-nu-pod-{alpha,beta,gamma}` |
| 049 — pivot charts | δ | Approved | `feat/sprint-nu-pod-delta` |
| 054 — launch hardening | ε | Approved | `feat/sprint-nu-pod-epsilon` |

See `Plans/sprint-nu.md` for the full plan, OOXML pivot anatomy
diagram, calendar, risk register, and acceptance criteria.
