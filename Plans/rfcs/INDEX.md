# RFC Index — openpyxl-Replacement Gap Closure

> **Source plan**: kept in conversation transcript; see `Plans/launch-posts.md` for the marketing artifacts that depend on this work landing.
> **Scope**: Modify-mode parity (T1.5) + structural worksheet ops + W5 rip-out. Excludes Phase 2-5 read gaps and construction-side stubs (chart construction, NamedStyle, etc.).
> **Goal**: Defensible "full openpyxl replacement" claim by end of Phase 4.

This file is regenerated from each RFC's frontmatter. Edit a source RFC, not this index.

---

## Status Table (16 RFCs)

| ID | Title | Status | Phase | Estimate | Depends-on | Unblocks |
|---|---|---|---|---|---|---|
| 001 | W5 — rust_xlsxwriter rip-out | Researched | 2 | S | — | (clean baseline) |
| 010 | Infra — rels graph parser/serializer | Shipped | 2 | M | 001 | 022, 023, 024, 035 |
| 011 | Infra — XML-block-merger primitive | Shipped | 2 | M | 001 | 022, 023, 024, 025, 026 |
| 012 | Infra — formula reference translator | Researched | 2 | L | 001 | 021, 030, 031, 034, 035 |
| 013 | Infra — patcher extensions (ZIP-add, ancillary parts, two-phase flush) | Shipped | 2 | M | 001 | 022, 023, 024, 035 |
| 020 | T1.5 — Document properties | Shipped | 3 | S | 013 | — |
| 021 | T1.5 — Defined names mutation | Researched | 3 | M | 011, 012 | 030, 031, 034, 035, 036 |
| 022 | T1.5 — Hyperlinks | Shipped | 3 | M | 010, 011, 013 | 030, 031, 035 |
| 023 | T1.5 — Comments + VML drawings | Researched | 3 | L | 010, 011, 013 | 030, 031, 035 |
| 024 | T1.5 — Tables | Researched | 3 | M | 010, 011, 013 | 030, 031, 035 |
| 025 | T1.5 — Data validations | Shipped | 3 | M | 011 | 030, 031, 035 |
| 026 | T1.5 — Conditional formatting | Shipped | 3 | M | 011 | 030, 031, 035 |
| 030 | Structural — insert_rows / delete_rows | Researched | 4 | L | 012, 021, 022, 024, 025, 026 | 034 |
| 031 | Structural — insert_cols / delete_cols | Researched | 4 | L | 012, 021, 022, 024, 025, 026 | 034 |
| 034 | Structural — move_range | Researched | 4 | L | 012, 030, 031 | — |
| 035 | Structural — copy_worksheet | Researched | 4 | XL | 010, 012, 013, 021, 022, 023, 024, 025, 026 | — |
| 036 | Structural — move_sheet | Researched | 4 | M | 021 | — |

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

0. **Pre-step (gates RFC-012 only)**: Manual 5-minute Excel verification of `$A$1` shift behavior for question #3. Open Excel, type `=$A$1+$B$1` in `B5`, right-click row 3 → Insert, observe whether `B6`'s formula reads `=$A$1+$B$1` (no shift, default `respect_dollar=true`) or `=$A$2+$B$2` (shift, default `respect_dollar=false`). If shifted: keep RFC-012 §5.5 default `false`. If not: flip default to `true` and update RFC-012 §5.5 + RFC-030 + RFC-031 reference-rewriting paths. Do not start RFC-012 without this check.
1. **Decisions**: Resolved 2026-04-25 for 7 of 8 questions above. Question #3 is the only outstanding decision; it gates RFC-012 only (everything else can proceed).
2. **P0 fix-it commit**: Shipped — see commit `1af6ba3` (`feat(api): add NotImplementedError stubs for 7 structural ops`).
3. **Phase 2 kickoff**: `runplan Plans/rfcs/001-w5-rip-out.md` first (RFC-001 shipped in `4a840e9`); then dispatch 010/011/013 in parallel pods (each pod: implementer → spec-reviewer → code-quality-reviewer per `superpowers:subagent-driven-development`). RFC-012 starts after step 0 above.
4. **Phase 3 dispatch**: Once Phase 2 lands, fan out 7 T1.5 RFCs in parallel pods; bottleneck is review, not coding.
5. **Phase 4 dispatch**: Sequenced — 030/031 in parallel first, then 034/035/036.

The dependency DAG above is the source of truth for scheduling. Anything blocking 012 (formula translator) is on the critical path.
