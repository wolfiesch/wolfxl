# RFC-NNN: <One-Line Title>

Status: Draft | Researched | Approved | In-Progress | Shipped
Owner: pod-PN
Phase: 2 | 3 | 4 | 5
Estimate: S | M | L | XL
Depends-on: RFC-NNN, RFC-NNN
Unblocks: RFC-NNN, RFC-NNN

> **S** = ≤2 days; **M** = 3-5 days; **L** = 1-2 weeks; **XL** = 2+ weeks (calendar, with parallel subagent dispatch + review).

## 1. Problem Statement

What user-visible operation is broken or missing today? Quote the
exact error or stub message and link the line numbers. Show one or
two short user code samples that hit the gap. State the **target
behaviour** in one paragraph.

## 2. OOXML Spec Surface

Which ECMA-376 sections govern the parts you touch? List the relevant
elements, attributes, content-type URIs, and rels-type URIs. Note any
schema-ordering constraints (`CT_Worksheet` child order is the most
common gotcha). Call out spec corners that the codebase has not yet
encountered (e.g., R1C1 references, structured table refs, threaded
comments).

## 3. openpyxl Reference

What does openpyxl do? Quote the relevant source files and line
numbers (`.venv/lib/python3.14/site-packages/openpyxl/...`). Note
each public-API spelling, alias, and edge case. Explicitly enumerate
**what we do NOT copy** (read-path helpers, lxml-specific quirks,
internal validators).

## 4. WolfXL Surface Area

### 4.1 Python coordinator

Where in `python/wolfxl/` does the new code land? What gets removed
(typically a `NotImplementedError` block)? What gets added (typically
a `_flush_<feature>_to_patcher` method)?

### 4.2 Patcher (modify mode)

Which Rust module under `src/wolfxl/` (or which `crates/` workspace
crate) owns the new logic? Streaming splice vs full rewrite — justify
the choice. List the new `XlsxPatcher` PyMethods. List any new
`SheetBlock` / `ContentTypeOp` variants.

### 4.3 Native writer (write mode)

If the feature already round-trips in `crates/wolfxl-writer/`, note
the seam. If not, declare the asymmetry and link the follow-up.

## 5. Implementation Sketch

Walk the implementation phase by phase. Use OOXML-pseudocode for the
emit shape. Call out any cross-sheet aggregation (Phase-2.5 ordering
matters). Specify the no-op invariant explicitly: empty queue → no
file change, byte-identical output.

### 5.1, 5.2, … Subsections per major design decision.

## 6. Verification Matrix

Six layers required (omitted layers must be justified in §10):

1. **Rust unit tests** — `cargo test -p <crate>` covering every
   branch.
2. **Golden round-trip (diffwriter)** — `WOLFXL_TEST_EPOCH=0 pytest
   tests/diffwriter/`.
3. **openpyxl parity** — `pytest tests/parity/`.
4. **LibreOffice cross-renderer** — manual; document the fixture +
   expected behaviour.
5. **Cross-mode** — write-mode + modify-mode produce equivalent
   files for the same input.
6. **Regression fixture** — checked into `tests/fixtures/`.

The standardized "done" gate is `python scripts/verify_rfc.py --rfc
NNN`.

## 7. Cross-Mode Asymmetries

Anywhere write-mode and modify-mode diverge intentionally (e.g.,
write-mode bug that modify-mode fixes; opposite is also possible).
Document the seam, file a follow-up if applicable.

## 8. Risks

| # | Risk | Likelihood | Impact | Mitigation |
|---|------|-----------|--------|-----------|
| 1 | … | low/med/high | low/med/high | … |

## 9. Effort Breakdown

| Slice | Estimate | Notes |
|-------|----------|-------|
| Research | … | |
| Rust impl | … | |
| Python wiring | … | |
| Tests | … | |
| Review | … | |

## 10. Out of Scope

What is intentionally deferred? Cite the follow-up RFC or issue.
Anything not listed here is either in scope or a bug. Do not allow
silent gaps.

## Acceptance

(Filled in after Shipped. Format:)

- Commit: `<sha>` — feat/test commit subject
- Verification: `python scripts/verify_rfc.py --rfc NNN` GREEN at `<sha>`
- Date: YYYY-MM-DD
