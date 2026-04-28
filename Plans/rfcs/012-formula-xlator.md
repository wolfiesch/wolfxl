# RFC-012: Formula Reference Translator

Status: Shipped
Owner: pod-P1
Phase: 2
Estimate: L
Depends-on: RFC-001
Unblocks: RFC-021, RFC-030, RFC-031, RFC-034, RFC-035, RFC-036

## 1. Problem Statement

Six Phase-3+ RFCs need to update cell references inside formulas as a
consequence of structural mutations:

| Trigger (RFC) | Required transform |
|---|---|
| `Worksheet.insert_rows(row, n)` (RFC-030) | Every relative AND absolute reference at row ≥ `row` shifts down by `n`. |
| `Worksheet.delete_rows(row, n)` (RFC-030) | Every reference inside the deleted range becomes `#REF!`; references below shift up by `n`. |
| `Worksheet.insert_cols(col, n)` / `delete_cols` (RFC-031) | Symmetric to rows. |
| `Worksheet.move_range(src, dst)` (RFC-034) | Relative refs from cells inside `src` re-anchor by the `dst-src` offset; absolute refs unchanged; refs into the moved range from outside also shift. |
| `Workbook.copy_worksheet(src, name)` (RFC-035) | Defined-name re-scoping (handled by RFC-021); within-sheet relative refs re-anchor (no change in same-sheet copy); 3-D refs to other sheets unchanged. |
| `Worksheet.title = "newname"` (RFC-036) | Every 3-D ref `'oldname'!A1` (or `oldname!A1`) anywhere in the workbook updates to the new sheet name. |

See full spec body in this file's git history (the RFC was originally
researched at length; the executive shape is captured below).

## 2. OOXML Spec Surface

ECMA-376 Part 1 §18.17 — Formulas.

Reference syntaxes the translator MUST handle:

| Form | Example | Notes |
|---|---|---|
| A1 relative | `A1`, `B5` | Shifts on row/col insert/delete; re-anchors on copy/move. |
| A1 absolute | `$A$1`, `$B$5` | Shifts on insert/delete (coordinate remap); does NOT re-anchor on paste. |
| A1 mixed | `$A1`, `A$1` | Each part independently absolute. |
| Range | `A1:B5`, `$A$1:$B$5` | Translate both endpoints independently. |
| Whole row | `1:1`, `2:5` | Only shifts on row insert/delete. |
| Whole column | `A:A`, `B:D` | Only shifts on column insert/delete. |
| 3-D unquoted | `Sheet2!A1` | Sheet portion updates only on rename. |
| 3-D quoted | `'My Sheet'!A1`, `'O''Brien'!A1` | Apostrophes inside doubled. |
| External book | `[Book2.xlsx]Sheet1!A1` | OUT OF SCOPE — pass through. |
| Structured table | `Table1[Col1]`, `Table1[[#This Row], [Col1]]` | Bound by table name; do NOT shift on row/col delta. |
| Defined name | `=MyTotal` | Pass through; RFC-021 owns rescoping. |
| Function | `=SUM(...)` | Pass through. |
| String literal | `=A1&"text"` | Refs INSIDE strings must NOT translate. |
| Error | `=#REF!`, `=#N/A` | Pass through. |

## 3. openpyxl Reference

Direct port of `openpyxl.formula.tokenizer.Tokenizer` (Bachtal-style).
See `crates/wolfxl-formula/src/tokenizer.rs`. Translator extends the
openpyxl shape (which is paste-style only) with insert/delete coordinate-
remap semantics + sheet-rename + tombstone clipping for `delete_rows`.

## 4. WolfXL Surface Area

### 4.1 Public crate

`crates/wolfxl-formula/`. Pure Rust (no PyO3 — same workspace
invariant as `wolfxl-core`).

```rust
pub fn shift(formula: &str, plan: &ShiftPlan) -> String;
pub fn rename_sheet(formula: &str, old: &str, new: &str) -> String;
pub fn move_range(formula: &str, src: &Range, dst: &Range, respect_dollar: bool) -> String;

pub fn translate(formula: &str, delta: &RefDelta) -> Result<String, TokenizeError>;
pub fn translate_with_meta(formula: &str, delta: &RefDelta) -> Result<TranslateResult, TokenizeError>;
pub fn tokenize(formula: &str) -> Result<Vec<Token>, TokenizeError>;
```

`RefDelta` carries: `rows`, `cols`, `anchor_row`, `anchor_col`,
`sheet_renames`, `deleted_range`, `formula_sheet`,
`deleted_range_sheet`, `respect_dollar`, `move_src`, `move_dst_*`.

### 4.2 Patcher integration

The PyO3 cdylib (`src/wolfxl/`) calls `wolfxl-formula` from RFC-030,
031, 034, 035, 036 patcher paths. The translator is pure (no I/O).

## 5. Algorithm

### 5.1 Approach

Pure-Rust port of openpyxl's tokenizer. Token-level translation: only
`Operand/Range`-subkind tokens are inspected. String literals
(`Operand/Text`) are opaque, so refs inside `INDIRECT("A1")` are NOT
modified.

### 5.2 Tokenizer

See `crates/wolfxl-formula/src/tokenizer.rs`. ~450 LOC port of
`tokenizer.py`.

### 5.3 Reference parsing

See `crates/wolfxl-formula/src/reference.rs`. Strips optional sheet
prefix (quoted or unquoted), classifies as Cell / Range / RowRange /
ColRange / Table / ExternalBook / Name / Error.

### 5.4 Per-kind translation

See `crates/wolfxl-formula/src/translate.rs`. Sheet rename → tombstone
check → move-range re-anchor → axis shift → bounds check.

### 5.5 respect_dollar — OPEN QUESTION (BLOCKER)

The default value of `respect_dollar` for `shift` operations has not
yet been verified in Excel. See
`Plans/rfcs/notes/excel-respect-dollar-check.md`.

Until verified:

- `ShiftPlan::respect_dollar` is a required field with no `Default`.
- `RefDelta::respect_dollar` likewise has no documented default for
  shift ops.
- README in `crates/wolfxl-formula/` flags this prominently.
- RFC-030 / RFC-031 must NOT proceed past patcher-wiring without
  resolution.

`move_range` always uses `respect_dollar=true` independently of this
question.

### 5.6 Re-emission

Tokenizer's `render(&[Token])` concatenates token values verbatim. For
identity (no-op) translation, output is byte-identical to input.

### 5.7 INDIRECT pass-through

`translate_with_meta` returns `has_volatile_indirect=true` if
INDIRECT/OFFSET/ADDRESS/INDEX/CHOOSE/HYPERLINK appears as a function
opener. Caller decides whether to surface a warning.

### 5.8 Performance

100k complex formulas in 137 ms on M1. Budget 1s; 7x headroom.

## 6. Test Plan

Implemented in `crates/wolfxl-formula/src/tests.rs` and
`tokenizer.rs` mod tests:

- 82 unit tests: every reference syntax in §2 + every semantic in §5.
- `t41_synthgl_corpus_round_trip`: identity translation byte-identical
  for ≥100 formulas drawn from `tests/parity/fixtures/synthgl_snapshot/`
  + `tests/fixtures/`. Ran with 103 formulas — 100% pass.
- `t42_perf_100k_formulas_under_1s`: 100k complex-formula shifts
  finish in 137 ms.

## 7. Migration / Compat Notes

No public Python API change in this RFC. Surfaced via RFC-030/031/
034/035/036.

Behavior diffs vs openpyxl:

1. Absolute refs DO shift on insert/delete row (coordinate-remap).
   openpyxl has no API for this transform; ours matches Excel.
2. 3-D refs CAN be rewritten on sheet rename (openpyxl: no API).
3. INDIRECT contents NOT rewritten (openpyxl: silently rewrites
   substring matches — incorrect).

## 8. Risks & Open Questions

1. (HIGH→MITIGATED) Tokenizer port coverage. Mitigated by 103-formula
   corpus round-trip test on synthgl + tests/fixtures.
2. (MED→PASSED) Performance budget. 137 ms for 100k formulas (7x
   headroom).
3. (OPEN) `respect_dollar` default — see §5.5. **BLOCKER** for
   RFC-030/031.

## 9. Effort Breakdown

Shipped at ~750 LOC (lib + tokenizer + reference + translate +
tests). Did not require regex or `once_cell` — pure stdlib + zip
(dev-only).

## 10. Out of Scope

- External-workbook references (passthrough).
- Defined-name target rewriting (RFC-021).
- Table-name rewriting (RFC-024-Y2).
- Shared-formula expansion (RFC-021).
- Array-formula `<f t="array" ref="...">` `ref` attribute updates
  (RFC-030).
- Calc-engine integration (RFC-040).
- R1C1 notation.
- Locale-specific function names.

## Acceptance

Shipped in commit `7872c1c` (`feat(formula): RFC-012 — wolfxl-formula
crate (tokenizer + translator)`).

Verification:

- `cargo test -p wolfxl-formula --quiet` → 82 passed, 0 failed.
- `cargo test -p wolfxl-core -p wolfxl-writer -p wolfxl-rels -p
  wolfxl-merger --quiet` → 258 passed, 0 failed (no regressions).
- `t41_synthgl_corpus_round_trip` — 103 formulas, 100% byte-identical
  on identity translation.
- `t42_perf_100k_formulas_under_1s` — 137 ms (7x under budget).

**BLOCKER carried forward**: `respect_dollar` default for shift ops
remains unverified. RFC-030 / RFC-031 implementers MUST run the
5-minute check at `Plans/rfcs/notes/excel-respect-dollar-check.md`
before proceeding past patcher-wiring.
