# Excel `respect_dollar` Verification (gates RFC-012)

> Source: `Plans/rfcs/INDEX.md` open question #3 + sprint plan `2026-04-26-sprint-plan-next-openpyxl-parity-slice-close-phase-3-unblock-phase-4`.
> Status: **PENDING — gates RFC-012 default-only.** RFC-012 itself is
> shipped (no default — `respect_dollar` is a required field on
> `ShiftPlan` / `RefDelta`). RFC-030 / RFC-031 must NOT proceed past
> patcher-wiring without this check.

## What we are verifying

When a row is inserted in Excel, do **absolute** references (`$A$1`)
**shift** alongside relative references (`A1`), or do they stay put?

- **If they shift**: keep RFC-012 §5.5 default `respect_dollar=false`
  (current spec — insert/delete is a coordinate-space remap, `$` is
  ignored for the shift). Add a `Default` impl that uses `false`.
- **If they stay put**: flip RFC-012 §5.5 default to `respect_dollar=true`
  and patch RFC-030 + RFC-031 reference-rewriting paths
  (~30 LOC patch each).

## Repro steps (5 minutes, any Excel)

1. Open Excel. New blank workbook.
2. In `A5`, type any value (e.g., `1`).
3. In `B5`, type the formula:

   ```
   =$A$1+$B$1
   ```

   Press Enter.
4. Right-click row **3** header → **Insert**. (Inserts a new row above
   row 3; existing rows 3+ shift down by 1, so old `B5` is now `B6`.)
5. Click the cell that used to be `B5` (now `B6`) and read the formula
   bar.

## Decision matrix

| Formula bar shows… | Decision | Patch needed in RFC-012 |
|---|---|---|
| `=$A$2+$B$2` | absolute refs DID shift | **none** — `respect_dollar=false` is correct (current implementation). Add `Default` impl. |
| `=$A$1+$B$1` | absolute refs did NOT shift | **flip default** `respect_dollar=true`; update §5.5 + RFC-030 + RFC-031 |

## Reporting back

Once verified, edit `Plans/rfcs/INDEX.md` open-questions table row #3
with the result and date. Add `Default` impl to
`crates/wolfxl-formula/src/translate.rs` for `ShiftPlan` and `RefDelta`
using whichever default the verification picks. Update
`crates/wolfxl-formula/README.md` to remove the BLOCKER notice.

## Why this matters

RFC-012 ships without a default to force callers to make an informed
decision. Once verified, the default becomes part of the canonical API
and downstream callers (RFC-030/031) can omit the field.
