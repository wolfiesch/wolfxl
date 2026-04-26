# Follow-up — Native writer VML margin uses default column widths

**Source:** RFC-023 (Comments + VML drawings) — INDEX decision #8
**Status:** Open (post-1.0)
**Owner:** TBD

## Problem

`crates/wolfxl-writer/src/emit/drawings_vml.rs::compute_margin` hard-codes
`COL_WIDTH_PT = 48.0`. That value is the OOXML default column width in
points and is correct only on sheets whose `<cols>` block declares no
overrides. Sheets with custom column widths render comment shapes at
visually incorrect margin positions (the marker triangle still sits
on the right cell, but the popup body floats over the wrong area).

## Why it ships unfixed

- Modify-mode (RFC-023) patches around the bug by parsing the sheet's
  `<cols>` block at flush time (`src/wolfxl/comments.rs::compute_margin_with_widths`).
  This restores correct positioning for any file the patcher mutates.
- Fixing the writer requires threading `Worksheet.cols` through
  `crates/wolfxl-writer/src/emit/drawings_vml.rs` so the emitter can
  honor per-column overrides at write time. That's a writer-side
  refactor disproportionate to the visual cost (excel still renders
  the comment, just slightly mispositioned).
- Matches openpyxl's preexisting behavior — users migrating from
  openpyxl will see no regression.

## Acceptance criteria for the fix

1. `crates/wolfxl-writer/src/sheet/worksheet.rs` exposes the parsed
   `<cols>` block to the emitter (or moves margin computation into
   the worksheet itself).
2. `compute_margin` accepts a `&ColWidthMap` (or equivalent) and
   walks per-column widths in points, mirroring the patcher's
   `compute_margin_with_widths`.
3. New round-trip test in `tests/test_native_writer_comments.py`
   covers a sheet whose first column is set to width=4 and asserts
   the emitted `margin-left` lands on the correct cell boundary.
4. Patcher's `compute_margin_with_widths` and the writer's new
   helper either share code or are tested for parity.

## Notes

- INDEX decision #8 (2026-04-25): "Accept; RFC-023 §7 documents the seam."
- RFC-023's seam doc lives at the top of `src/wolfxl/comments.rs`.
