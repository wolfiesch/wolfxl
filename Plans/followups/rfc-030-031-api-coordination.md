# RFC-030 / RFC-031 API Coordination

> Filed by pod-031 on 2026-04-25 during the RFC-031 implementation slice.

## Context

Pod-031 reached the implementation slice (Half 2) before pod-030 had
landed. Per the RFC-031 task brief: "If pod-030 has not finished by
the time you reach Half 2: stub `wolfxl-structural` yourself with a
minimal API that takes `Axis`. Coordinate via the API contract:
`apply_axis_shift(axis, idx, n, sheet_state)`,
`apply_workbook_shift(axis, idx, n, &mut workbook_state)`. RFC-030
must accept the same API or fix it during reconciliation."

Pod-031 created `crates/wolfxl-structural/` with the following
shape:

```rust
pub struct AxisShift {
    pub axis: Axis,    // wolfxl_formula::Axis
    pub at: u32,
    pub n: i32,
}

pub fn apply_axis_shift(shift: AxisShift, sheet_xml: &[u8])
    -> Result<Vec<u8>, StructuralError>;

// Stub — implementation deferred (see "Open question" below).
pub fn apply_workbook_shift(shift: AxisShift, sheet_name: &str,
                             workbook_xml: &[u8]) -> Vec<u8>;
```

RFC-030 should adopt this signature on rebase. The patcher's
`queue_axis_shift(sheet, axis: &str, at: u32, n: i32)` already
accepts `axis="row"`, so the only missing piece on the row side is
filling in the `Axis::Row` branches in `sheet_shift.rs` (which
currently handles `<row r>` attribute only).

## Open question — `apply_workbook_shift` signature

RFC-031's first slice ships `apply_workbook_shift` as a stub that
returns the workbook XML unchanged. Real implementation needs:

1. The per-defined-name **scope** (workbook-scope vs
   `localSheetId="N"`-scope) — RFC-021's `defined_names::merge_defined_names`
   parses this but does NOT expose it to a non-flush caller.
2. A way to translate every formula text inside `<definedName>` via
   `wolfxl_formula::shift` AND filter by sheet name (only names
   pointing at the shifted sheet should be rewritten).

Two paths forward:

- **Path A** — extend RFC-021's `defined_names::merge_defined_names`
  with an optional `Vec<RefMutation>` argument that gets applied to
  each name's formula text before re-serialization.
- **Path B** — add a new `defined_names::apply_axis_shift_inplace`
  function that walks the workbook XML, finds the `<definedNames>`
  block, and rewrites in place. Symmetric to RFC-031's per-sheet
  shift.

Either RFC-030 or RFC-034 is likely the right place to land this
(both consume the same path). pod-031 leaves it as a stub so the
patcher's call site is stable.

## Open question — patcher-block shifts (hyperlinks, DV, CF, tables)

RFC-031's first slice does NOT shift these because they live in
their own queue paths (`queued_hyperlinks`, `queued_dv_patches`,
`queued_cf_patches`, `queued_tables`). Their existing flush logic
emits the user's queued patches verbatim into the merged sheet
XML; if the user calls `insert_cols` AFTER `add_hyperlink`, the
hyperlink's `ref` is at pre-shift coordinates.

Two paths:

- **Coordinator-side shift** — when `insert_cols`/`delete_cols`
  runs, sweep the Python-side pending queues (`_pending_hyperlinks`,
  `_pending_data_validations`, `_pending_conditional_formats`,
  `_pending_tables`) and apply the shift before flushing.
- **Patcher-side post-merge shift** — extend the Phase 2.6 axis-shift
  pass to also walk the merged sheet XML's `<hyperlinks>`,
  `<dataValidations>`, `<conditionalFormatting>`, `<tableParts>`
  blocks. (Tables also need their own
  `xl/tables/tableN.xml` rewritten — see §5.4 of RFC-031.)

Pod-031 recommends path 2 (post-merge): it requires no Python-side
plumbing and naturally handles user code that mixes mutate+shift in
either order. RFC-030 should adopt the same approach for its tests
to mirror the col tests.

## Open question — `<tableColumn>` removal on delete

RFC-031 §5.4 specifies the algorithm. Pod-031's first slice does
NOT implement it (the `tables.rs` extension referenced in the spec
isn't present in the shipped `wolfxl-structural` crate). Follow-up:
land alongside the patcher-block shifts above.

## Action items

- [ ] pod-030: rebase onto pod-031's `feat/rfc-031-insert-delete-cols`
  branch, fill in the `Axis::Row` paths in
  `crates/wolfxl-structural/src/sheet_shift.rs`, and confirm the
  `AxisShift` / `apply_axis_shift` signature matches expectations.
- [ ] pod-030 or follow-up pod: implement the patcher-block shifts
  (hyperlinks / DV / CF / tables / `<tableColumn>` removal).
- [ ] pod-030 or follow-up pod: implement `apply_workbook_shift`
  for defined names.
- [ ] Update RFC-031 §5 and §10 once the follow-ups land; promote
  the spec deviations from "deferred" to "shipped".
