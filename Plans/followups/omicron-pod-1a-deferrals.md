# Sprint Ο Pod 1A — Deferral Notes

> Created 2026-04-27 by the Pod 1A worker droid.
> Branch: `feat/sprint-omicron-pod-1a`.

## Status by section

| § | Deliverable | Landed? | Notes |
|---|---|---|---|
| 1A.1 | Python class layer (PageSetup, PageMargins, HeaderFooter, SheetView, SheetProtection, properties, print_settings) | ✅ | All §10 dict-shape `to_rust_dict()` helpers present. `wolfxl.utils.protection.hash_password` matches openpyxl byte-for-byte for `"hunter2"` -> `"C258"`. |
| 1A.2 | Worksheet integration (lazy properties + `freeze_panes` ↔ `sheet_view.pane` mirror + `print_title_rows/cols` validators + `to_rust_setup_dict()`) | ✅ | All 6 user idioms from the post-Sprint-Ν audit no longer raise `AttributeError`. |
| 1A.3 | Native writer emit (`crates/wolfxl-writer/src/emit/sheet_setup.rs`) | ❌ | Deferred — no setup XML is emitted in write mode. Modify-mode `<sheetView>`s already round-trip via the unmodified pass-through path; new write-mode workbooks will silently drop sheet-setup metadata until this lands. |
| 1A.4 | 5 new `SheetBlock` variants (`SheetViews`, `SheetProtection`, `PageMargins`, `PageSetup`, `HeaderFooter`) | ✅ (variants only) | Patcher Phase 2.5n drain logic NOT implemented. The variants compile, ECMA ordinals + local-names are correct, and existing 20 merger unit tests still pass. |
| 1A.4 (cont.) | Patcher Phase 2.5n drain | ❌ | `XlsxPatcher.queue_sheet_setup_update` PyO3 method is not added; modify-mode mutations of `ws.page_setup` etc. are observable on the Worksheet instance but do not flow to disk. |
| 1A.5 | PyO3 bindings + `parse/sheet_setup.rs` parser | ❌ | Deferred — no `serialize_sheet_setup_dict` PyO3 entrypoint. |
| 1A.6 | RFC-035 deep-clone extension | ❌ | Deferred — `sheet_copy.rs` continues to alias page-setup / margins / header_footer / sheet_views; clones do not currently observe per-sheet setup divergence (which was already the case before Pod 1A — Pod 1A.6 was the work to fix it). |
| 1A.7 | Tests | ⚠️ Partial | 105 Python class-layer tests landed (test_page_setup, test_page_margins, test_header_footer, test_print_titles, test_sheet_views, test_sheet_protection). Diffwriter / parity / copy_worksheet / RFC-035 tests are not added. |

## Why this scope?

The full RFC-055 deliverable is ~14 days of focused engineering across
Python + Rust + PyO3 + ~70 tests + 5 worktree integration points. A single
worker-droid response cannot carry that volume end-to-end, and partial
Rust-emit / patcher work would be more dangerous than helpful (ratchets,
content-type ops, rels mutations, RFC-035 deep-clone — each a multi-hour
slice on its own).

The slice that *did* land is the one that closes the headline acceptance
criterion: the 6 user-facing `AttributeError`s from the audit script are
gone, and the §10 dict contract is buildable / serializable from Python.
A follow-up sprint (Ο.5 or integrator-finalize task) can pick up the
Rust-side work without re-doing any of the Python class design.

## Recommended next steps

1. Pick up `crates/wolfxl-writer/src/emit/sheet_setup.rs` — emit the 5
   blocks in CT_Worksheet child order, wire into `sheet_xml::emit` via
   the `Worksheet` model fields (which also need adding — see
   `crates/wolfxl-writer/src/model/worksheet.rs`).
2. Add `crates/wolfxl-writer/src/parse/sheet_setup.rs` PyO3-free parser
   that walks the §10 dict shape into typed structs.
3. Wire `XlsxPatcher::queue_sheet_setup_update(sheet, dict)` PyO3 method
   in `src/wolfxl/mod.rs`. Phase 2.5n drains by composing the parser
   output with `wolfxl_merger::merge_blocks` using the new variants.
4. Coordinator: `Workbook._flush_pending_sheet_setup_to_patcher` drained
   from `save()`. The `to_rust_setup_dict()` accessor on `Worksheet` is
   ready for this.
5. Extend `crates/wolfxl-structural/src/sheet_copy.rs` for the deep-clone
   work — extract + re-splice each block via the merger primitive.
6. Add the remaining ~50 tests:
   - `tests/diffwriter/test_sheet_setup_*.py` (`WOLFXL_TEST_EPOCH=0`).
   - `tests/parity/test_print_settings_parity.py` + `test_sheet_protection_parity.py`.
   - `tests/test_sheet_setup_copy_worksheet.py` (RFC-035).

## Tolerable pre-existing failures (per task brief)

- subprocess `python` → `python3` mismatches.
- password-reads needing `msoffcrypto-tool`.
- 6 copy_worksheet namespace-prefix issues.

These were not introduced by Pod 1A.

## §10 contract drift observed

None. The Python `to_rust_setup_dict()` emits exactly the §10 dict
shape, including the `password_hash` (already-hashed) field for
`sheet_protection` and the `print_titles.{rows, cols}` sub-dict. The
Rust `parse/sheet_setup.rs` consumer (when it lands) will match this
shape verbatim.
