# Sprint Ο Pod 1A — Deferral Notes

> Created 2026-04-27 by the Pod 1A worker droid.
> **CLOSED 2026-04-27 by Pod 1A.5 follow-up — every deferred item is now landed.**
> Branch: `feat/sprint-omicron-pod-1a` (base) + `feat/sprint-omicron-pod-1a5` (closure).

## Status by section

| § | Deliverable | Landed? | SHA / notes |
|---|---|---|---|
| 1A.1 | Python class layer (PageSetup, PageMargins, HeaderFooter, SheetView, SheetProtection, properties, print_settings) | ✅ | Pod 1A — `7cdfceb`. All `to_rust_dict()` helpers present. `wolfxl.utils.protection.hash_password` matches openpyxl byte-for-byte for `"hunter2"` -> `"C258"`. |
| 1A.2 | Worksheet integration (lazy properties + `freeze_panes` ↔ `sheet_view.pane` mirror + `print_title_rows/cols` validators + `to_rust_setup_dict()`) | ✅ | Pod 1A — `6dd19b0`. All 6 user idioms from the post-Sprint-Ν audit no longer raise `AttributeError`. |
| 1A.3 | Native writer emit (`crates/wolfxl-writer/src/emit/sheet_setup.rs`) | ✅ | Pod 1A.5 — `f5ff68e` (parser+emitter) + `20ade12` (sheet_xml wiring). 5 emit helpers (`emit_sheet_views`, `emit_sheet_protection`, `emit_page_margins`, `emit_page_setup`, `emit_header_footer`) drive the new typed Worksheet model fields (slots 3 / 8 / 21 / 22 / 23). |
| 1A.4 | 5 new `SheetBlock` variants (`SheetViews`, `SheetProtection`, `PageMargins`, `PageSetup`, `HeaderFooter`) | ✅ | Pod 1A — `e6788ec`. |
| 1A.4 (cont.) | Patcher Phase 2.5n drain | ✅ | Pod 1A.5 — `6f475e1`. `XlsxPatcher::queue_sheet_setup_update` PyO3 method drains the §10 dict via `wolfxl_merger::merge_blocks`, sequenced AFTER pivots (Phase 2.5m) and BEFORE autoFilter (Phase 2.5o). |
| 1A.5 | PyO3 bindings + `parse/sheet_setup.rs` parser | ✅ | Pod 1A.5 — `f5ff68e` (parser) + `6f475e1` (PyO3 entrypoint in `src/wolfxl/sheet_setup.rs`). |
| 1A.6 | RFC-035 deep-clone extension | ✅ | Pod 1A.5 — `8d23b58`. `copy_worksheet` (modify + write modes) deep-copies the 5 sheet-setup slots + `print_titles` into the destination Worksheet proxy so per-sheet divergence is preserved. Workbook coordinator also propagates through Phase 2.5n. |
| 1A.7 | Tests | ✅ | Pod 1A — 105 class-layer tests (`1becee7`). Pod 1A.5 — 9 diffwriter + 16 parity + 7 copy_worksheet (= 32 new) on top, all passing. |

## Why this scope was originally deferred

The full RFC-055 deliverable spans Python + Rust + PyO3 + ~70 tests +
5 worktree integration points. A single worker-droid response could
not carry that volume end-to-end, so Pod 1A shipped only the
self-contained Python class layer plus the merger SheetBlock
variants. Pod 1A.5 picked up the Rust + patcher + native-writer +
deep-clone work as a follow-up and closed it without re-doing any of
Pod 1A's design.

## Pod 1A.5 commit map

| SHA | Subject |
|---|---|
| `b513daf` | Merge feat/sprint-omicron-pod-1a (Python + merger variants base) |
| `f5ff68e` | RFC-055 §10 — Rust `parse/sheet_setup.rs` parser + 5 emitters + 23 unit tests |
| `6f475e1` | RFC-055 Phase 2.5n — patcher sheet-setup drain + PyO3 `queue_sheet_setup_update` |
| `20ade12` | RFC-055 §5 — native writer sheet-setup emit (5 new Worksheet model fields) |
| `5422060` | RFC-055 Workbook flush hook + write-mode `set_sheet_setup_native` |
| `8d23b58` | RFC-055 §7 — 32 new tests + `copy_worksheet` deep-clone |

## Tolerable pre-existing failures (per task brief)

- subprocess `python` → `python3` mismatches.
- password-reads needing `msoffcrypto-tool`.
- 6 copy_worksheet namespace-prefix issues.

These were not introduced by Pod 1A or 1A.5.

## §10 contract drift observed

None. The Python `to_rust_setup_dict()` emits exactly the §10 dict
shape; the Rust `parse_sheet_setup_payload` consumes it verbatim
without re-shaping. The shared `SheetSetupBlocks` struct in
`crates/wolfxl-writer/src/parse/sheet_setup.rs` is the single source
of truth for the Rust side and is consumed by both the native writer
(`emit/sheet_xml.rs`) and the patcher (`src/wolfxl/sheet_setup.rs`)
via the same set of emit helpers.
