# Codex Handoff — G06: Image replace / remove public API

> **Sprint**: S1 — Bridge Completion
> **Branch**: `feat/parity-G06-image-replace-remove`
> **Sibling pods**: G03 (diagonal borders), G04 (protection), G05 (NamedStyle), G07 (DataTableFormula)
> **Merge gate**: independent codepaths — no sibling-blocking

## 1. Goal

Add `ws.remove_image(index_or_image)` and `ws.replace_image(index_or_image, new_image)` to the public Worksheet API so the `images_replace_remove` probe in the openpyxl compat oracle stops xfailing.

The probe loads a workbook with `modify=True`, calls `ws.remove_image(0)`, saves, reloads, and expects the saved file to have one fewer image than it started with. Replace is the same dance plus an add.

## 2. Files to touch

| File | What to do |
|---|---|
| `python/wolfxl/_worksheet.py:1259–1265` | Promote `_images` (currently underscore-prefixed property returning the live list) to a public `ws.images` if it is not already public, and add `ws.remove_image(...)` + `ws.replace_image(...)` methods next to the existing `add_image(...)` (line 1267–1282). |
| `python/wolfxl/_worksheet_media.py:391–406` | The add-side queues `_pending_images`. Add a queue concept for removals (e.g. `_pending_image_deletions: list[int]`) and a queue for replacements (or model replace as remove+add internally). Mirror the comments-delete pattern. |
| `python/wolfxl/_workbook_patcher_flush.py:82–95` | The current flush drains `_pending_images`. Extend it to drain the deletion queue and call into a new patcher entry point. Mirror the comments delete pattern at `:98–112` (`queue_comment_delete`). |
| `src/wolfxl/patcher_drawing.rs:130, 159` | Add a `queue_image_remove(idx)` (or similar) entry point that rewrites `xl/drawings/drawingN.xml` and `xl/drawings/_rels/drawingN.xml.rels` to drop the indicated drawing element + its rel. The add entry point is the template. |
| `src/wolfxl/mod.rs:1117` | Expose the new patcher entry through PyO3 next to the existing `queue_image_add`. |
| `tests/test_images_modify.py` (extend) | Add focused round-trip tests for `remove_image(0)` and `replace_image(0, new)` in modify mode: assert the resulting file has the expected count of images, with the survivor at the expected anchor. |

## 3. Reference exemplar

**Comments** (`_pending_comments` dict on `Worksheet`, drained at `_workbook_patcher_flush.py:98–112` via `queue_comment_delete`) is the closest analogue: queue add + queue delete via separate patcher entry points; flush handles both kinds. The image case is simpler — images are positionally indexed in a list, not keyed by cell address.

A second useful reference: `remove_chart` in `_worksheet_media.py:113–123` already removes from the in-process `_pending_charts` list. Mirror that pattern for the in-process image case (i.e. removing an image that was queued in this same session but not yet flushed). The harder case — removing an image that was loaded from an existing file — needs the patcher path.

## 4. Acceptance tests

```bash
# 1. The compat-oracle probe flips from xfail to pass.
uv run --no-sync pytest tests/test_openpyxl_compat_oracle.py -k images_replace_remove -q
# expect: 1 passed (was xfail)

# 2. Existing image suite stays green.
uv run --no-sync pytest tests/test_images_write.py tests/test_images_modify.py -q

# 3. Full compat-oracle pass count rises by exactly 1.
uv run --no-sync pytest tests/test_openpyxl_compat_oracle.py -q
# expect: 31 passed, 19 xfailed (was 30 / 20)

# 4. External-oracle preservation stays green when fixtures present.
uv run --no-sync pytest tests/test_external_oracle_preservation.py -q
```

After the probe passes, flip `images.replace_remove` (and any sibling rows tracked under G06) in `docs/migration/_compat_spec.py` from `partial`/`not_yet` (gap_id `G06`) to `supported`, drop `gap_id`, re-render the matrix, and mark G06 `landed`.

## 5. Out-of-scope guards

- **Do not** add `ws.replace_image` semantics that mutate the existing `Image` object in place. Replace = add new + remove old; keep object identity boundaries clean.
- **Do not** touch the Rust write-mode add path (`native_writer_backend.rs:416 add_image`). Modify mode is the only mode this handoff cares about.
- **Do not** redesign the `Image` class itself. The existing constructor + anchor model is fine.
- **Do not** sweep `_worksheet_media.py` for unrelated cleanup.
- **Do not** add a `clear_all_images()` shortcut. The probe + openpyxl API surface only need single-image removal.
- **Do not** introduce a separate "delete by anchor" lookup; index-based `remove_image(idx)` is the openpyxl-shaped contract. If users want anchor-based removal, that is a follow-up handoff.

## 6. Verification commands

```bash
uv run --no-sync maturin develop  # required — this handoff touches Rust

cargo test -p wolfxl --lib patcher_drawing
uv run --no-sync pytest -q --ignore=tests/test_external_oracle_preservation.py --ignore=tests/diffwriter
uv run --no-sync pytest tests/test_openpyxl_compat_oracle.py -q
uv run --no-sync python scripts/render_compat_matrix.py
```

Mark G06 `landed` only after all four acceptance commands pass and the matrix is regenerated.
