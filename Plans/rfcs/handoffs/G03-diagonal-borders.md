# Codex Handoff — G03: Diagonal borders Python bridge

> **Sprint**: S1 — Bridge Completion
> **Branch**: `feat/parity-G03-diagonal-borders`
> **Sibling pods**: G04 (workbook protection), G05 (NamedStyle), G06 (image replace), G07 (DataTableFormula)
> **Merge gate**: independent codepaths — no sibling-blocking

## 1. Goal

Plumb `Border(diagonal=Side(...), diagonalUp=True, diagonalDown=True)` from the Python style API end-to-end through both write mode and modify mode so the diagonal-border probe in the openpyxl compat oracle stops xfailing.

The Rust write emit already supports diagonal borders; the gap is two missing extractors on the bridge between Python and Rust.

## 2. Files to touch

| File | What to do |
|---|---|
| `python/wolfxl/_cell_payloads.py:168–181` | Extend `border_to_rust_dict()` to read `border.diagonal` (`Side`), `border.diagonalUp` (`bool`), `border.diagonalDown` (`bool`) and include them in the dict. |
| `src/wolfxl/patcher_payload.rs:78–99` | Extend `dict_to_border_spec()` to extract the same three fields and populate `BorderSpec.diagonal`, `BorderSpec.diagonal_up`, `BorderSpec.diagonal_down`. |
| `tests/test_diagonal_borders_python_bridge.py` (NEW) | Add a focused write-mode + modify-mode round-trip test for diagonal borders. |

Do **not** touch the Rust write emit at `crates/wolfxl-writer/src/emit/styles_xml.rs:201–220` — it already emits `diagonalUp`/`diagonalDown` attrs and the `<diagonal>` element correctly. The two existing emit-side tests (`border_diagonal_style_and_color_emit`, `border_diagonal_up_and_down_attrs_present`, in the same file at lines 811–851) are the spec for what the dict shape must produce.

## 3. Reference exemplar

The `top` border side is the template. Trace it end-to-end:

1. Python: `border.top` is a `Side`; `border_to_rust_dict()` writes `{"top": side_to_dict(border.top)}`.
2. Modify-mode: `dict_to_border_spec()` calls `extract_side("top", &dict)` → `BorderSideSpec`.
3. Write-mode: `BorderSpec.top: Option<BorderSideSpec>` is consumed by `border_side_xml("top", &spec.top)` in `styles_xml.rs`.

Mirror this exactly for `diagonal`. The two boolean attributes (`diagonalUp`, `diagonalDown`) are simpler — they do not go through `extract_side`; they are just `bool` extracted from the dict.

## 4. Acceptance tests

The handoff is done when **all** of the following pass:

```bash
# 1. The compat-oracle probe flips from xfail to pass.
uv run --no-sync pytest tests/test_openpyxl_compat_oracle.py -k cell_diagonal_borders -q
# expect: 1 passed (was xfail)

# 2. Existing emit tests stay green.
cargo test -p wolfxl-writer border_diagonal

# 3. The new bridge test passes (write mode + modify mode).
uv run --no-sync pytest tests/test_diagonal_borders_python_bridge.py -q

# 4. Compat-oracle full run shows pass count rises by exactly 1, with no regressions.
uv run --no-sync pytest tests/test_openpyxl_compat_oracle.py -q
# expect: 31 passed, 19 xfailed (was 30 / 20)
```

After the probe passes, flip the spec entry in `docs/migration/_compat_spec.py` for `cell.diagonal_borders` from `partial` (gap_id `G03`) to `supported`, drop `gap_id`, and re-render via `python scripts/render_compat_matrix.py`.

## 5. Out-of-scope guards

- **Do not** change the Rust emit signature (`border_to_xml`, `border_side_xml`).
- **Do not** refactor `BorderSpec` or `BorderSideSpec` shapes.
- **Do not** touch the read-side (`crates/wolfxl-reader/`) — diagonals already round-trip on read; only the write/modify bridge is broken.
- **Do not** add diagonal-related fields to the Cell-level shortcut API; this is a `Border` extension only.
- **Do not** sweep in unrelated `_cell_payloads.py` cleanups.

## 6. Verification commands

```bash
# rebuild bindings (only needed if a .rs file changed)
uv run --no-sync maturin develop

# Rust unit tests
cargo test -p wolfxl-writer

# Python suite (excluding fixture-gated slices)
uv run --no-sync pytest -q --ignore=tests/test_external_oracle_preservation.py --ignore=tests/diffwriter

# compat-oracle (the program-level metric)
uv run --no-sync pytest tests/test_openpyxl_compat_oracle.py -q

# regenerate matrix after spec flip
uv run --no-sync python scripts/render_compat_matrix.py
```

Mark the G03 row in `Plans/openpyxl-parity-program.md` `landed` only after all four acceptance commands pass clean and the matrix has been regenerated and committed.
