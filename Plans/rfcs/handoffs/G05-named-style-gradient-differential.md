# Codex Handoff — G05: NamedStyle / GradientFill / DifferentialStyle Python flow

> **Sprint**: S1 — Bridge Completion
> **Branch**: `feat/parity-G05-named-style-bridge`
> **Sibling pods**: G03 (diagonal borders), G04 (protection), G06 (image replace), G07 (DataTableFormula)
> **Merge gate**: independent codepaths — no sibling-blocking
> **Background RFC**: [`Plans/rfcs/064-styles-namedstyle-gradient.md`](../064-styles-namedstyle-gradient.md) — read this first; it has the design contract for all three classes. This handoff closes the residual Python-side bridge work that the RFC's implementation didn't fully reach.

## 1. Goal

Make three openpyxl-style style-table idioms round-trip through `wb.save()` and `load_workbook()` so the corresponding compat-oracle probes stop xfailing:

- **NamedStyle**: `wb.add_named_style(NamedStyle(name="Highlight", font=Font(bold=True)))` + `cell.style = "Highlight"` round-trips so `wb2.active["A1"].style == "Highlight"`. (probe: `cell_named_style`)
- **GradientFill**: `cell.fill = GradientFill(type="linear", degree=90, stop=(Color(...), Color(...)))` round-trips so the reloaded fill's `type` is `"linear"` or `"gradient"`. (probe: `cell_gradient_fill`)
- **Combined styles preservation**: setting `cell.font`, `cell.fill`, `cell.alignment`, `cell.border`, and `cell.number_format` all on one cell preserves all five through save+reload. Today only the last setter wins. (probe: `cell_font_fill_border_alignment`)

DifferentialStyle is referenced in RFC-064 as the dxf-table emitter; it is exercised by S3 conditional-formatting probes, not by S1 directly. Land DifferentialStyle's Python class + emit only if it falls naturally out of the same change; do not write a dedicated probe in this handoff.

## 2. Files to touch

| File | What to do |
|---|---|
| `python/wolfxl/styles/_named_style.py` | Confirm `NamedStyle` matches the dataclass shape in RFC-064 §2.1. Make sure `wb.add_named_style(style)` registers it on `wb.named_styles` (registry already at `_workbook.py:427`). |
| `python/wolfxl/_workbook.py:385`, `:427`, `:549` | The named-style registry getter, lazy-init, and `add_named_style` already exist. Wire the registry through to the patcher / writer flush so `<cellStyle>` and `<cellStyleXfs>` are emitted when `wb.save()` runs. |
| `python/wolfxl/_cell.py` | `cell.style = "Highlight"` setter must look up the registry on the parent workbook and resolve to a style xf id at flush time. (Note from RFC-064 §3.) |
| `python/wolfxl/_cell.py` (combined-style fix) | The current single-attribute setters collapse to last-setter only. Rewrite the per-cell style flush so font / fill / border / alignment / number_format all coexist on the same cell. |
| `python/wolfxl/styles/fills.py` | Confirm `GradientFill` constructor matches the openpyxl shape (`type`, `degree`, `stop`, `top`, `bottom`, `left`, `right`). |
| `python/wolfxl/_cell_payloads.py` (or equivalent fill bridge) | Extend `fill_to_rust_dict()` (or analogue) to recognise GradientFill and emit a payload the Rust writer's existing gradient emitter can consume. RFC-064 §4 has the dict shape. |
| `crates/wolfxl-writer/src/emit/styles_xml.rs` | If `<gradientFill>` emit is incomplete (RFC-064 §4 spec), finish it. If complete, leave alone. |
| `tests/test_named_style_bridge.py` (NEW) | Focused unit test: register a NamedStyle, assign it to a cell, save+reload, assert the lookup survives. |
| `tests/test_combined_cell_styles.py` (NEW) | Focused unit test for the font/fill/border/alignment/number_format combined-preservation case. This test also guards the broader G05 contract — combined style setting must not collapse to last-setter. |

## 3. Reference exemplar

The single-attribute setters that already round-trip in isolation (`cell.font`, `cell.alignment`, `cell.border`, `cell.number_format`) are the per-attribute templates. The bug is that each one's flush replaces the cell's xf id rather than merging into it. The combined-style test proves both directions: set five attrs, expect five attrs preserved.

For the NamedStyle registry-lookup pattern, see how `Cell.number_format` resolves through `wb._number_formats` (or equivalent named registry). The named-style xf id allocation in `<cellStyleXfs>` is described in RFC-064 §3.

## 4. Acceptance tests

```bash
# 1. The three compat-oracle probes flip from xfail to pass.
uv run --no-sync pytest tests/test_openpyxl_compat_oracle.py \
  -k 'cell_named_style or cell_gradient_fill or cell_font_fill_border_alignment' -q
# expect: 3 passed (was 3 xfail)

# 2. New focused tests pass.
uv run --no-sync pytest tests/test_named_style_bridge.py tests/test_combined_cell_styles.py -q

# 3. Existing style suite stays green.
uv run --no-sync pytest tests/ -k 'styles or font or fill or border' -q

# 4. Full compat-oracle pass count rises by exactly 3.
uv run --no-sync pytest tests/test_openpyxl_compat_oracle.py -q
# expect: 33 passed, 17 xfailed (was 30 / 20)
```

After all three probes pass, flip `cell.named_style`, `cell.gradient_fill`, and `cell.font_fill_border_alignment` in `docs/migration/_compat_spec.py` from `partial` (gap_id `G05`) to `supported`, drop `gap_id` on each, re-render the matrix, and mark G05 `landed` in the tracker.

## 5. Out-of-scope guards

- **Do not** rewrite the entire styles emitter. RFC-064 already designed the structure; this handoff fills in the Python bridge gaps, not the OOXML semantics.
- **Do not** add support for Border-on-the-named-style or alignment-on-the-named-style edge cases beyond what the probes test. Those are S3-adjacent (CF + dxf table integration).
- **Do not** introduce a DifferentialStyle public probe in this handoff. That class is exercised by S3 (conditional formats); the read/write plumbing is in scope here only insofar as it supports the three probes above.
- **Do not** sweep cleanup of unrelated `_styles.py` shapes. Keep the diff focused on the three failing probes.
- **Do not** change the public API shape of `Font`, `PatternFill`, `Alignment`, `Border` (they are partly green today). Adding new attributes is acceptable; renaming or removing is not.

## 6. Verification commands

```bash
uv run --no-sync maturin develop  # only if any .rs file changed

uv run --no-sync pytest -q --ignore=tests/test_external_oracle_preservation.py --ignore=tests/diffwriter
uv run --no-sync pytest tests/test_openpyxl_compat_oracle.py -q
uv run --no-sync python scripts/render_compat_matrix.py
```

Mark G05 `landed` only after all three probes pass, RFC-064's design constraints are upheld, and the matrix reflects three new `supported` rows.
