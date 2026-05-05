# Codex Handoff — G07 (residual): DataTableFormula round-trip

> **Sprint**: S1 — Bridge Completion
> **Branch**: `feat/parity-G07-data-table-formula`
> **Sibling pods**: G03 (diagonal borders), G04 (protection), G05 (NamedStyle), G06 (image replace)
> **Merge gate**: independent codepaths — no sibling-blocking
> **Note**: basic `ArrayFormula(ref, text)` was promoted to `supported` during S0 triage. This handoff closes the residual G07 case: `DataTableFormula`.

## 1. Goal

Make `cell.value = DataTableFormula(ref="C1:C3", t="dataTable", r1="A1", dt2D=False)` round-trip through `wb.save()` and `load_workbook()` so the `array_formula_data_table` probe in the openpyxl compat oracle stops xfailing.

The Python class, write emit, and read parse are all wired today. The probe xfails, so something in the chain breaks for the data-table specifically (most likely a missing dispatch case or an attribute drop). Diagnose, then fix the smallest surface.

## 2. Files to touch

| File | What to do |
|---|---|
| `python/wolfxl/cell/cell.py:78` | Confirm `DataTableFormula` constructor accepts `ref`, `ca`, `dt2D`, `dtr`, `r1`, `r2`. The probe at `tests/test_openpyxl_compat_oracle.py:1125` constructs it as `DataTableFormula(ref="C1:C3", t="dataTable", r1="A1", dt2D=False)` — the `t=` kwarg is openpyxl's convention; confirm the wolfxl class accepts or ignores it. |
| `python/wolfxl/_cell.py:400–520` | `_queue_data_table_formula()` puts the value into `ws._pending_array_formulas` with `kind="data_table"`. Confirm every probe attribute (ref, ca, dt2D, dtr, r1, r2) reaches the payload dict. |
| `crates/wolfxl-writer/src/emit/sheet_data.rs:172–202` | Emits `<f t="dataTable" ref="..." ca="1" dt2D="1" dtr="1" r1="..." r2="..."/></c>`. Confirm every attribute the Python payload provides is preserved on the wire. |
| `crates/wolfxl-reader/src/lib.rs:5087–5088, 5195–5202` | Read parser extracts `t="dataTable"` and constructs `ArrayFormulaInfo::DataTable{...}`. Confirm the dict it hands back to Python contains every attribute the probe needs. |
| `python/wolfxl/cell/cell.py` (or wherever `DataTableFormula` is reconstructed on load) | Confirm load_workbook reconstructs a `DataTableFormula` (not a string, not an `ArrayFormula`) from the read payload. This is the most-likely source of the xfail. |
| `tests/test_data_table_formula.py` (extend) | Add an end-to-end round-trip test that mirrors the probe but with explicit assertions on every attribute (`ref`, `r1`, `dt2D`, etc.), not just isinstance. |

## 3. Reference exemplar

`tests/test_array_formula.py:151–165` — basic `ArrayFormula` round-trip. Save, load, assert isinstance + field equality. The DataTable case should mirror this exactly. The S0 hardening of the basic-array probe (now reading via openpyxl as well as wolfxl) is the gold-standard contract; aim for the same coverage in the new test.

For the cell-value reconstruction step, the basic `ArrayFormula` reconstruction path is the template — find where the loader reads the array-formula dict and constructs an `ArrayFormula` object. The DataTable branch should slot in next to it.

## 4. Acceptance tests

```bash
# 1. The compat-oracle probe flips from xfail to pass.
uv run --no-sync pytest tests/test_openpyxl_compat_oracle.py -k array_formula_data_table -q
# expect: 1 passed (was xfail)

# 2. The basic ArrayFormula probe stays passing (no regression).
uv run --no-sync pytest tests/test_openpyxl_compat_oracle.py -k array_formula_basic -q
# expect: 1 passed

# 3. Existing data-table tests stay green.
uv run --no-sync pytest tests/test_data_table_formula.py -q

# 4. Full compat-oracle pass count rises by exactly 1.
uv run --no-sync pytest tests/test_openpyxl_compat_oracle.py -q
# expect: 31 passed, 19 xfailed (was 30 / 20)
```

After the probe passes, flip `array_formulas.data_table` in `docs/migration/_compat_spec.py` from `partial` (gap_id `G07`) to `supported`, drop `gap_id`. With `array_formulas.array_formula` already promoted in S0, both G07 rows now read `supported` → mark **G07 fully closed** in the tracker.

## 5. Out-of-scope guards

- **Do not** redesign the `_pending_array_formulas` storage shape — its current dict-with-`kind` discriminator is fine.
- **Do not** add support for shared-formula range expansion (`<f t="shared">`) in this handoff. That is a separate gap, not under G07.
- **Do not** change `ArrayFormula` semantics. G07's basic case is already green; this handoff must not regress it.
- **Do not** rename PyO3 entry points. The existing dispatch is correct.
- **Do not** sweep unrelated cell-value parsing cleanup; keep the diff focused on the data-table case.

## 6. Verification commands

```bash
uv run --no-sync maturin develop  # only if any .rs file changed

cargo test -p wolfxl-writer sheet_data
uv run --no-sync pytest -q --ignore=tests/test_external_oracle_preservation.py --ignore=tests/diffwriter
uv run --no-sync pytest tests/test_openpyxl_compat_oracle.py -q
uv run --no-sync python scripts/render_compat_matrix.py
```

Mark G07 `landed` only after the probe passes, the existing array-formula probe still passes, and the matrix shows both array-formula rows as `supported`.
