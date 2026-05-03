# Codex Handoff — G04 (residual): Protection Python flow

> **Sprint**: S1 — Bridge Completion
> **Branch**: `feat/parity-G04-protection`
> **Sibling pods**: G03 (diagonal borders), G05 (NamedStyle), G06 (image replace), G07 (DataTableFormula)
> **Merge gate**: independent codepaths — no sibling-blocking
> **Note**: the *sheet* half of G04 was promoted to `supported` during S0 triage. This handoff closes the two remaining halves: **workbook protection** and **cell-level protection setter**.

## 1. Goal

Two independent deliverables, one branch, both gated under G04:

- **(a)** Make `wb.security = WorkbookProtection(lockStructure=True, workbookPassword="secret")` (openpyxl camelCase) round-trip through `wb.save()` and `load_workbook()`. The compat oracle probe is `protection_workbook` (currently xfail).
- **(b)** Make `cell.protection = Protection(locked=False, hidden=True)` actually settable. The getter exists; the setter is missing. The compat oracle probe is `cell_protection` (currently xfail).

Both round trips already work end-to-end through the Rust write/read backends. The breaks are Python-class API: a missing kwarg-alias map for (a) and a missing setter for (b).

## 2. Files to touch

**Deliverable (a) — workbook protection camelCase:**

| File | What to do |
|---|---|
| `python/wolfxl/workbook/protection.py:36–115` | Add camelCase kwargs to `WorkbookProtection.__init__` (or accept `**kwargs` and map). Add property aliases (`@property lockStructure`, etc.) for every attribute openpyxl exposes. Mirror `SheetProtection`. |
| `tests/test_workbook_protection.py` (extend) | Unit test: snake_case and camelCase kwargs construct equivalent objects; `getattr(prot, "lockStructure")` matches `prot.lock_structure`. |

Do **not** touch the workbook-level wiring (`python/wolfxl/_workbook.py:611–632`), the patcher flush (`python/wolfxl/_workbook_patcher_flush.py:254–275`), the Rust emit (`crates/wolfxl-writer/src/parse/workbook_security.rs`), or the reader (`crates/wolfxl-reader/src/lib.rs:948–1530`). All four are correctly wired; the only break is the constructor signature.

**Deliverable (b) — cell.protection setter:**

| File | What to do |
|---|---|
| `python/wolfxl/_cell.py` | Find the `Cell.protection` getter; add a matching setter that accepts an openpyxl-shaped `Protection` object and stores it the same way `cell.font` / `cell.fill` setters do today (the existing single-attribute pattern collapses to last-setter only — this handoff just needs `cell.protection = Protection(...)` to round-trip on its own, *not* combined with other style setters; the combined-style case is G05). |
| `tests/test_cell_protection_setter.py` (NEW) | Focused unit test that constructs a cell, assigns `Protection(locked=False, hidden=True)`, saves, reloads, and asserts both attributes survive. |

Do **not** redesign the style-merge path. The setter just needs to make a single `cell.protection = Protection(...)` call work; it does not need to compose with `cell.font = ...` and `cell.fill = ...` on the same cell. That is G05's job.

## 3. Reference exemplar

For deliverable (a): `python/wolfxl/worksheet/protection.py:17–50` — `SheetProtection`. snake_case dataclass with camelCase property aliases that read/write the canonical snake_case fields. The S0 triage proved this pattern round-trips through openpyxl's reader.

For deliverable (b): the `cell.font` and `cell.fill` setters already exist on `Cell`. Find the simplest of those two and mirror the pattern for `protection`. (The combined-style preservation issue noted in the spec for `cell.font_fill_border_alignment` is **out of scope** here.)

## 4. Acceptance tests

```bash
# 1. Both compat-oracle probes flip from xfail to pass.
uv run --no-sync pytest tests/test_openpyxl_compat_oracle.py -k 'protection_workbook or cell_protection' -q
# expect: 2 passed (was 2 xfail)

# 2. Existing protection unit and parity tests stay green.
uv run --no-sync pytest tests/test_workbook_protection.py tests/parity/test_workbook_security_parity.py -q

# 3. The new cell-protection setter test passes.
uv run --no-sync pytest tests/test_cell_protection_setter.py -q

# 4. Full compat-oracle pass count rises by exactly 2.
uv run --no-sync pytest tests/test_openpyxl_compat_oracle.py -q
# expect: 32 passed, 18 xfailed (was 30 / 20)
```

After both probes pass, flip `protection.workbook` and `cell.protection` in `docs/migration/_compat_spec.py` from `partial` (gap_id `G04`) to `supported`, drop `gap_id` on each, re-render the matrix. With G04's sheet half already promoted in S0 triage, all three protection rows now read `supported` → mark **G04 fully closed** in the tracker.

## 5. Out-of-scope guards

- **Do not** rename or remove the snake_case canonical fields on `WorkbookProtection`. Both naming styles must continue to work.
- **Do not** change the Rust struct (`WorkbookProtectionSpec`) field names — Python is the only layer that needs alias coverage.
- **Do not** alter the encryption / hash-and-salt path. The `secret` password in the probe is already correctly hashed by existing emit code.
- **Do not** redesign the per-cell style-merge path. The `cell.protection` setter just needs to work in isolation; combined-style preservation is G05.
- **Do not** sweep in `WorkbookProtection` API additions (e.g. `revisionsPassword`) that aren't referenced by the probe — separate ratchet, separate handoff.

## 6. Verification commands

```bash
uv run --no-sync maturin develop  # only if any .rs file changed (this handoff should not need it)

uv run --no-sync pytest -q --ignore=tests/test_external_oracle_preservation.py --ignore=tests/diffwriter
uv run --no-sync pytest tests/test_openpyxl_compat_oracle.py -q
uv run --no-sync python scripts/render_compat_matrix.py
```

Mark G04 `landed` only after all of the above pass and the matrix is regenerated and committed.
