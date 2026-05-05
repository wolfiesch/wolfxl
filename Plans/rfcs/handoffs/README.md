# Codex Handoffs — Openpyxl Parity Program

Each file in this directory is a self-contained Codex handoff spec for one parity-program gap. The format follows the openpyxl parity program plan ([`Plans/openpyxl-parity-program.md`](../../openpyxl-parity-program.md), §"Codex delegation strategy").

## Contract

A handoff is **only ready** to dispatch when it contains all six fields:

1. **Goal** — one sentence describing the user-facing capability.
2. **Files to touch** — explicit paths.
3. **Reference exemplar** — a pointer to an analogous, already-shipped feature.
4. **Acceptance tests** — specific test names and the compat-oracle slice.
5. **Out-of-scope guards** — explicit list of what the pod must not touch.
6. **Verification commands** — the exact gate the pod runs before opening a PR.

## Active handoffs (Sprint 1)

| Gap | Handoff | Branch | Status |
|---|---|---|---|
| G03 | [`G03-diagonal-borders.md`](G03-diagonal-borders.md) | `feat/parity-G03-diagonal-borders` | ready |
| G04 | [`G04-workbook-protection.md`](G04-workbook-protection.md) | `feat/parity-G04-protection` | ready |
| G05 | [`G05-named-style-gradient-differential.md`](G05-named-style-gradient-differential.md) | `feat/parity-G05-named-style-bridge` | ready |
| G06 | [`G06-image-replace-remove.md`](G06-image-replace-remove.md) | `feat/parity-G06-image-replace-remove` | ready |
| G07 | [`G07-data-table-formula.md`](G07-data-table-formula.md) | `feat/parity-G07-data-table-formula` | ready |

## Dispatch protocol

For an N-way parallel sprint:

1. Open a separate git worktree per handoff (`git worktree add ../wolfxl-G03 feat/parity-G03-diagonal-borders` etc.).
2. Spawn one Codex pod per worktree from the same Claude session (or sequentially if you only have one Codex seat).
3. Each pod runs the verification commands listed in its handoff before opening a PR.
4. Merge serially — the parity-program tracker shows landing order; flip the `Status` column from `ready` → `in-progress` → `review` → `landed` per row.
5. After each merge, regenerate the compat matrix (`python scripts/render_compat_matrix.py`) and verify the oracle pass count rose by the expected number.
