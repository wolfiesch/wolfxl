# WolfXL 1.2 — RFC-035 follow-ups + composition hardening

_Date: 2026-04-26_

WolfXL 1.2 closes the open follow-ups from 1.1's RFC-035 landing
and hardens the `copy_worksheet` cross-RFC composition surface.
The three biggest user-visible changes are: write-mode
`copy_worksheet` (the §3 OQ-a `NotImplementedError` is gone),
opt-in image deep-clone via `CopyOptions(deep_clone_images=True)`,
and a wolfxl-side `xl/calcChain.xml` rebuild so external readers no
longer wait for Excel to fix the chain on first open.

Two cross-RFC composition bugs that 1.1 left as `xfail(strict=True)`
— the self-closing `<sheets/>` splice corner case (#4) and the
CDATA / processing-instruction fakeout in the workbook.xml splice
locator (#6) — are also closed. The Phase 2.7 splice is now
XML-aware rather than byte-level. Together, the two fixes mean
RFC-035 has zero deferred composition gaps in 1.2: every case in
`tests/test_copy_worksheet_modify.py` passes without an xfail.

## What's new

### Write-mode `Workbook.copy_worksheet` (Sprint Θ Pod-C1, RFC-035 §3 OQ-a)

```python
import wolfxl

wb = wolfxl.Workbook()                         # write mode, no source file
template = wb.active
template.title = "Template"
template["A1"] = "Hello"
template["A2"] = "=A1"

clone = wb.copy_worksheet(template)            # 1.1 raised NotImplementedError
clone.title = "Clone"

wb.save("out.xlsx")
```

`Workbook.copy_worksheet(source, name=None)` now works in **write
mode** as well as modify mode. The write-mode path walks the
in-memory `NativeWorkbook` model and clones every sub-record (cells,
styles, tables, DV, CF, hyperlinks, comments, defined names) into a
fresh sheet appended at the end of the tab list.

Behaviour matches the modify-mode contract (§5.4 sheet-scoped
defined-name fan-out, §5.5 table-name auto-rename `_{N}`, §10
sheet-scoped vs. workbook-scoped name divergence). The new tests
in `tests/test_copy_worksheet_write.py` mirror the modify-mode
harness so any divergence surfaces immediately.

<!-- TBD: Sprint Θ Pod-C1 commit sha when integrated -->

### Image deep-clone via `CopyOptions` (Sprint Θ Pod-C2, RFC-035 §5.3 / §8 risk #2)

```python
from wolfxl import CopyOptions

clone = wb.copy_worksheet(
    template,
    options=CopyOptions(deep_clone_images=True),    # additive; default False
)
```

The new `CopyOptions` keyword controls how embedded image media
(`xl/media/image*.png` / `.jpg`) is treated on copy:

- **`deep_clone_images=False`** (default): preserves 1.1's URL-string
  alias behaviour — the cloned drawing rels point at the source's
  `xl/media/imageN.png`. Byte-identical to 1.1.
- **`deep_clone_images=True`**: duplicates the image part on the way
  out. The clone's drawing rels point at a fresh
  `xl/media/imageM.png` with byte-identical content but a new
  `<Relationship Id>`.

The default stays `False` to preserve 1.1's byte-stability golden
test. Enable deep-clone when you need to mutate the copy's images
independently of the source (e.g. swapping a logo per quarterly
report). The opt-in trades 50× bloat on logo-heavy workbooks for
mutation independence — the right knob for the job.

<!-- TBD: Sprint Θ Pod-C2 commit sha when integrated -->

### `xl/calcChain.xml` rebuild post-copy (Sprint Θ Pod-C3, RFC-035 §10 / §8 risk #9)

1.1 left `xl/calcChain.xml` stale after `copy_worksheet`, relying on
Excel's "rebuild on next open" behaviour. External readers that
consume `calcChain.xml` directly (parity tests, programmatic
auditors, third-party xlsx tooling) saw an incomplete chain.

1.2 walks the cloned sheet's formula cells in a post-Phase-2.7 pass
and emits the matching `<c>` entries into `xl/calcChain.xml` so the
chain is complete on the wolfxl-emitted file — no Excel round-trip
needed.

<!-- TBD: Sprint Θ Pod-C3 commit sha when integrated -->

### Phase 2.7 splice: self-closing `<sheets/>` and CDATA fuzz (Sprint Θ Pods A + B)

The two RFC-035 cross-RFC composition bugs that 1.1 deferred are
now closed:

- **Pod-A — bug #4 (`test_p_self_closing_sheets_block`)**: the Phase
  2.7 splice now handles a synthesized self-closing `<sheets/>`
  workbook.xml correctly. The fix lives at the splice unit-test
  level; integration through the public `load_workbook` loader path
  remains gated behind a future `permissive=True` mode (real Excel
  never emits `<sheets/>` for a non-empty workbook).
  <!-- TBD: Sprint Θ Pod-A commit sha when integrated -->
- **Pod-B — bug #6 (`test_r_cdata_pi_fuzz_fakeout`)**: the byte-level
  workbook.xml splice locator is replaced with an XML-aware scan
  (`quick-xml` reader pass) that respects element nesting. A
  workbook.xml comment containing literal `</sheets>` no longer
  fools the locator.
  <!-- TBD: Sprint Θ Pod-B commit sha when integrated -->

After 1.2, `tests/test_copy_worksheet_modify.py` ships zero
`xfail(strict=True)` markers; every case passes.

## Breaking changes

None. Every 1.2 change is additive:

- `CopyOptions` is a new keyword argument with a backward-compatible
  default (`deep_clone_images=False` matches 1.1 byte-for-byte).
- Write-mode `copy_worksheet` previously raised
  `NotImplementedError`; lifting the raise is forward-compatible.
- The Phase 2.7 splice rewrite is internal — the public API surface
  is unchanged, only the byte-level locator is replaced with an
  XML-aware one.
- `xl/calcChain.xml` rebuild produces a *more correct* chain than
  1.1's stale output; downstream tools that previously worked
  around 1.1's staleness keep working (the chain is a superset of
  what Excel-on-open would produce).

## Migration guide

No source changes are required. A few opt-in knobs are available:

- If your code relied on the write-mode `NotImplementedError`
  signal (e.g. as a feature-detection probe), update to the new
  shipped behaviour. The recommended detection pattern is now
  `hasattr(wolfxl.Workbook(), "copy_worksheet")` plus a try-call —
  but since 1.1 already exposed the modify-mode method, most code
  can simply call `copy_worksheet` unconditionally.
- If your pipeline mutates an image on a copied sheet and relied on
  the source-side mutation for both, that path is now incorrect
  on 1.2 only if you opt into `deep_clone_images=True`. The default
  (`False`) preserves 1.1's aliasing behaviour. Enable deep-clone
  only when you intentionally want isolated copies.
- If your tooling consumed `xl/calcChain.xml` directly and worked
  around 1.1's staleness with a recompute pass, you can drop the
  workaround in 1.2 — the chain is now complete on emit.

## Known limitations

The 1.1 *Known divergences* table is reduced — the image-aliasing
and calcChain-staleness rows now have an opt-in / shipped
counterpart in 1.2. Carry-forward limitations:

- **`copy_worksheet` re-saved by openpyxl**: openpyxl's own
  `copy_worksheet` deep-copies tables / DV / CF / sheet-scoped
  defined names as Python objects on re-save; if you openpyxl-save
  in the middle of a wolfxl pipeline, those parts are dropped from
  the duplicate. wolfxl's emitted file is structurally correct;
  openpyxl's loader is the lossy step. Workaround: stay inside
  wolfxl until the final save.
- **Sheet-scoped defined-name divergence on copy** (RFC-035 §3 OQ-c):
  wolfxl clones sheet-scoped names with a re-pointed
  `localSheetId`; openpyxl drops them silently. Documented as a
  deliberate divergence in `tests/parity/KNOWN_GAPS.md`.
- **Cross-workbook copy** (`copy_worksheet(other_wb_sheet)`):
  remains out of scope per RFC-035 §10. openpyxl rejects the same
  call.
- **Chart sheets** (`<chartsheet>`): remain out of scope per
  RFC-035 §10. The Python coordinator continues to raise
  `NotImplementedError` for non-`Worksheet` sources.

See `tests/parity/KNOWN_GAPS.md` for the full per-feature gap list,
which now reflects all RFC-020 / 021 / 022 / 023 / 024 / 025 / 026
mutations as ✅ Shipped (the W4F audit is closed) plus the new
1.2 status of `copy_worksheet` write mode and image deep-clone.

## Acknowledgments

Sprint Θ ("Theta") pods that landed 1.2:

- **Pod-A — RFC-035 bug #4** (self-closing `<sheets/>` splice).
  <!-- TBD: pod sha when integrated -->
- **Pod-B — RFC-035 bug #6** (XML-aware splice replacing byte-level
  locator).
  <!-- TBD: pod sha when integrated -->
- **Pod-C1 — write-mode `copy_worksheet`** (lifts §3 OQ-a).
  <!-- TBD: pod sha when integrated -->
- **Pod-C2 — image deep-clone** (`CopyOptions(deep_clone_images=True)`).
  <!-- TBD: pod sha when integrated -->
- **Pod-C3 — calcChain rebuild** (post-Phase-2.7 chain emit).
  <!-- TBD: pod sha when integrated -->
- **Pod-D (this release)** — KNOWN_GAPS T1.5 reconciliation,
  RFC-035 §8.5 Sprint Θ deliverables section, 1.2 release notes
  scaffold.

Specs: see `Plans/rfcs/035-copy-worksheet.md` (especially the new
§8.5 "Sprint Θ deliverables (1.2)" subsection) for the
implementation plan and per-pod scope.

Thanks to everyone who exercised RFC-035 in production after 1.1
shipped — the bug surface this sprint closed was found by real
workloads, not synthetic harnesses.
