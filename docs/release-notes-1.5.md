# wolfxl 1.5.0 (2026-04-26) — encrypted writes + image construction + streaming-datetime fix

_Date: 2026-04-26_

WolfXL 1.5 lifts the last two "construction" gaps that the 1.0–1.4
arc deferred as out-of-scope: write-side OOXML encryption (Pod-α,
RFC-044) and image construction (Pod-β, RFC-045). It also closes the
streaming-reads datetime divergence Pod-β surfaced in 1.3 (Pod-γ).
After Sprint Λ ("Lambda") the openpyxl-parity surface is exhausted at
the construction level too — only chart construction (deferred to
v1.6.0 / Sprint Μ) and pivot table construction (deferred to v2.0.0
/ Sprint Ν) remain.

## TL;DR

- **Encrypted writes** — `wb.save(path, password="...")` now emits an
  Agile (AES-256 / SHA-512) encrypted xlsx via `msoffcrypto-tool`.
  Install via `pip install wolfxl[encrypted]`. Closes the
  `NotImplementedError` at `python/wolfxl/_workbook.py:1032`.
- **Image construction** — `Image(...)` and `ws.add_image(...)` are
  real. PNG / JPEG / GIF / BMP. One-cell, two-cell, and absolute
  anchors. Works in both write mode and modify mode.
- **Streaming-datetime fix** — `iter_rows(values_only=True)` now
  returns `datetime` objects for date-formatted cells under
  `read_only=True`, matching openpyxl. Previously surfaced as Excel
  serial floats.

## What's new

### Encrypted writes (Sprint Λ Pod-α, RFC-044)

```python
import wolfxl

# Write an encrypted file from scratch
wb = wolfxl.Workbook()
wb.active["A1"] = "secret data"
wb.save("budget.xlsx", password="hunter2")     # NEW

# Round-trip an encrypted file (read → mutate → write)
wb = wolfxl.load_workbook("budget.xlsx", password="hunter2", modify=True)
wb.active["A1"] = "edited"
wb.save("budget.xlsx", password="hunter2")     # NEW — re-encrypted in place
```

`wb.save(path, password="...")` flushes the workbook to plaintext
bytes via the existing writer / patcher pipeline, then re-encrypts
those bytes via `msoffcrypto-tool`'s high-level `OOXMLFile.encrypt()`
helper. The encrypted file is written atomically (tempfile +
`os.replace`) so an interrupted save can never leave the target path
in a half-written state.

**Algorithm scope**: Agile (AES-256 / SHA-512) only. `msoffcrypto-tool`
is decrypt-only on the Standard (AES-128, Office 2007) and XOR /
RC4 / 40-bit (legacy) paths; cloning that work into wolfxl is
deferred until a customer specifically asks for it. Agile is the
Office 2010+ default and is what every modern Excel emits.

**Optional dependency**: `pip install wolfxl[encrypted]` adds
`msoffcrypto-tool >= 5.4`. Saving with `password=...` without the
extra installed raises `RuntimeError` with the install hint
(matching Sprint Ι Pod-γ's read-side error). The dep is lazy-imported
— users who never pass `password=` never pay the import cost.

**Empty-string password**: `password=""` is **not** equivalent to
`password=None`. Empty-string produces a real encryption envelope
with a literal empty key (matches the read-side semantics RFC-042
already documents).

Pod-α feat commit: `4bc806c` (merged via `738656a`). RFC:
`Plans/rfcs/044-encryption-writes.md`.

### Image construction (Sprint Λ Pod-β, RFC-045)

```python
from wolfxl.drawing.image import Image
from wolfxl.drawing.spreadsheet_drawing import (
    TwoCellAnchor, AbsoluteAnchor, AnchorMarker,
)
from wolfxl.drawing.xdr import XDRPoint2D, XDRPositiveSize2D
import wolfxl

wb = wolfxl.Workbook()
ws = wb.active

# One-cell anchor (default — pin to the top-left of B2)
img = Image("logo.png")
ws.add_image(img, "B2")

# Two-cell anchor (resize with both corners)
ws.add_image(img, anchor=TwoCellAnchor(
    _from=AnchorMarker(col=1, colOff=0, row=1, rowOff=0),
    to=AnchorMarker(col=5, colOff=0, row=10, rowOff=0),
))

# Absolute anchor (floating, EMU coordinates)
ws.add_image(img, anchor=AbsoluteAnchor(
    pos=XDRPoint2D(x=914400, y=914400),               # 1 inch from top-left
    ext=XDRPositiveSize2D(cx=2438400, cy=914400),     # 2.67" × 1"
))

wb.save("with_logo.xlsx")
```

`Image(filename)` accepts a path, `BytesIO`, or raw `bytes`. The
constructor sniffs the format from magic bytes and extracts width
and height from the format-specific header (PNG IHDR, JPEG SOF,
GIF LSD, BMP DIB), so callers don't have to specify dimensions.

**Format support**: PNG, JPEG, GIF, BMP. WebP, TIFF, SVG, EMF, WMF
are out of scope; the four supported formats cover ~99% of
real-world spreadsheet images.

**Anchor types**: one-cell (`<xdr:oneCellAnchor>`), two-cell
(`<xdr:twoCellAnchor>`), absolute (`<xdr:absoluteAnchor>`). String
anchors (`ws.add_image(img, "B2")`) are coordinate shorthand; the
default is one-cell pinned to the top-left of the named cell at the
image's intrinsic dimensions.

**Cross-mode**: works in both write mode (native writer emits
`drawingN.xml` + media + rels through the new images-emit pass) and
modify mode (patcher's new Phase 2.5j drains queued images and
routes them through `file_adds`). Composes with RFC-035
`copy_worksheet`: copying a sheet that has an image preserves the
aliased original; adding a new image to the copy works.

Pod-β feat commit: `d9cb569` (Image+add_image, merged via `7dc00d2`). RFC:
`Plans/rfcs/045-image-construction.md`.

### Streaming-datetime fix (Sprint Λ Pod-γ)

```python
import wolfxl

wb = wolfxl.load_workbook("daily_log.xlsx", read_only=True)
for row in wb.active.iter_rows(values_only=True, max_row=5):
    print(row)
# Before 1.5: (45382.0, 'sale', 99.50)         ← Excel serial float
# After 1.5:  (datetime(2024, 3, 15, ...), 'sale', 99.50)   ← real datetime
```

In 1.3 Pod-β, the streaming reader's `values_only` path returned
Excel serial floats for date-formatted cells, while the eager path
returned proper `datetime` objects. The divergence was documented as
a "Phase 4 known divergence" in `tests/parity/KNOWN_GAPS.md` lines
116-122 and tracked under Phase 3 rich-text follow-ups.

Pod-γ closes the gap by teaching the streaming reader to consult the
styles table for the cell's number format and convert serial floats
inline. Both `values_only=True` and `StreamingCell.value` now match
openpyxl's `read_only=True` behavior.

Pod-γ fix commit: `98cd147` (merged via `974b9b5`).

## Migration notes

### Encryption (RFC-044)

* **Install the extra**: `pip install wolfxl[encrypted]`. The wheel
  is unchanged for users who don't need encryption — `msoffcrypto-tool`
  is an optional dep (matches the Sprint Ι Pod-γ read-side
  pattern).
* **`save(password=...)` was previously a `NotImplementedError`**.
  Code that defensively caught the exception (e.g. wrapped saves
  in try/except to fall through to a different code path) now
  succeeds and writes an encrypted file. Audit any `try`/`except`
  around `wb.save(..., password=...)` calls and remove the fallback
  branch.
* **Algorithm**: Agile (AES-256) only. If you need Standard
  (Office 2007) for interop with very old Excel installs, the save
  raises with a pointer at the RFC-044 §3 algorithm-scope
  rationale. File a ticket with the workload requirement.
* **Round-trip flows that were forced through an explicit re-encrypt
  step** (load encrypted → wolfxl → save plaintext → manual
  msoffcrypto encrypt) can simplify to a single `save(password=...)`
  call. Plaintext never touches disk.

### Images (RFC-045)

* **`Image()` previously raised `NotImplementedError`**. Code that
  caught the exception and fell through to "render the image as a
  comment" or "skip the image" branches now succeeds — make sure
  callers handle the success case. Search for `_make_stub` /
  `Image(` in your codebase and audit error paths.
* **Anchor classes live under `wolfxl.drawing.spreadsheet_drawing`
  and `wolfxl.drawing.xdr`** to mirror openpyxl's module layout.
  Existing openpyxl callers can search-replace `from openpyxl.
  drawing.spreadsheet_drawing` → `from wolfxl.drawing.
  spreadsheet_drawing` directly.
* **No `remove_image` or `replace_image` API yet** — the v1.5
  surface is additive only. Existing images on round-tripped
  workbooks are still preserved verbatim.
* **EMU vs pixel sizing**: anchor coordinates are EMUs (914,400 per
  inch). Use `wolfxl.utils.units.pixels_to_EMU(...)` /
  `points_to_EMU(...)` if you have dimensions in pixels or points.

### Streaming datetime fix (Pod-γ)

* **Behavior change for `read_only=True` workbooks with date-formatted
  cells**: `iter_rows(values_only=True)` now returns `datetime`
  objects where it previously returned Excel serial `float`s. This
  matches openpyxl's behavior and is what most callers expect.
* **If you were relying on the Excel serial floats** (e.g. doing
  arithmetic on them), either drop the `read_only=True` flag (the
  eager path always returned datetimes), or convert manually in
  caller code via `wolfxl.utils.datetime.from_excel(serial_float)`.
* **`StreamingCell.value`** also returns `datetime` for date cells
  in 1.5 — same change, same migration path.

## Out of scope (documented, planned)

The 1.5 release closes "encryption writes" and "image construction"
as the last construction-side T3 stubs. Two large construction items
remain on the roadmap, both now scheduled rather than open-ended:

* **Chart construction** — scheduled for **v1.6.0** (Sprint Μ).
  Modify-mode round-trip already preserves charts verbatim;
  construction (`BarChart`, `LineChart`, `Reference`, `Series`,
  axes) is the headline v1.6.0 deliverable. Full openpyxl
  `openpyxl.chart` parity is the target.
* **Pivot table construction** — scheduled for **v2.0.0** (Sprint Ν).
  Pivot caches and pivot tables are preserved on round-trip but
  cannot be added programmatically. v2.0.0 is the public-launch
  milestone; pivots ship alongside the launch.

Other out-of-scope items (OpenDocument, image transformations / alt
text / hyperlinks, image replace / delete) are tracked in
`tests/parity/KNOWN_GAPS.md` "Out of scope" with deferral
rationales.

## RFCs

- `Plans/rfcs/044-encryption-writes.md` (Sprint Λ Pod-α) — `4bc806c`
- `Plans/rfcs/045-image-construction.md` (Sprint Λ Pod-β) — `d9cb569`

## Stats (post-1.5)

- `cargo test --workspace --exclude wolfxl`: ~660 + N green
  (Pod-β adds image-meta sniffer + dim-extraction tests; Pod-γ adds
  styles-table-aware streaming-datetime tests).
- `pytest tests/`: **1175+ → ~1235+ passed** (Pod-α/β/γ each add
  test cases; the exact count is filled in on integrator merge).
- `pytest tests/parity/`: **140+ → ~165+ passed** (Pod-α adds
  encrypted-write parity, Pod-β adds image parity).
- `KNOWN_GAPS.md` "Out of scope" pruned: write-side encryption and
  image construction lifted to "Closed in 1.5"; chart and pivot
  construction explicitly scheduled for v1.6.0 / v2.0.0.

## Acknowledgments

Sprint Λ ("Lambda") pods that landed 1.5:

- **Pod-α — RFC-044 write-side OOXML encryption.** Feat `4bc806c`,
  test `55dc4c6`, docs `9e9555a`. Merged via `738656a`.
- **Pod-β — RFC-045 image construction (`Image`, anchors, write +
  modify mode emit).** Writer `0ace8c5`, Image+add_image `d9cb569`,
  tests `a73737e`. Merged via `7dc00d2`.
- **Pod-γ — Streaming-datetime correctness fix (Pod-β follow-up
  from 1.3).** Failing test `8af260c`, fix `98cd147`, docs `409837e`.
  Merged via `974b9b5`.
- **Pod-δ (this release scaffold)** — RFC-044 + RFC-045 specs,
  INDEX update, KNOWN_GAPS reconciliation, this release notes
  scaffold, and CHANGELOG entry. Commits: `9c10ce0`, `fff2142`,
  `28caf08`, `deb8155`. Merged via `99079d1`.

## SHA log

| Pod | Branch | Commits | Merge |
|---|---|---|---|
| α | `feat/sprint-lambda-pod-alpha` | `4bc806c`, `55dc4c6`, `9e9555a` | `738656a` |
| β | `feat/sprint-lambda-pod-beta` | `0ace8c5`, `d9cb569`, `a73737e` | `7dc00d2` |
| γ | `feat/sprint-lambda-pod-gamma` | `8af260c`, `98cd147`, `409837e` | `974b9b5` |
| δ | `feat/sprint-lambda-pod-delta` | `9c10ce0`, `fff2142`, `28caf08`, `deb8155` | `99079d1` |

Integrator finalize commit fills these placeholders, performs the
post-merge ratchet flip on `openpyxl.Workbook.save (password kwarg)`,
and tags `v1.5.0`.

After Sprint Λ the openpyxl-parity surface is exhausted at both the
read level (1.0 → 1.4) and the construction level (1.5). The
remaining roadmap is feature-additive: charts (v1.6.0) and pivots
(v2.0.0). Thanks to everyone who file-bugged the encryption and
image construction stubs over the 1.0 → 1.4 cycle — every workload
that hit `NotImplementedError("Image is preserved on modify-mode
round-trip but cannot be added programmatically.")` drove this slice.
