# WolfXL 1.3 — Read-side parity (rich text + streaming + password)

_Date: 2026-04-26_

WolfXL 1.3 closes the three biggest read-side gaps that survived
1.0–1.2's modify-mode and structural-ops focus: rich-text reads
(Pod-α), streaming reads on huge sheets (Pod-β), and
password-protected workbook reads via `msoffcrypto-tool` (Pod-γ).
Sprint Ι ("Iota") wraps the read-path parity story — after 1.3,
everything in `tests/parity/KNOWN_GAPS.md`'s Phase 2 / Phase 3 /
Phase 4 sections is closed; only Phase 5 (`.xls` / `.xlsb`) carries
forward.

The Pod-δ slice that ships alongside the read-path pods adds a
fix-it for the long-standing native-writer VML margin bug on sheets
with custom column widths (D3), exposes
`Workbook.defined_names["X"] = DefinedName(...)` end-to-end so
write-mode users can construct named ranges without dropping into
the Rust API (D4), re-engages the parity ratchet on the open
KNOWN_GAPS rows (D1), and registers four custom pytest marks to
silence the long-running `PytestUnknownMarkWarning` noise (D2).

## What's new

### Rich-text reads: `Cell.rich_text` (Sprint Ι Pod-α, RFC-040)

```python
import wolfxl

wb = wolfxl.load_workbook("template.xlsx")
ws = wb["Notes"]
cell = ws["B7"]                        # has bold "WARNING: " + plain "do not delete"

print(cell.value)                      # → "WARNING: do not delete"  (flat str — unchanged)
print(cell.rich_text)                  # → CellRichText([
                                       #     TextBlock(font=InlineFont(bold=True), text="WARNING: "),
                                       #     "do not delete",
                                       #   ])
```

`Cell.rich_text` returns a `CellRichText`-shaped iterable of
`TextBlock | str` runs when the underlying `<is>` / `<si>` carries
`<r>`/`<rPr>` formatting. Plain-text cells return `None` from
`Cell.rich_text` and continue to surface their value via
`Cell.value` exactly as before.

`InlineFont` mirrors `<rPr>`:
`name | size | bold | italic | underline | strike | color | family
| scheme`. `str(cell.rich_text) == cell.value` is invariant —
stringifying a `CellRichText` concatenates run text in document
order so callers that only need the flat string can keep using
`cell.value`.

Pod-α commit: <!-- TBD: Pod-α commit sha when integrated -->.
RFC: `Plans/rfcs/040-rich-text.md`.

### Streaming reads on huge sheets (Sprint Ι Pod-β, RFC-041)

```python
import wolfxl

# Explicit opt-in
wb = wolfxl.load_workbook("huge.xlsx", read_only=True)
for row in wb.active.iter_rows(values_only=True, max_row=100):
    print(row)
# Memory: O(row width), not O(sheet). 1M-row workbooks now stream
# at < 200 MB peak RSS instead of OOM-ing the kernel.

# Implicit auto-trigger (warn-once per workbook)
wb = wolfxl.load_workbook("huge.xlsx")
# RuntimeWarning: WolfXL auto-enabled streaming reads for "huge.xlsx"
# (sheet 'Sheet1': 1,200,000 rows × 8 cols). Pass read_only=False to opt out.
```

`read_only=True` engages a `quick-xml` SAX path that walks
`<sheetData>` row-by-row and yields one Python tuple per `<row>`.
Sheets that exceed the auto-trigger threshold (default: 50,000 rows
**or** 5 million cells) flip to the streaming path implicitly with
a one-time `RuntimeWarning` so users notice and can opt out via
`read_only=False`.

Streaming-mode workbooks are read-only — `cell.value = …`,
`ws.append(...)`, `wb.create_sheet(...)`, and `wb.save(...)` raise.
`read_only=True + modify=True` raises at load time. Random-access
cell reads (`ws["B7"]`) raise per openpyxl. Match the openpyxl
contract end-to-end.

Pod-β commit: <!-- TBD: Pod-β commit sha when integrated -->.
RFC: `Plans/rfcs/041-streaming-reads.md`.

### Password-protected reads via `msoffcrypto-tool` (Sprint Ι Pod-γ, RFC-042)

```python
import wolfxl

wb = wolfxl.load_workbook("budget.xlsx", password="hunter2")
# 1.2 raised CalamineError; 1.3 decrypts via msoffcrypto-tool
# and parses the plaintext through CalamineStyledBook.open_bytes.
```

Add the optional `crypto` extra to your install:

```bash
pip install 'wolfxl[crypto]'
# or, if you maintain pinned deps directly:
# msoffcrypto-tool>=5.4,<6
```

`load_workbook(path, password=...)` decrypts the .xlsx via
`msoffcrypto-tool` and feeds the plaintext through the existing
`CalamineStyledBook.open_bytes()` reader. `password=None` (default)
short-circuits — no msoffcrypto import, no perf hit on the common
unencrypted path.

Errors:

- Wrong password → `wolfxl.PasswordError` (subclass of `ValueError`)
  with the file path.
- Missing `crypto` extra → `RuntimeError("password kwarg requires
  the 'crypto' extra: pip install wolfxl[crypto]")`.
- Other msoffcrypto errors propagate as-is (corrupt envelope, etc.).

`save()` on a workbook opened with `password=` raises until the
post-1.3 follow-up RFC ships re-encryption. Decryption-then-save
to a *plaintext* file works in modify mode today.

Pod-γ commit: <!-- TBD: Pod-γ commit sha when integrated -->.
RFC: `Plans/rfcs/042-password-reads.md`.

### Pod-δ — sweep and follow-ups

Four small, independent items that round out the release:

- **D1 — Parity ratchet re-enabled.** Five new fine-grained
  KNOWN_GAPS entries land in `tests/parity/openpyxl_surface.py`
  with `wolfxl_supported=False`, so the
  `test_known_gap_still_gaps` test now actually pins the open
  rows. The integrator flips the rich-text / streaming / password
  rows to `True` as the matching pods land; the `.xls` / `.xlsb`
  rows stay open. Pod-δ commit: `751760f`.
- **D2 — Custom pytest marks registered.** `rfc035`, `rfc031`,
  `rfc036`, and `manual` are added to `pyproject.toml`'s
  `[tool.pytest.ini_options].markers` list, silencing the
  recurring `PytestUnknownMarkWarning` noise. Pod-δ commit: `ce9dda3`.
- **D3 — Native-writer VML margin honors per-column widths.**
  `crates/wolfxl-writer/src/emit/drawings_vml.rs::compute_margin`
  hard-coded `COL_WIDTH_PT = 48.0`; sheets with custom column
  widths rendered comment popups over the wrong cell area. The
  new `compute_margin_with_widths` walks `worksheet.columns` and
  sums per-column widths in points, mirroring the modify-mode
  patcher's `compute_margin_with_widths`. Empty `<cols>` falls
  back to the legacy math so existing fixtures stay byte-stable.
  Pod-δ commit: `92c901d`. Closes
  `Plans/followups/native-writer-vml-margin-fix.md`.
- **D4 — `Workbook.defined_names["X"] = DefinedName(...)` shipped.**
  The Python proxy already routed through `_pending_defined_names`;
  Pod-δ adds Excel-compliant name validation (no whitespace,
  no leading digit, not an A1-style ref, not the R/C R1C1
  reserved tokens) and fixes the writer payload so sheet-scope
  names route via `scope=sheet` plus the resolved sheet name.
  Pod-δ commit: `b64c364`. Closes the Phase 1 KNOWN_GAPS row.

Pod-δ also ships RFC-040 / 041 / 042 spec drafts (commit
`6bc120c`) and this release-notes scaffold (commit
<!-- TBD: D6 commit sha when committed -->).

## Breaking changes

### Cell.value behaviour on rich-text cells (RFC-040)

`Cell.value` continues to return a flat `str` for cells with
`<is>`/`<si>` rich text — this preserves the 1.2 contract
end-to-end. The new `Cell.rich_text` accessor is the official
path for callers that need the per-run formatting.

If you previously **relied on `Cell.value` flattening rich text to
plain `str`** (i.e. wrote `if isinstance(cell.value, str): …`),
your code keeps working as-is — `Cell.value` still returns a
plain `str` even on rich-text cells. The break only matters for
callers who want to OPT IN to rich-text awareness; those callers
add a `cell.rich_text` lookup and fall back to `cell.value` when
it returns `None`.

If a future release decides to flip `Cell.value` to return
`CellRichText` directly on rich-text cells (currently filed as
RFC-040 §11 OQ-1), a `Cell.value_str` accessor will land in the
same release as a guaranteed-`str` escape hatch. Until then, no
behaviour change is required for `Cell.value` consumers.

## Migration guide

No source changes are required for callers that worked on 1.2 —
every Pod-α / β / γ / δ change is additive. Optional adjustments:

- **Rich text consumers**: replace any "I want runs but had to
  re-parse `Cell.value`" workaround with `Cell.rich_text`.
  `str(cell.rich_text)` matches `cell.value` for the same cell, so
  callers that previously did `cell.value` keep working until they
  decide to opt in to runs.
- **Large-fixture ingest**: drop the explicit `pyexcelerate` /
  `python-calamine` workaround you may have used to stream huge
  sheets — `wolfxl.load_workbook(path, read_only=True)` now
  matches the openpyxl streaming contract. The auto-trigger means
  long-running pipelines that ingest mixed file sizes will see
  one `RuntimeWarning` per huge workbook; suppress with
  `warnings.simplefilter("ignore", RuntimeWarning)` or pass
  `read_only=False` to opt out.
- **Password-protected files**: `pip install 'wolfxl[crypto]'` and
  pass `password=`. Existing pipelines that detected the encrypted
  file via the previous `CalamineError` should switch to the
  pre-flight `msoffcrypto.OfficeFile(...).is_encrypted()` check or
  catch the new `wolfxl.PasswordError`.
- **VML comment positioning**: if you re-saved a file with
  `wolfxl.Workbook()` (write mode) and noticed comment popups in
  the wrong place when columns had custom widths, the bug is
  fixed in 1.3. No action required — the patch ships in the
  emitter, not the API surface.
- **Sheet-scope defined names via `wb.defined_names`**: previously
  `wb.defined_names["X"] = DefinedName(name="X", value="...",
  localSheetId=0)` round-tripped, but the Rust writer received
  `scope=workbook` regardless of `localSheetId`, so the saved
  file did not carry a `<definedName localSheetId="0">`. 1.3
  routes the scope correctly. If you depended on the old
  silent-workbook-scope behaviour, drop `localSheetId=` from the
  call site.

## Known limitations

Carry-forward from 1.2:

- **`.xls` / `.xlsb`**: still deferred (Phase 5 / `tests/parity/KNOWN_GAPS.md`).
  openpyxl itself doesn't read `.xls` or `.xlsb`; closing this
  gap requires migrating WolfXL from `calamine-styles` to
  upstream `calamine`. No timeline yet.
- **`copy_worksheet` re-saved by openpyxl**: as in 1.2, openpyxl's
  loader is the lossy step on a wolfxl-emitted clone. Stay inside
  wolfxl until the final save.
- **Cross-workbook copy** (`copy_worksheet(other_wb_sheet)`):
  remains out of scope per RFC-035 §10. openpyxl rejects the same
  call.
- **Chart sheets** (`<chartsheet>`): remain out of scope per
  RFC-035 §10.

New limitations introduced by 1.3 (deliberately deferred):

- **Rich-text writes**: RFC-040 ships read-only. Constructing a
  `CellRichText` and saving is post-1.3 work — see RFC-040 §10.
- **Streaming writes** (openpyxl's `write_only=True`): not in
  scope for RFC-041. Users who need bulk-write performance keep
  using the eager write path.
- **Encrypted writes**: `wb.save(..., password=...)` raises until
  the post-1.3 follow-up RFC ships CFB-envelope emission.

See `tests/parity/KNOWN_GAPS.md` for the full per-feature gap list.

## Acknowledgments

Sprint Ι ("Iota") pods that landed 1.3:

- **Pod-α — RFC-040 rich-text reads.** <!-- TBD: Pod-α commit sha when integrated -->
- **Pod-β — RFC-041 streaming reads.** <!-- TBD: Pod-β commit sha when integrated -->
- **Pod-γ — RFC-042 password-protected reads.** <!-- TBD: Pod-γ commit sha when integrated -->
- **Pod-δ (this release scaffold)** — D1 ratchet (`751760f`),
  D2 pytest marks (`ce9dda3`), D3 VML margin fix (`92c901d`),
  D4 `defined_names.__setitem__` (`b64c364`), D5 RFC drafts
  (`6bc120c`), D6 release-notes scaffold
  (<!-- TBD: D6 commit sha when committed -->).

Specs: `Plans/rfcs/040-rich-text.md`,
`Plans/rfcs/041-streaming-reads.md`,
`Plans/rfcs/042-password-reads.md`. Each ships with a §11 Open
Questions block — Pod owners (and the integrator) should resolve
these in the merge PR rather than carrying them into 1.4.

Thanks to everyone who file-bugged the read-side gaps over the
1.0 → 1.2 cycle — every row in the Phase 2 / 3 / 4 KNOWN_GAPS
tables came from a real workload that hit the limitation in
production.
