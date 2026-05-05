# WolfXL 1.4 — `.xlsb` / `.xls` reads + bytes/BytesIO/file-like input

_Date: 2026-04-26_

WolfXL 1.4 closes Phase 5 — the last open row in
`tests/parity/KNOWN_GAPS.md`. Sprint Κ ("Kappa") ships runtime-dispatched
calamine backends for `.xlsb` and `.xls`, plus a unified bytes /
`io.BytesIO` / file-like input path on `load_workbook` across all
three formats. After 1.4 the openpyxl-parity roadmap is exhausted;
only out-of-scope items remain (write-side encryption, OpenDocument,
chart construction, etc.).

## Highlights

- **`.xlsb` and `.xls` reads** via runtime-dispatched calamine
  backends — values + cached formula results round-trip. Closes
  Phase 5 of the openpyxl-parity roadmap.
- **Bytes / `BytesIO` / file-like input** on `wolfxl.load_workbook(...)`
  across all formats. The Sprint Ι Pod-γ tempfile workaround for
  password reads is gone; decrypted bytes flow direct.
- **`Workbook._format` introspection attribute** (`'xlsx' | 'xlsb' | 'xls'`)
  for callers that need to branch on the underlying format.

## What's new

### `.xlsb` reads (values + cached formula results)

```python
import wolfxl

wb = wolfxl.load_workbook("finance_q3.xlsb")
ws = wb.active
for row in ws.iter_rows(values_only=True):
    print(row)
# Values + cached formula results come out the same shape as the
# matching .xlsx workbook.
```

`Workbook._format == "xlsb"` for these workbooks. Style accessors
(`cell.font`, `cell.fill`, `cell.border`, `cell.alignment`,
`cell.number_format`) raise `NotImplementedError` — see
"Limitations" below.

Pod-α commit: `b805aac`

### `.xls` reads (values + cached formula results)

```python
import wolfxl

wb = wolfxl.load_workbook("legacy_erp.xls")
ws = wb.active
for row in ws.iter_rows(values_only=True):
    print(row)
# Same contract as .xlsb: values + cached formula results, no styles.
```

`Workbook._format == "xls"`. Same style-accessor restriction as
`.xlsb` (the binary formats encode styles inline differently from
xlsx's separate `xl/styles.xml`; the calamine-styles fork only
exposes the xlsx style path).

Pod-α commit: `b805aac`

### Bytes / `BytesIO` / file-like input

```python
import io
import wolfxl

# Bytes (raw)
blob = open("data.xlsx", "rb").read()
wb = wolfxl.load_workbook(blob)

# BytesIO
wb = wolfxl.load_workbook(io.BytesIO(blob))

# Any Read-able file-like
with open("data.xlsb", "rb") as f:
    wb = wolfxl.load_workbook(f)

# Path (unchanged)
wb = wolfxl.load_workbook("data.xlsx")
```

Works across all three formats (`xlsx`, `xlsb`, `xls`). Each backend
exposes a `Source` enum with `File(BufReader<File>)` and
`Bytes(Cursor<Vec<u8>>)` arms — both implement `Read + Seek`, so
calamine's underlying `Reader` trait dispatches uniformly.

This refactor lands the bytes-direct path that Sprint Ι Pod-γ
worked around with a tempfile. Password reads now route through
`open_from_bytes` end-to-end without the tempfile hop.

Pod-β commit: `ddf0dc5`

### `_rust.classify_format(path_or_bytes)` magic-byte sniffer

```python
from wolfxl import _rust

_rust.classify_format("data.xlsx")        # → "xlsx"
_rust.classify_format("data.xlsb")        # → "xlsb"
_rust.classify_format("data.xls")         # → "xls"
_rust.classify_format(b"\xD0\xCF\x11\xE0...")   # → "xls"
_rust.classify_format(b"PK\x03\x04...")          # ZIP probe → "xlsx" or "xlsb"
_rust.classify_format(b"<garbage>")              # → "unknown"
```

The detector runs once per `load_workbook` call and selects the
backend pyclass. Magic-byte rules:

* `D0 CF 11 E0 A1 B1 1A E1` → OLE compound file → `.xls` (or
  encrypted ooxml — disambiguated by stream layout).
* `50 4B 03 04` → ZIP local file header → probe central directory
  for `xl/workbook.xml` (xlsx) vs `xl/workbook.bin` (xlsb).
* OpenDocument (`mimetype` containing
  `application/vnd.oasis.opendocument`) → `"ods"` → raises with the
  ODS limitation pointer (out of scope).
* Anything else → `"unknown"` → raise `ValueError` with the first
  8 bytes hex-dumped.

Pod-α commit: `b805aac`

### Pre-built `.xlsb` / `.xls` parity fixtures

`tests/parity/test_xlsb_reads.py` and `tests/parity/test_xls_reads.py`
land alongside committed binary fixtures under
`tests/parity/fixtures/xlsb/` and `tests/parity/fixtures/xls/`. Each
pins shape + values element-wise against
`pandas.read_excel(engine="calamine")` —
openpyxl reads neither format, so pandas+calamine is the de-facto
parity target for "binary Excel decoded correctly" in the Python
ecosystem.

Pod-γ commits: `97585a5` (fixtures) + `49e95d5` (parity tests)

## Limitations

The 1.4 read path for `.xlsb` and `.xls` is intentionally minimal.
Each call below raises `NotImplementedError` with a pointer at
RFC-043 so users hit a documented wall, not a mystery error.

* **Modify-mode is xlsx-only.**
  `wolfxl.load_workbook("foo.xlsb", modify=True)` raises. Workaround:
  load values, reconstruct as a fresh `Workbook()`, save as `.xlsx`.
  The modify-mode patcher is xlsx-ZIP-specific and porting it to
  the binary formats would require a separate binary patcher.
* **Streaming SAX (`read_only=True`) is xlsx-only.**
  `.xlsb` / `.xls` load eagerly — calamine's binary-format readers
  don't expose a SAX-style row iterator the way the xlsx XML parser
  does.
* **`password=` is xlsx-only.**
  `msoffcrypto-tool` only handles the OOXML encryption envelope.
  Encrypted `.xlsb` (rare in the wild) and encrypted `.xls`
  (legacy CryptoAPI / RC4) are out of scope.
* **Style accessors raise on `.xlsb` / `.xls` workbooks.**
  `cell.font`, `cell.fill`, `cell.border`, `cell.alignment`,
  `cell.number_format` all raise `NotImplementedError`. Branch via
  `wb._format != "xlsx"` to handle this in caller code.
* **`Workbook.save("out.xlsb")` and `Workbook.save("out.xls")` are
  unsupported.** The native writer is xlsx-only.
* **OpenDocument (`.ods`) is out of scope.** Detected and rejected
  by `classify_format` with a friendly error.

## API additions

```python
# Non-path inputs on load_workbook (NEW)
wolfxl.load_workbook(b"PK\x03\x04...")        # raw bytes
wolfxl.load_workbook(io.BytesIO(blob))         # BytesIO
wolfxl.load_workbook(open("data.xlsb", "rb"))  # file-like

# Format introspection (NEW)
wb._format  # → 'xlsx' | 'xlsb' | 'xls'
```

## Migration notes from 1.3

- **No breaking changes for xlsx workflows.** Every existing
  `.xlsx` read path is unchanged; the new dispatch only kicks in
  on non-xlsx inputs and the non-path input branches.
- **`_from_bytes` (private) refactored.** Sprint Ι password reads
  used a tempfile to bridge the bytes/path gap; 1.4 routes the
  decrypted bytes directly through `open_from_bytes` on each
  backend. The user-facing `password=` contract is unchanged
  (still cleans up via `Workbook.close()`), but the tempfile is
  gone — workloads that scrutinise `/tmp` will see one fewer
  artefact per password load.
- **`load_workbook` raises `ValueError` on unknown formats.** Code
  that catches `IOError` on bad `load_workbook` inputs (corrupt
  bytes, wrong file type) should also catch `ValueError` — the
  magic-byte sniffer raises with the first 8 bytes hex-dumped when
  it can't classify the input.
- **Style access on `.xlsb` / `.xls` cells raises
  `NotImplementedError`.** Call sites that branch on workbook
  format should use `wb._format` rather than try/except.

## RFCs

- `Plans/rfcs/043-xlsb-xls-reads.md` (Sprint Κ) (`fe8b677`)

## Stats (post-1.4)

- `cargo test --workspace --exclude wolfxl`: ~660 + N green
  (Pod-α adds magic-byte sniffer + bytes-input round-trip tests).
- `pytest tests/`: **1106 → ~1175+ passed** (Pod-α/β/γ each add
  test cases; the exact count is filled in on integrator merge).
- `pytest tests/parity/`: **102 → ~140+ passed** (Pod-γ ships
  `tests/parity/test_xlsb_reads.py` + `tests/parity/test_xls_reads.py` element-wise vs
  `pandas.read_excel(engine="calamine")`).
- KNOWN_GAPS.md Phase 5 section removed; openpyxl-parity roadmap
  exhausted post-1.4.

## Acknowledgments

Sprint Κ ("Kappa") pods that landed 1.4:

- **Pod-α — RFC-043 Rust backends + magic-byte sniffer.** `b805aac`
- **Pod-β — RFC-043 Python dispatcher + bytes/BytesIO/file-like.** `ddf0dc5`
- **Pod-γ — RFC-043 parity fixtures + pandas+calamine assertions.** `97585a5` (fixtures) + `49e95d5` (parity tests)
- **Pod-δ (this release scaffold)** — RFC-043 spec, INDEX update,
  KNOWN_GAPS Phase 5 reconciliation, this release notes scaffold,
  and CHANGELOG entry. `fe8b677` (RFC-043) + `9aaf918` (release notes) + `5d22b82` (KNOWN_GAPS)

Spec: `Plans/rfcs/043-xlsb-xls-reads.md`. Each pod owner (and the
integrator) should resolve any §10 open questions in the merge PR
rather than carrying them into 1.5.

Phase 5 closure marks the end of the openpyxl-parity roadmap that
launched in 0.3 and crystallised across 1.0 (modify-mode), 1.1 (
structural ops), 1.2 (RFC-035 follow-ups), and 1.3 (read-side
parity for rich-text / streaming / password). Thanks to everyone
who file-bugged the binary-format gaps over the 1.0 → 1.3 cycle —
every workload that hit `CalamineError` on a `.xlsb` / `.xls`
input drove this slice.
