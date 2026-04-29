# RFC-043: `.xlsb` / `.xls` reads via runtime-dispatched calamine backends

Status: Shipped 1.4 (Sprint Κ) — integrator finalize at e2089f9
Owner: Sprint Κ Pods α / β / γ / δ
Phase: Read-side parity (1.4)
Estimate: L
Depends-on: existing `calamine-styles` fork (xlsx); upstream `calamine` (xlsb / xls value paths).
Unblocks: SynthGL ingest of finance-team .xlsb exports + legacy ERP .xls dumps; closes the
last open `KNOWN_GAPS.md` row (Phase 5).

> **S** = ≤2 days; **M** = 3-5 days; **L** = 1-2 weeks; **XL** = 2+ weeks.

## 1. Status

Shipped in **WolfXL 1.4** (Sprint Κ — "Kappa"). Closes Phase 5 of the
openpyxl-parity roadmap. The remaining KNOWN_GAPS surface after 1.4 is
out-of-scope work only (write-side encryption, OpenDocument, etc.).

## 2. Summary

WolfXL now reads `.xlsb` and `.xls` workbooks via runtime-dispatched
calamine backends. Values and cached formula results round-trip; styles
raise `NotImplementedError`. The same dispatcher accepts bytes,
`io.BytesIO`, and arbitrary file-like objects (in addition to the
existing path-based load) across all three formats.

## 3. Motivation

Phase 5 was the last open row in `tests/parity/KNOWN_GAPS.md`. SynthGL
ingestion sees occasional `.xlsb` from finance teams (Excel's default
"big workbook" binary format) and rare `.xls` from legacy ERP exports
that haven't been migrated. openpyxl itself reads neither format, so
the parity target is the `pandas.read_excel(engine="calamine")`
output: same values, same cached formula results, same sheet ordering.

A secondary motivation is the bytes-input path: Sprint Ι Pod-γ's
password-reads workaround round-tripped decrypted bytes through a
tracked tempfile because the path-only reader couldn't accept an
in-memory buffer. Sprint Κ Pod-β replaces that with a direct
bytes/BytesIO/file-like path on every backend, eliminating the
tempfile hop.

## 4. Architecture

### 4.1 Format detection

`_rust.classify_format(path_or_bytes) -> "xlsx" | "xlsb" | "xls" | "ods" | "unknown"`

Magic-byte sniffer, runs once per `load_workbook` call:

* `D0 CF 11 E0 A1 B1 1A E1` → OLE compound file → likely `.xls` (or
  encrypted `.xlsx` which Pod-γ already routes through msoffcrypto).
  Confirm by parsing the CFB directory and checking for the
  `Workbook` stream (xls) vs `EncryptedPackage` stream (encrypted
  ooxml).
* `50 4B 03 04` → ZIP local file header → probe central directory
  for `xl/workbook.xml` (xlsx) vs `xl/workbook.bin` (xlsb).
* `PK 03 04` + `mimetype` containing `application/vnd.oasis.opendocument`
  → `"ods"` (out of scope; raises a friendly error pointing at the
  ODS limitation).
* Anything else → `"unknown"` → raise `ValueError` with the first 8
  bytes hex-dumped.

### 4.2 Three backend pyclasses

| pyclass | format | styles | reads |
|---|---|---|---|
| `CalamineStyledBook` | xlsx | full (font/fill/border/alignment/numfmt) | values + formulas + styles + rich text + streaming |
| `CalamineXlsbBook` | xlsb | **NotImplementedError** on access | values + cached formula results |
| `CalamineXlsBook` | xls | **NotImplementedError** on access | values + cached formula results |

Backend selection happens once at `load_workbook` time based on
`classify_format`'s answer. The Python `Workbook` carries a
`_format: Literal["xlsx", "xlsb", "xls"]` attribute so call sites
that need to branch (e.g. style accessors) can do so explicitly.

### 4.3 Bytes-input path

Each backend exposes a `Source` enum:

```rust
enum Source {
    File(BufReader<File>),
    Bytes(Cursor<Vec<u8>>),
}
// Implements Read + Seek for both arms.
```

`open_from_bytes(bytes: &[u8])` constructs the `Bytes` arm; the
existing `open` constructor stays on the `File` arm for the
path-input path. calamine's underlying `Reader` trait already
accepts any `Read + Seek`, so the dispatcher is mechanical.

This refactor lands the bytes-direct path that Sprint Ι Pod-γ
worked around with a tempfile; password reads now flow through
`open_from_bytes` end-to-end.

## 5. API surface

```python
import io
import wolfxl

# Path (xlsx / xlsb / xls — auto-detected)
wb = wolfxl.load_workbook("data.xlsx")
wb = wolfxl.load_workbook("data.xlsb")
wb = wolfxl.load_workbook("data.xls")

# File-like
wb = wolfxl.load_workbook(open("data.xlsb", "rb"))

# Bytes
wb = wolfxl.load_workbook(b"PK\x03\x04...")  # raw bytes
wb = wolfxl.load_workbook(io.BytesIO(blob))   # BytesIO

# Format introspection (NEW)
wb._format  # → 'xlsx' | 'xlsb' | 'xls'

# Values come out the same shape regardless of source format
ws = wb.active
for row in ws.iter_rows(values_only=True):
    print(row)
```

## 6. Limitations (explicit "what we don't do")

The 1.4 read path for `.xlsb` and `.xls` is intentionally minimal.
Each limitation below raises with a pointer at this RFC so users
hit a documented wall, not a mystery error.

* **Modify-mode**: `load_workbook("foo.xlsb", modify=True)` raises
  `NotImplementedError`. Workaround: load values, reconstruct as a
  fresh `Workbook()`, save as `.xlsx`. Out of scope because the
  modify-mode patcher is xlsx-ZIP-specific.
* **Streaming SAX**: `read_only=True` is xlsx-only. `.xlsb` / `.xls`
  load eagerly — calamine's binary-format readers don't expose a
  SAX-style row iterator the way the xlsx XML parser does.
* **Password**: `password=` is xlsx-only. `msoffcrypto-tool` only
  handles the OOXML encryption envelope; encrypted `.xlsb` (rare)
  and encrypted `.xls` (legacy CryptoAPI / RC4) are out of scope.
* **Style accessors**: `cell.font`, `cell.fill`, `cell.border`,
  `cell.alignment`, `cell.number_format` raise
  `NotImplementedError("style access on .xlsb/.xls workbooks is RFC-043
  out-of-scope; see _format attribute")` on non-xlsx workbooks. The
  binary formats encode styles inline differently from xlsx's
  separate `xl/styles.xml`; calamine-styles' fork only exposes the
  xlsx style path.
* **Write**: `Workbook.save("out.xlsb")` and
  `Workbook.save("out.xls")` raise on save. The native writer is
  xlsx-only.
* **OpenDocument (`.ods`)**: explicitly out of scope. Detected and
  rejected with a friendly error.

## 7. Compatibility / parity target

openpyxl reads neither `.xlsb` nor `.xls`; openpyxl-comparison parity
is not the gate. Instead the parity target is
`pandas.read_excel(engine="calamine")`, which uses the same
underlying calamine engine and is the de-facto reference for
"binary Excel formats decoded correctly" in the Python ecosystem.

Test fixtures and parity assertions live in:

* `tests/parity/test_xlsb_reads.py` — values + cached formula
  results on a hand-built `.xlsb` fixture.
* `tests/parity/test_xls_reads.py` — same shape on an `.xls`
  fixture exported from a known reference workbook.

Each test file pins shape + values element-wise against the pandas
output and round-trips through wolfxl's `iter_rows(values_only=True)`.

## 8. Implementation pods

Sprint Κ ships in four parallel pods. Each pod's merge SHA is
filled in by the integrator post-merge.

| Pod | Scope | Merge SHA |
|---|---|---|
| Pod-α | Rust backends: `CalamineXlsbBook` + `CalamineXlsBook` + `open_from_bytes` on all three pyclasses + `_rust.classify_file_format` magic-byte sniffer (renamed from the spec's `classify_format` to avoid collision with the existing SynthGL archetype classifier). | `b805aac` |
| Pod-β | Python dispatcher: `load_workbook` accepts `bytes` / `BytesIO` / file-like in addition to path; routes through `classify_file_format`; sets `Workbook._format`; raises on style access for non-xlsx. Drops Sprint Ι Pod-γ's tempfile workaround for the bytes path. | `ddf0dc5` |
| Pod-γ | Pre-built `.xlsb` / `.xls` parity fixtures (`97585a5`) + parity assertions vs `pandas.read_excel(engine="calamine")` (`49e95d5`). | `97585a5` + `49e95d5` |
| Pod-δ | Docs: this RFC (`fe8b677`), 1.4 release notes scaffold (`9aaf918`), KNOWN_GAPS reconciliation Phase 5 closed (`5d22b82`). | `fe8b677` + `9aaf918` + `5d22b82` |
| Integrator | Drift: encrypted-xlsx CFB disambiguation, `open_from_bytes` permissive-arg fallback, `wolfxl.classify_file_format` re-export, Pod-γ test `data_only=True` + epoch normalization, ratchet flip Phase-5 → shipped-1.4, release-notes-1.4.md path normalization. | `e2089f9` |

## 9. Verification Matrix

1. **Rust unit tests** — magic-byte sniffer (xlsx / xlsb / xls / ods
   / garbage); each backend's `open_from_bytes` round-trip; `Source`
   enum's `Read+Seek` impl on both arms.
2. **Python round-trip** — `tests/parity/test_xlsb_reads.py` +
   `tests/parity/test_xls_reads.py` cover values + formulas + sheet
   ordering, including path / BytesIO / bytes / file-like input across
   both binary formats.
3. **Parity** — `tests/parity/test_xlsb_reads.py` +
   `tests/parity/test_xls_reads.py` element-wise vs
   `pandas.read_excel(engine="calamine")`.
4. **Cross-mode** — `password=` + xlsx still works (Pod-γ's
   tempfile workaround is gone, but the user-facing contract is
   unchanged); `modify=True` + xlsb raises with the RFC pointer;
   `read_only=True` + xlsb raises.
5. **Regression fixtures** — committed samples under
   `tests/parity/fixtures/xlsb/` and `tests/parity/fixtures/xls/`
   cover formulas, multi-sheet workbooks, dates, numbers, and strings.
6. **LibreOffice cross-renderer** — N/A (read-only path; no bytes
   emitted).

## 10. Open questions / future work

* **Style approximation for xlsb**: would require reading the
  binary `xl/styles.bin` part. Out of scope unless the calamine-styles
  fork ever ports the styles path; SynthGL's xlsb workloads are
  values-only today.
* **Modify-mode for xlsb**: would require a binary patcher
  (separate from the xlsx ZIP patcher). Out of scope; the current
  "load + transcribe to xlsx" workflow is the documented path.
* **Write-side xlsb / xls**: explicit non-goal. Users who need to
  emit binary formats can post-process wolfxl's `.xlsx` output via
  Excel / LibreOffice / external converters.
* **Encrypted `.xlsb`**: rare in the wild. If a workload needs it,
  re-evaluate post-1.4 with a small follow-up RFC.

## Acceptance

(Filled in after Sprint Κ pods merge.)

- Pod-α commit: `b805aac`
- Pod-β commit: `ddf0dc5`
- Pod-γ commits: `97585a5` (fixtures) + `49e95d5` (parity tests)
- Pod-δ commits: `fe8b677` (RFC) + `9aaf918` (release notes) + `5d22b82` (KNOWN_GAPS)
- Integrator drift commit: `e2089f9`
- Verification: pytest 1190 passed / 14 skipped / 2 xfailed; cargo --workspace --exclude wolfxl ~676 green
- Date: 2026-04-26
