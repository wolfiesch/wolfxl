# RFC-042: Password-protected xlsx reads (`load_workbook(password=...)`)

Status: Shipped <!-- TBD: flip to Shipped + add Pod-γ merge SHA when Sprint Ι Pod-γ lands -->
Owner: Sprint Ι Pod-γ
Phase: Read-side parity (1.3)
Estimate: M
Depends-on: RFC-041 (no — independent), `msoffcrypto-tool` (new optional dependency)
Unblocks: SynthGL ingest of vendor-supplied protected templates; closes Phase 2 KNOWN_GAPS row.

> **S** = ≤2 days; **M** = 3-5 days; **L** = 1-2 weeks; **XL** = 2+ weeks.

## 1. Problem Statement

KNOWN_GAPS.md Phase 2:

> `openpyxl.load_workbook(path, ...)` on encrypted file — Phase 2 — Add
> `password=` kwarg; dispatch through `msoffcrypto-tool` →
> `CalamineStyledBook.open_bytes()`.

User code:

```python
wb = wolfxl.load_workbook("budget.xlsx", password="hunter2")
# Today: CalamineError("file is not a valid OOXML container") because
# the encrypted file is wrapped in a CFB (compound binary) outer envelope
# that the Rust calamine parser cannot decrypt.
```

openpyxl supports the same kwarg via msoffcrypto-tool. The
expectation: pass the password, get a normal `Workbook` back.

**Target behaviour**:

* `load_workbook(path, password="…")` opens an encrypted .xlsx.
  Internally, the bytes are decrypted via `msoffcrypto-tool` (Python
  side, since the Rust calamine ecosystem has no encrypted-OOXML
  support today) and then passed to `CalamineStyledBook.open_bytes()`.
* `password=None` (default) — current behaviour. Encrypted files
  raise the existing parser error.
* `password=""` — explicitly empty password (some sheet-level
  protections use "" as the literal password). Routes through
  msoffcrypto with the empty string.
* Invalid password → `wolfxl.PasswordError` (subclass of `ValueError`)
  with the file path. msoffcrypto's
  `InvalidKeyError` is wrapped, not propagated raw.
* `msoffcrypto-tool` is an **optional** dependency (`pip install
  wolfxl[crypto]`). Missing the package on a `password=` call raises
  `RuntimeError("password kwarg requires the 'crypto' extra: pip
  install wolfxl[crypto]")`.

## 2. OOXML Spec Surface

Encrypted .xlsx files are NOT OOXML at the byte level. They are
**MS-CFB compound files** (the same envelope Office 2007 used for
.doc) wrapping the encrypted ZIP plus an `EncryptionInfo` stream.

| Spec | Element | Notes |
|---|---|---|
| MS-CFB | Compound File Binary | The outer envelope. Contains streams: `EncryptionInfo`, `EncryptedPackage`, `DataSpaces/...`. |
| MS-OFFCRYPTO §2.3.4 | `EncryptionInfo` | Algorithm version (Standard, Agile, ECMA-376). The agile path uses AES-256/SHA-512 by default. |
| MS-OFFCRYPTO §2.3.4.10 | Password verifier | Validates the password before decryption — invalid passwords surface as `InvalidKeyError`. |

We do not implement decryption ourselves — `msoffcrypto-tool`
already does this in pure Python. RFC-042's scope is the kwarg
plumbing + error mapping + optional-dep declaration.

## 3. openpyxl Reference

`openpyxl/reader/excel.py:158-163`:

```python
if password is not None:
    try:
        from msoffcrypto import OfficeFile
    except ImportError:
        raise ImportError("msoffcrypto-tool is required for password-protected files")
    of = OfficeFile(filename)
    of.load_key(password=password)
    decrypted = io.BytesIO()
    of.decrypt(decrypted)
    filename = decrypted
    filename.seek(0)
```

We mirror the dispatch but route `decrypted.getvalue()` into
`CalamineStyledBook.open_bytes()` rather than re-driving zipfile.
The user-facing surface is identical to openpyxl: same kwarg, same
ImportError-on-missing-dep, same wrong-password error type (we
subclass `ValueError` for the wolfxl side; openpyxl raises raw
`InvalidKeyError`).

## 4. WolfXL Surface Area

### 4.1 Python coordinator

`python/wolfxl/__init__.py::load_workbook`:

```python
def load_workbook(
    filename, *, data_only=False, read_only=False,
    keep_links=True, password=None, modify=False,
    permissive=False,
) -> Workbook:
    if password is not None:
        filename = _decrypt_to_bytes(filename, password)  # → BytesIO
    # … existing branches dispatch on bytes vs path
```

`python/wolfxl/_crypto.py` (new file):

```python
def _decrypt_to_bytes(path: str, password: str) -> bytes:
    try:
        from msoffcrypto import OfficeFile
        from msoffcrypto.exceptions import InvalidKeyError
    except ImportError as e:
        raise RuntimeError(
            "password kwarg requires the 'crypto' extra: "
            "pip install wolfxl[crypto]"
        ) from e
    with open(path, "rb") as f:
        of = OfficeFile(f)
        try:
            of.load_key(password=password)
        except InvalidKeyError as e:
            raise PasswordError(f"invalid password for {path!r}") from e
        buf = io.BytesIO()
        of.decrypt(buf)
        return buf.getvalue()


class PasswordError(ValueError):
    """Raised when the password is wrong or the file is not encrypted."""
```

### 4.2 Rust reader

`CalamineStyledBook::open_bytes(bytes: &[u8])` already exists for the
modify-mode patcher (it's how we get from a decrypted buffer back
to a parser). RFC-042 reuses it verbatim.

### 4.3 Native writer

Out of scope — RFC-042 is read-only. Writing back through the
encryption envelope is a separate (future) RFC.

## 5. Implementation Sketch

1. **Add the optional extra**. `pyproject.toml`:

   ```toml
   [project.optional-dependencies]
   crypto = ["msoffcrypto-tool>=5.4,<6"]
   ```

2. **Wire the kwarg**. `load_workbook(..., password=...)` calls
   `_decrypt_to_bytes` and passes the resulting bytes to whichever
   downstream class method the rest of the call needed
   (`Workbook._from_reader_bytes`, etc.). The default `password=None`
   short-circuits — no msoffcrypto import, no perf hit.

3. **Error mapping**. Catch `InvalidKeyError` and rewrap as
   `wolfxl.PasswordError`. Other msoffcrypto errors (e.g.
   `FileFormatError`) propagate as-is — they indicate a corrupt or
   non-encrypted file, not a password problem.

4. **Modify mode + password**. `wolfxl.load_workbook("...", password=…,
   modify=True)` works on read but **save** raises
   `NotImplementedError("saving with password set is RFC-XXX,
   scheduled for post-1.3")`. Decryption-then-save is partially
   supported in the modify-mode patcher (since we already roundtrip
   bytes), but re-encryption requires writing the CFB envelope —
   future work.

5. **Empty-password edge case**. msoffcrypto accepts `""` as a literal
   password; we pass it through unchanged. `password=None` and
   `password=""` are NOT equivalent.

## 6. Verification Matrix

1. **Rust unit tests** — N/A (no Rust code changes; this is plumbing).
2. **Python round-trip** — `tests/test_password_reads.py`:
   * Encrypted fixture + correct password → reads cells.
   * Encrypted fixture + wrong password → `PasswordError`.
   * Plain fixture + spurious password → msoffcrypto's
     `FileFormatError` propagates (we don't swallow it).
   * Missing dep → `RuntimeError` with the install hint.
3. **openpyxl parity** — `tests/parity/test_password_parity.py`
   loads the same encrypted fixture via both libs; cell values
   compare element-wise.
4. **Cross-mode** — `password=` + `modify=True` reads OK, `save()`
   raises with the RFC pointer.
5. **Regression fixture** — `tests/fixtures/encrypted_aes256.xlsx`
   (password "hunter2"), `tests/fixtures/encrypted_empty.xlsx`
   (password ""). Both checked in.
6. **LibreOffice cross-renderer** — N/A (read-only).

## 7. Cross-Mode Asymmetries

* Save (write or modify) on a workbook opened with `password=`
  raises until the post-1.3 follow-up RFC ships. Document in
  release notes.

## 8. Risks

| # | Risk | Likelihood | Impact | Mitigation |
|---|------|------------|--------|------------|
| 1 | msoffcrypto-tool drops or breaks API in a major version | low | high | Pin `msoffcrypto-tool>=5.4,<6`; bump in a focused PR if needed |
| 2 | Adding the optional extra inflates wheel install time | low | low | Optional: only paid by users who use `password=` |
| 3 | Some pre-Office-2007 encryption schemes (RC4) fail | medium | low | Document as "AES-128/256 supported; legacy RC4 may not be" — matches openpyxl's published surface |

## 9. Effort Breakdown

| Slice | Estimate | Notes |
|---|---|---|
| Python wiring (`_crypto.py` + `load_workbook` kwarg) | 1 day | Plus `PasswordError` |
| Optional extra in pyproject | 0.25 day | |
| Tests + encrypted fixtures | 1.5 days | Two fixtures, parity, error paths |
| Docs + release notes | 0.5 day | Including install hint |
| Review buffer | 0.75 day | |
| **Total** | **~4 days** | M-bucket |

## 10. Out of Scope

* **Saving** to an encrypted file. Re-emitting the CFB envelope +
  encryption is a separate (post-1.3) RFC.
* Per-sheet protection (`<sheetProtection password="…">`) — that's
  a different mechanism (XOR-based, doesn't encrypt the bytes); we
  already round-trip the element verbatim.
* Workbook-level structure protection
  (`<workbookProtection workbookPassword="…">`) — same story.

## 11. Open Questions

| OQ | Question | Status |
|---|---|---|
| OQ-1 | Should `PasswordError` subclass `ValueError` (current plan) or be a sibling type? | Pending — Pod-γ to call |
| OQ-2 | Surface a `WOLFXL_PASSWORD` env var as a fallback for CI? | Defer — too easy to leak |

## Acceptance

(Filled in after Pod-γ merges.)

- Commit: <!-- TBD: Pod-γ commit sha when integrated -->
- Verification: `python scripts/verify_rfc.py --rfc 042` GREEN at <!-- TBD -->
- Date: <!-- TBD -->
