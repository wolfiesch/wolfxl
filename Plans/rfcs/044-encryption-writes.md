# RFC-044: Write-side OOXML encryption (`Workbook.save(path, password=...)`)

Status: Shipped 1.5 (Sprint Λ Pod-α) — feat `4bc806c`, test `55dc4c6`, docs `9e9555a`
Owner: Sprint Λ Pod-α
Phase: 5 (1.5)
Estimate: M
Depends-on: RFC-042 (`msoffcrypto-tool` optional dep), RFC-013 (patcher infra)
Unblocks: T3 closure for encrypted xlsx writes; round-trip-encrypted modify-mode flows

> **S** = ≤2 days; **M** = 3-5 days; **L** = 1-2 weeks; **XL** = 2+ weeks
> (calendar, with parallel subagent dispatch + review).

## 1. Background — Problem Statement

`python/wolfxl/_workbook.py:1032` is a hard stub:

```python
if password is not None:
    raise NotImplementedError(
        "write-side encryption is not yet supported; "
        "see docs/encryption.md (tracked as T3 out-of-scope)"
    )
```

Sprint Ι Pod-γ (RFC-042) shipped read-side decryption end-to-end via
`msoffcrypto-tool`. Modify-mode workbooks opened with `password=`
already round-trip values, but every `wb.save(...)` call drops back
to plaintext. The Sprint Ι release notes documented this asymmetry
as T3 ("Writing encrypted xlsx") and pinned it to a follow-up RFC.

User code that hits the stub today:

```python
import wolfxl

wb = wolfxl.load_workbook("budget.xlsx", password="hunter2")
wb["Sheet1"]["A1"] = "edited"
wb.save("budget.xlsx", password="hunter2")     # NotImplementedError
```

**Target behaviour**: `Workbook.save(path, password="...")` writes
the workbook to disk, encrypted on the way out. The on-disk bytes
are an MS-CFB compound file (the same envelope `msoffcrypto-tool`
already decrypts on the read side), wrapping the plaintext xlsx
ZIP under AES-256 / SHA-512 (Agile algorithm). `password=None`
(default) ≡ plaintext save — current behaviour, unchanged.

## 2. Approach

`msoffcrypto-tool` 5.4+ ships a high-level `OOXMLFile.encrypt()`
helper that takes an output stream, a password, and an algorithm
selector. RFC-044 reuses it verbatim; the patcher and writer keep
emitting plaintext bytes, and the encryption pass is a final
post-flush step on the way to the user-supplied path.

```python
# python/wolfxl/_crypto.py — RFC-044 additions
def _encrypt_to_path(plaintext: bytes, out_path: str, password: str) -> None:
    try:
        from msoffcrypto.format.ooxml import OOXMLFile
    except ImportError as e:
        raise RuntimeError(
            "password kwarg requires the 'encrypted' extra: "
            "pip install wolfxl[encrypted]"
        ) from e
    src = io.BytesIO(plaintext)
    of = OOXMLFile(src)
    with tempfile.NamedTemporaryFile(
        suffix=".xlsx", delete=False, dir=os.path.dirname(out_path) or None,
    ) as tmp:
        of.encrypt(password, tmp)
        tmp_path = tmp.name
    os.replace(tmp_path, out_path)              # atomic
```

The dispatch is wired into `Workbook.save`:

```python
def save(self, filename, *, password=None) -> None:
    filename = str(filename)
    if password is None:
        # Existing plaintext path — unchanged.
        return self._save_plaintext(filename)

    # Buffer plaintext bytes from the existing pipeline (writer or
    # patcher), then re-encrypt to disk under the requested password.
    buf = io.BytesIO()
    self._save_plaintext(buf)
    _encrypt_to_path(buf.getvalue(), filename, password)
```

`_save_plaintext` is the plain-bytes seam that already accepts a
`Path | str | BinaryIO` (Sprint Κ Pod-β unified the bytes input
path; Pod-α reuses the same plumbing on the bytes-output side).

## 3. Algorithm scope — Agile (AES-256) only

`msoffcrypto-tool` exposes three OOXML encryption families:

| Family | OOXML version | Reads (RFC-042) | Writes (RFC-044) |
|---|---|---|---|
| **Agile** (AES-128/256, SHA-1/SHA-512) | Office 2010+ | ✅ | ✅ (this RFC, AES-256/SHA-512 default) |
| **Standard** (AES-128, SHA-1) | Office 2007 | ✅ | ❌ — `OOXMLFile.encrypt()` raises `NotImplementedError`; library is decrypt-only on this path |
| **XOR / RC4 / 40-bit** | pre-2007 legacy | partial | ❌ — same; library is decrypt-only |

**Decision: Agile only for v1.5.0.** The msoffcrypto-tool
authors document Standard + XOR as "decrypt-only because we never
implemented the write-side key-derivation flow". Cloning that work
into wolfxl is out of proportion: Agile is the Office 2010+ default
and is what every modern tool emits; Office still reads it back
without prompting. Sprint Λ Pod-α therefore ships only the AES-256
write path. A future RFC can revisit Standard if a customer
specifically asks for it (no current request).

The default key length is **AES-256**; the default hash is
**SHA-512**. These match the Office 2013+ defaults and exceed
NIST's current recommendations. Users who need to interoperate
with a specific Office version can override via
`Workbook.save(path, password=..., algorithm="agile")` (the
parameter is single-valued in v1.5.0; future-proofing the kwarg
is cheap).

## 4. Public API

```python
# Plaintext save (default — unchanged)
wb.save("out.xlsx")

# Encrypted save (NEW)
wb.save("out.xlsx", password="hunter2")

# Empty-string password (per OOXML spec — some workflows use "")
wb.save("out.xlsx", password="")
```

`password=None` ≡ plaintext (default). `password=""` is **NOT**
equivalent to `password=None`; an empty-string password produces a
real encryption envelope with a literal empty key, matching the
read-side semantics RFC-042 already documents.

**Errors**:

* `RuntimeError` on missing optional dep, with the
  `pip install wolfxl[encrypted]` install hint (mirrors RFC-042's
  `_decrypt_to_bytes` error path).
* `TypeError` if `password` is not `str | bytes | None`.
* Any underlying `msoffcrypto-tool` exception propagates — these
  indicate a tempfile / IO failure, not a wolfxl-specific bug.

## 5. Modify mode + write mode coverage

Both flows reach the same final byte stream:

| Mode | Bytes source | Encryption pass |
|---|---|---|
| **Write mode** (`Workbook()`-constructed) | `crates/wolfxl-writer` emits a fresh ZIP | `_save_plaintext(BytesIO)` → `_encrypt_to_path` |
| **Modify mode** (`load_workbook(modify=True)`) | `XlsxPatcher::do_save` rewrites the source ZIP | `_save_plaintext(BytesIO)` → `_encrypt_to_path` |

The encryption pass is **mode-agnostic** because it operates on the
plaintext ZIP bytes, not on the upstream writer or patcher. This is
the same architectural seam Sprint Κ Pod-β used to unify
bytes-input across `xlsx`/`xlsb`/`xls` reads — encryption sits one
layer above the writer/patcher choice.

The Sprint Ι read-side path also flows through this seam:
`load_workbook(p, password=..., modify=True)` decrypts on read,
the patcher mutates the in-memory state, and `wb.save(p2,
password=...)` re-encrypts on the way out. The full encrypted
modify-mode round-trip is thus end-to-end self-contained — no
plaintext ever touches disk.

## 6. Round-trip

The Sprint Ι read-side decryption (RFC-042) and Sprint Λ write-side
encryption (this RFC) compose:

```python
import wolfxl

# Read encrypted → mutate → write encrypted, never touching plaintext on disk.
wb = wolfxl.load_workbook("budget_2026.xlsx", password="hunter2", modify=True)
wb["Q1"]["B7"] = "updated"
wb.save("budget_2026.xlsx", password="hunter2")   # re-encrypted in place
```

Tests pin both directions: encrypted → plaintext (decrypt only),
plaintext → encrypted (this RFC), and encrypted → encrypted
(round-trip via the modify-mode seam).

## 7. Test plan

New files:

* `tests/test_encrypted_writes.py` — write-side coverage. Cases:
  * Write mode + password — plaintext fixture, save with password,
    re-read with msoffcrypto-tool standalone, byte-compare ZIP
    contents to a plaintext save of the same workbook.
  * Modify mode + password — load encrypted fixture, mutate a
    cell, save with same password, re-read and assert the
    mutation persisted.
  * Modify mode + password change — load encrypted with password
    A, save with password B, re-read with B, assert mutation
    + new password are both honored.
  * Empty-string password — `password=""` is NOT equivalent to
    `password=None`; produces an encrypted file that round-trips
    with the empty literal.
  * Missing dep → `RuntimeError` with install hint (mocked
    import failure).
  * `password=` not `str | bytes | None` → `TypeError`.

* `tests/parity/test_encrypted_write_parity.py` — parity vs
  openpyxl. openpyxl's `save(...)` doesn't support `password=`
  natively (you have to post-process with msoffcrypto-tool
  yourself), so the parity target is "openpyxl save → manual
  msoffcrypto encrypt → re-read". wolfxl `save(password=...)` must
  produce a re-readable file under both msoffcrypto-tool standalone
  AND wolfxl's own RFC-042 read path.

Verification matrix coverage (six-layer):

| Layer | Coverage |
|---|---|
| 1. Rust unit tests | N/A — the encryption pass is pure Python (msoffcrypto-tool). |
| 2. Golden round-trip (diffwriter) | `tests/diffwriter/cases/encrypted_save.py` — opt-in (skipped without `wolfxl[encrypted]`). |
| 3. openpyxl parity | `tests/parity/test_encrypted_write_parity.py` (above). |
| 4. LibreOffice cross-renderer | Manual: open the encrypted output in LibreOffice with the password prompt; document in PR. |
| 5. Cross-mode | `tests/test_encrypted_writes.py` covers write mode + modify mode. |
| 6. Regression fixture | `tests/fixtures/encrypted_writeback.xlsx` (encrypted source, password "hunter2") — checked in. |

## 8. Drift / open questions

**What's NOT shipped in 1.5**:

* **Standard (AES-128, Office 2007) encryption.** msoffcrypto-tool
  is decrypt-only on this path; implementing the write side would
  require porting the Office 2007 key-derivation routine. Tracked
  as a follow-up; no current customer ask.
* **XOR obfuscation / RC4** — legacy, pre-Office-2007. Same
  rationale.
* **Custom EncryptionInfo XML** (e.g. choosing block size,
  spinning iterations, embedding custom certificates). Agile
  defaults are sensible; a future "advanced encryption" RFC can
  add the override knobs if a workload demands them.
* **Per-sheet protection** (`<sheetProtection password="…">`) and
  **workbook structure protection** (`<workbookProtection
  workbookPassword="…">`) — these are XOR-based attribute hashes,
  not envelope encryption. wolfxl already round-trips them
  verbatim; a separate (future) RFC could expose Python-side
  setters.

| OQ | Question | Status |
|---|---|---|
| OQ-1 | Default to AES-256/SHA-512, or expose `algorithm=` selector with `agile_aes128` / `agile_aes256` as values? | Resolution: `algorithm="agile"` accepted as a single-valued forward-compat kwarg in 1.5; AES-256/SHA-512 is the only behavior. Pod-α may broaden if msoffcrypto-tool exposes the knobs cleanly. |
| OQ-2 | If the user passes `password=` to a workbook opened **without** `password=`, do we silently encrypt (current plan) or warn? | Resolution: **silent encrypt**, matching openpyxl's plaintext-source-then-encrypt-on-save flow. The save path is independent of the load path. |
| OQ-3 | Should `Workbook.save(path, password=)` fail-loudly if the optional dep is missing AT IMPORT time (eager guard) or AT SAVE time (lazy guard)? | Resolution: **lazy** — matches RFC-042's read-side behavior. Users who never pass `password=` never pay the import cost. |

## 9. SHA log

Sprint Λ Pod-α landed in three atomic commits on `feat/sprint-lambda-pod-alpha`:

- `4bc806c` `feat(encrypt): add Agile/AES-256 write encryption via msoffcrypto OOXMLFile.encrypt`
- `55dc4c6` `test(encrypt): write-side encryption coverage + parity vs read path`
- `9e9555a` `docs(encrypt): add docs/encryption.md (Agile-only scope; Standard/XOR rationale)`

Merged to `feat/native-writer` as merge commit `738656a` on 2026-04-26.

## Acceptance

(Filled in after Pod-α merges.)

- Commit: `4bc806c` (Pod-α feat); merged via `738656a` on `feat/native-writer`
- Verification: `pytest tests/test_encrypted_writes.py tests/parity/test_encrypted_write_parity.py` GREEN — 13 cases passing
- Date: 2026-04-26
