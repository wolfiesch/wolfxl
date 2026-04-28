# Password-Protected Workbooks

wolfxl supports OOXML password protection on both the read and write
paths via the optional [`msoffcrypto-tool`][mso] dependency. Install
it with:

```bash
pip install wolfxl[encrypted]
```

[mso]: https://github.com/nolze/msoffcrypto-tool

## Quick Start

```python
import wolfxl

# Encrypt-on-save: pass password= to Workbook.save
wb = wolfxl.Workbook()
wb.active["A1"] = "secret"
wb.save("confidential.xlsx", password="my-password")

# Decrypt-on-load: pass password= to load_workbook
wb2 = wolfxl.load_workbook("confidential.xlsx", password="my-password")
print(wb2.active["A1"].value)  # -> "secret"
```

`password` accepts `str` or `bytes` (UTF-8 decoded). Passing
`password=None` or omitting the argument produces / consumes a
plaintext xlsx — same as openpyxl's behaviour.

Empty passwords (`""` / `b""`) raise `ValueError("empty password not
allowed")` on save. On read, an empty / wrong password raises
`ValueError` with a clear "wrong password?" hint.

## Supported Algorithms

OOXML supports three families of password protection. wolfxl's coverage
matches what `msoffcrypto-tool`'s upstream library is able to produce
or consume:

| Algorithm                       | Read | Write | Notes                                                                                              |
| ------------------------------- | ---- | ----- | -------------------------------------------------------------------------------------------------- |
| **Agile (AES-256, SHA-512)**    | ✅   | ✅    | The modern Excel default. Used for every wolfxl `save(..., password=...)` call.                    |
| Standard / ECMA-376 (AES-128)   | ✅   | ❌    | Read-only. msoffcrypto-tool does not implement Standard *encryption*, only decryption.             |
| XOR obfuscation (legacy `.xls`) | ✅   | ❌    | Decrypt-only. XOR obfuscation is BIFF-era and not part of the OOXML spec.                          |

If you need to write a Standard or XOR-encrypted file, wolfxl is the
wrong tool — please file an issue describing your use case so we can
discuss alternatives (e.g. shelling out to LibreOffice).

## Architecture (write side)

`Workbook.save(path, password=...)` is layered on top of the existing
plaintext save path:

1. Validate the password (empty rejected up front so we don't leak a
   plaintext tempfile).
2. Materialise the plaintext xlsx via the normal Rust writer / patcher
   into a `tempfile.NamedTemporaryFile` next to `path`'s parent dir.
3. Hand the bytes to `wolfxl._encryption.encrypt_xlsx_to_path`, which
   wraps `msoffcrypto.format.ooxml.OOXMLFile.encrypt(password, outfile)`.
4. Atomic-rename the encrypted file into place.
5. Always clean up the plaintext tempfile, including on error paths.

The Rust crates are not modified — encryption stays Python-side, same
as Sprint Ι Pod-γ's read path.

### Implementation note: tiny-file workaround

msoffcrypto-tool's OOXML container writer routes any `EncryptedPackage`
stream ≤ 4096 bytes through the OLE2 *MiniFAT* sectors, but the
directory entry's `StartingSectorLocation` is set to the regular-FAT
offset. The result is a misaligned stream that fails AES-CBC decrypt
on the way back. Real-world xlsx files (with sheets, styles, etc.)
encrypt to >4096 bytes naturally, so this edge case only affects
synthetic / minimal workbooks.

wolfxl works around the issue by inflating very small plaintext blobs
to 5120 bytes via the ZIP End-Of-Central-Directory comment field
(`_pad_zip_via_eocd_comment`). The padding is part of the formal ZIP
comment, so every standards-compliant ZIP reader (including
msoffcrypto-tool's internal re-read after decrypt) accepts it without
complaint. The padded plaintext is *not* observable on disk — the
caller only sees the encrypted output.

## Modify-mode

`Workbook.save(..., password=...)` works in both write mode (`Workbook()`)
and modify mode (`load_workbook(path, modify=True)`). Encryption is
applied to the final byte stream regardless of which Rust backend
produced it.

```python
wb = wolfxl.load_workbook("plaintext.xlsx", modify=True)
wb.active["B2"] = "edited"
wb.save("encrypted.xlsx", password="pw")
```

A workbook that was *opened* with `password=` and saved without one
produces a plaintext output (T3 out-of-scope behaviour from Sprint Ι
Pod-γ — passing `password=` on save explicitly re-encrypts).

## Round-trip verification

```python
import wolfxl

wb = wolfxl.Workbook()
wb.active["A1"] = "round-trip"
wb.save("rt.xlsx", password="pw")

wb2 = wolfxl.load_workbook("rt.xlsx", password="pw")
assert wb2.active["A1"].value == "round-trip"
```

The same workflow is exercised end-to-end in
`tests/test_encrypted_writes.py` and
`tests/parity/test_encrypted_write_parity.py`.
