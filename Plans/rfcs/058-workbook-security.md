# RFC-058 — Workbook-level security (`wb.security`, `WorkbookProtection`, `FileSharing`)

> **Status**: Approved
> **Phase**: 5 (2.0 — Sprint Ο)
> **Depends-on**: 044 (encryption — for password hash helpers), 011 (xml-merger)
> **Unblocks**: 060
> **Pod**: 1D

## 1. Goal

Expose openpyxl's workbook-level protection: lock-structure,
lock-windows, revision tracking passwords, file-sharing
read-only-recommended flag.

## 2. Public API

```python
from wolfxl.workbook.protection import WorkbookProtection, FileSharing

wb.security = WorkbookProtection(
    workbook_password="hunter2",
    lock_structure=True,
    lock_windows=False,
    lock_revision=False,
)
wb.security.revisions_password = "review-only"

wb.fileSharing = FileSharing(
    read_only_recommended=True,
    user_name="alice",
    reservation_password="alice-pw",
)
```

## 3. Class definitions

```python
class WorkbookProtection:
    workbook_password: str | None = None             # plaintext setter
    workbook_password_character_set: str | None = None
    lock_structure: bool = False
    lock_windows: bool = False
    lock_revision: bool = False
    revisions_algorithm_name: str | None = None
    revisions_hash_value: str | None = None
    revisions_salt_value: str | None = None
    revisions_spin_count: int | None = None
    workbook_algorithm_name: str | None = None
    workbook_hash_value: str | None = None
    workbook_salt_value: str | None = None
    workbook_spin_count: int | None = None

    def set_workbook_password(self, plaintext: str,
                              algorithm: str = "SHA-512") -> None: ...
    def set_revisions_password(self, plaintext: str,
                               algorithm: str = "SHA-512") -> None: ...
    def check_workbook_password(self, plaintext: str) -> bool: ...
    def check_revisions_password(self, plaintext: str) -> bool: ...

class FileSharing:
    read_only_recommended: bool = False
    user_name: str | None = None
    reservation_password: str | None = None        # plaintext setter
    algorithm_name: str | None = None
    hash_value: str | None = None
    salt_value: str | None = None
    spin_count: int | None = None

    def set_reservation_password(self, plaintext: str,
                                 algorithm: str = "SHA-512") -> None: ...
```

## 4. Password hashing

Reuses `wolfxl.utils.protection.hash_password` from Sprint Ι
Pod-γ. Default algorithm SHA-512 with 100,000 spin count
(matches Excel's default). Hash output base64-encoded.

When the user calls `set_workbook_password("plaintext")`:
1. Generate 16-byte salt.
2. Run SHA-512 hash with spin_count iterations.
3. Set `workbook_algorithm_name="SHA-512"`,
   `workbook_hash_value=<base64>`,
   `workbook_salt_value=<base64-of-salt>`,
   `workbook_spin_count=100000`.

## 5. OOXML output

```xml
<workbook>
  ...
  <bookViews>...</bookViews>

  <workbookProtection
    workbookAlgorithmName="SHA-512"
    workbookHashValue="..."
    workbookSaltValue="..."
    workbookSpinCount="100000"
    lockStructure="1"
    lockWindows="0"/>

  <fileSharing
    readOnlyRecommended="1"
    userName="alice"
    algorithmName="SHA-512"
    hashValue="..."
    saltValue="..."
    spinCount="100000"/>

  ...
</workbook>
```

`<workbookProtection>` and `<fileSharing>` are siblings of
`<bookViews>` per CT_Workbook child order:

```
fileVersion → fileSharing → workbookPr → workbookProtection
→ bookViews → sheets → ...
```

Note: `<fileSharing>` must come BEFORE `<workbookPr>`;
`<workbookProtection>` must come AFTER `<workbookPr>` but
BEFORE `<bookViews>`. wolfxl-merger gets two new
`WorkbookBlock` variants ordered correctly.

## 6. Modify mode

XlsxPatcher Phase 2.5q (after 2.5p slicers, before 2.5h
defined-names): drains `Workbook.security` / `fileSharing` and
splices into `xl/workbook.xml`.

## 7. Native writer

`crates/wolfxl-writer/src/emit/workbook.rs` extended to emit
the two new blocks at canonical positions. ~150 LOC delta.

## 8. Testing

- `tests/test_workbook_protection.py` (~10 tests).
- `tests/test_file_sharing.py` (~6 tests).
- `tests/test_workbook_password_hash.py` (~5 tests verifying
  hash matches openpyxl output for fixed input).
- `tests/parity/test_workbook_security_parity.py` (~4 tests).

## 9. References

- ECMA-376 Part 1 §18.2.29 (CT_WorkbookProtection)
- ECMA-376 Part 1 §18.2.10 (CT_FileSharing)
- openpyxl 3.1.x `openpyxl.workbook.protection` source.

## 10. Dict contract

`Workbook.to_rust_security_dict()`:

```python
{
    "workbook_protection": {
        "lock_structure": bool,
        "lock_windows": bool,
        "lock_revision": bool,
        "workbook_algorithm_name": str | None,
        "workbook_hash_value": str | None,
        "workbook_salt_value": str | None,
        "workbook_spin_count": int | None,
        "revisions_algorithm_name": str | None,
        "revisions_hash_value": str | None,
        "revisions_salt_value": str | None,
        "revisions_spin_count": int | None,
    } | None,
    "file_sharing": {
        "read_only_recommended": bool,
        "user_name": str | None,
        "algorithm_name": str | None,
        "hash_value": str | None,
        "salt_value": str | None,
        "spin_count": int | None,
    } | None,
}
```

PyO3 binding: `serialize_workbook_security_dict(d) -> bytes`
returns the XML fragments (two separate fragments — caller
splices each at canonical position).

## 11. Acceptance

- `wb.security.set_workbook_password("hunter2")` round-trips.
- openpyxl reads back the same hash value.
- LibreOffice prompts for password on open.
- ~25 tests green.
