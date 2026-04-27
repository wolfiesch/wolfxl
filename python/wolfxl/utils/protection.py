"""Password hashing for sheet/workbook protection (RFC-055 / RFC-058).

The Excel "Legacy Password Hash" algorithm is the historical 16-bit
hash openpyxl emits in ``<sheetProtection password="…">`` and
``<workbookProtection workbookPassword="…">`` attributes. It is NOT
cryptographically secure — it's a UI-level deterrent.

The algorithm matches openpyxl.utils.protection.hash_password byte for
byte; tests in tests/test_sheet_protection.py round-trip ``"hunter2"``
against openpyxl's output to pin the contract.
"""

from __future__ import annotations


def hash_password(plaintext: str) -> str:
    """Compute the legacy 16-bit Excel password hash.

    Returns a 4-char uppercase hex string (no ``0x`` prefix). Empty
    input returns the literal string ``""`` so the caller can
    round-trip "no password" without an extra branch.

    Cribbed from openpyxl.utils.protection (the algorithm has been
    public since Excel 95). See also Trail of Bits' "Microsoft Office
    file password recovery" write-up — same arithmetic, different
    serialization.
    """
    if not plaintext:
        return ""
    password = 0x0000
    for idx, ch in enumerate(plaintext):
        char_code = ord(ch)
        # Spec: rotate left by (idx + 1) within a 15-bit field.
        rotated_bits = (idx + 1) % 15
        # 15-bit window per the algorithm — the high bit is XOR'd back
        # into the low 15 bits.
        rotated = ((char_code << rotated_bits) | (char_code >> (15 - rotated_bits))) & 0x7FFF
        password ^= rotated
    password ^= len(plaintext)
    password ^= 0xCE4B
    return f"{password:04X}"


def check_password(plaintext: str, expected_hash: str) -> bool:
    """Verify ``plaintext`` matches a previously-stored hash.

    Returns False on empty/None input regardless of the stored hash —
    "no password supplied" never matches.
    """
    if not plaintext:
        return False
    return hash_password(plaintext) == (expected_hash or "").upper()


__all__ = ["hash_password", "check_password"]
