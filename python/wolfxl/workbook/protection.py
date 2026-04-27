"""Workbook-level security classes — RFC-058.

Two openpyxl-shaped containers:

* :class:`WorkbookProtection` — toggles structure / window / revision
  locks plus optional SHA-512 spin-hashed passwords for the
  ``workbookPassword`` and ``revisionsPassword`` slots.
* :class:`FileSharing` — read-only-recommended flag plus optional
  user name and reservation password.

Both classes match openpyxl's attribute names (camelCase + snake_case
aliases) so existing code using ``wb.security`` / ``wb.fileSharing``
continues to work. Hash algorithm defaults to ``"SHA-512"`` with a
spin count of 100,000 (Excel 2013+ default).
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from wolfxl.utils.protection import (
    hash_password,
    hash_password_sha512,
    verify_password_sha512,
)

if TYPE_CHECKING:
    pass


# ---------------------------------------------------------------------------
# WorkbookProtection
# ---------------------------------------------------------------------------


class WorkbookProtection:
    """Workbook-level protection block (``<workbookProtection>``).

    Two independent password slots:

    * **Workbook password** — protects the workbook structure (sheet
      add/remove/reorder, hidden-sheet visibility). Set via
      :meth:`set_workbook_password`.
    * **Revisions password** — protects revision tracking history.
      Set via :meth:`set_revisions_password`.

    Each password is stored as a SHA-512 spin hash plus a salt; the
    plaintext is never persisted. Passing ``workbook_password=`` /
    ``revisions_password=`` to the constructor is a convenience that
    routes through the corresponding ``set_*_password`` method.

    The three lock booleans (``lock_structure``, ``lock_windows``,
    ``lock_revision``) work independently of the password slots —
    a True flag with no password means "locked but anyone can unlock".
    """

    __slots__ = (
        "lock_structure",
        "lock_windows",
        "lock_revision",
        "workbook_password",
        "workbook_password_character_set",
        "workbook_algorithm_name",
        "workbook_hash_value",
        "workbook_salt_value",
        "workbook_spin_count",
        "revisions_password",
        "revisions_password_character_set",
        "revisions_algorithm_name",
        "revisions_hash_value",
        "revisions_salt_value",
        "revisions_spin_count",
    )

    def __init__(
        self,
        *,
        workbook_password: str | None = None,
        revisions_password: str | None = None,
        lock_structure: bool = False,
        lock_windows: bool = False,
        lock_revision: bool = False,
        workbook_password_character_set: str | None = None,
        workbook_algorithm_name: str | None = None,
        workbook_hash_value: str | None = None,
        workbook_salt_value: str | None = None,
        workbook_spin_count: int | None = None,
        revisions_password_character_set: str | None = None,
        revisions_algorithm_name: str | None = None,
        revisions_hash_value: str | None = None,
        revisions_salt_value: str | None = None,
        revisions_spin_count: int | None = None,
    ) -> None:
        self.lock_structure: bool = bool(lock_structure)
        self.lock_windows: bool = bool(lock_windows)
        self.lock_revision: bool = bool(lock_revision)

        self.workbook_password: str | None = None
        self.workbook_password_character_set: str | None = workbook_password_character_set
        self.workbook_algorithm_name: str | None = workbook_algorithm_name
        self.workbook_hash_value: str | None = workbook_hash_value
        self.workbook_salt_value: str | None = workbook_salt_value
        self.workbook_spin_count: int | None = workbook_spin_count

        self.revisions_password: str | None = None
        self.revisions_password_character_set: str | None = revisions_password_character_set
        self.revisions_algorithm_name: str | None = revisions_algorithm_name
        self.revisions_hash_value: str | None = revisions_hash_value
        self.revisions_salt_value: str | None = revisions_salt_value
        self.revisions_spin_count: int | None = revisions_spin_count

        if workbook_password is not None:
            self.set_workbook_password(workbook_password)
        if revisions_password is not None:
            self.set_revisions_password(revisions_password)

    # ------------------------------------------------------------------
    # Workbook password
    # ------------------------------------------------------------------

    def set_workbook_password(
        self,
        plaintext: str,
        algorithm: str = "SHA-512",
        *,
        salt: bytes | None = None,
        spin_count: int = 100_000,
    ) -> None:
        """Hash ``plaintext`` and store it on the workbook-password slot.

        After this call ``workbook_algorithm_name``,
        ``workbook_hash_value``, ``workbook_salt_value``, and
        ``workbook_spin_count`` are all populated. The legacy
        ``workbook_password`` attribute is also set to the openpyxl-style
        16-bit XOR hash so byte-equal round-trips against openpyxl-saved
        files keep working.
        """
        h, s = hash_password_sha512(
            plaintext,
            salt=salt,
            spin_count=spin_count,
            algorithm=algorithm,
        )
        self.workbook_algorithm_name = algorithm
        self.workbook_hash_value = h
        self.workbook_salt_value = s
        self.workbook_spin_count = spin_count
        # The legacy password attribute mirrors openpyxl: the
        # short XOR hash, NOT the plaintext.
        self.workbook_password = hash_password(plaintext)

    def check_workbook_password(self, plaintext: str) -> bool:
        """Return ``True`` iff ``plaintext`` matches the stored hash.

        Uses the SHA-512 spin hash; the legacy ``workbook_password``
        attribute is ignored. Returns ``False`` when no SHA-512 hash
        is present (i.e. :meth:`set_workbook_password` was never
        called).
        """
        if (
            self.workbook_hash_value is None
            or self.workbook_salt_value is None
            or self.workbook_algorithm_name is None
            or self.workbook_spin_count is None
        ):
            return False
        return verify_password_sha512(
            plaintext,
            self.workbook_hash_value,
            self.workbook_salt_value,
            spin_count=self.workbook_spin_count,
            algorithm=self.workbook_algorithm_name,
        )

    # ------------------------------------------------------------------
    # Revisions password
    # ------------------------------------------------------------------

    def set_revisions_password(
        self,
        plaintext: str,
        algorithm: str = "SHA-512",
        *,
        salt: bytes | None = None,
        spin_count: int = 100_000,
    ) -> None:
        """Hash ``plaintext`` and store it on the revisions-password slot."""
        h, s = hash_password_sha512(
            plaintext,
            salt=salt,
            spin_count=spin_count,
            algorithm=algorithm,
        )
        self.revisions_algorithm_name = algorithm
        self.revisions_hash_value = h
        self.revisions_salt_value = s
        self.revisions_spin_count = spin_count
        self.revisions_password = hash_password(plaintext)

    def check_revisions_password(self, plaintext: str) -> bool:
        """Return ``True`` iff ``plaintext`` matches the stored hash."""
        if (
            self.revisions_hash_value is None
            or self.revisions_salt_value is None
            or self.revisions_algorithm_name is None
            or self.revisions_spin_count is None
        ):
            return False
        return verify_password_sha512(
            plaintext,
            self.revisions_hash_value,
            self.revisions_salt_value,
            spin_count=self.revisions_spin_count,
            algorithm=self.revisions_algorithm_name,
        )

    # ------------------------------------------------------------------
    # Dict serialisation (RFC-058 §10)
    # ------------------------------------------------------------------

    def to_dict(self) -> dict[str, object]:
        """Return the patcher/writer-side flat dict (RFC-058 §10)."""
        return {
            "lock_structure": self.lock_structure,
            "lock_windows": self.lock_windows,
            "lock_revision": self.lock_revision,
            "workbook_algorithm_name": self.workbook_algorithm_name,
            "workbook_hash_value": self.workbook_hash_value,
            "workbook_salt_value": self.workbook_salt_value,
            "workbook_spin_count": self.workbook_spin_count,
            "revisions_algorithm_name": self.revisions_algorithm_name,
            "revisions_hash_value": self.revisions_hash_value,
            "revisions_salt_value": self.revisions_salt_value,
            "revisions_spin_count": self.revisions_spin_count,
        }

    def __repr__(self) -> str:  # pragma: no cover - debug aid
        return (
            "WorkbookProtection("
            f"lock_structure={self.lock_structure}, "
            f"lock_windows={self.lock_windows}, "
            f"lock_revision={self.lock_revision}, "
            f"workbook_password_set={self.workbook_hash_value is not None}, "
            f"revisions_password_set={self.revisions_hash_value is not None})"
        )


# ---------------------------------------------------------------------------
# FileSharing
# ---------------------------------------------------------------------------


class FileSharing:
    """``<fileSharing>`` block — read-only-recommended + reservation password.

    The ``read_only_recommended`` flag suggests Excel show the workbook
    in read-only mode by default. The reservation password (set via
    :meth:`set_reservation_password`) gates write access — without it,
    Excel falls back to read-only.
    """

    __slots__ = (
        "read_only_recommended",
        "user_name",
        "reservation_password",
        "algorithm_name",
        "hash_value",
        "salt_value",
        "spin_count",
    )

    def __init__(
        self,
        *,
        read_only_recommended: bool = False,
        user_name: str | None = None,
        reservation_password: str | None = None,
        algorithm_name: str | None = None,
        hash_value: str | None = None,
        salt_value: str | None = None,
        spin_count: int | None = None,
    ) -> None:
        self.read_only_recommended: bool = bool(read_only_recommended)
        self.user_name: str | None = user_name
        self.reservation_password: str | None = None
        self.algorithm_name: str | None = algorithm_name
        self.hash_value: str | None = hash_value
        self.salt_value: str | None = salt_value
        self.spin_count: int | None = spin_count

        if reservation_password is not None:
            self.set_reservation_password(reservation_password)

    def set_reservation_password(
        self,
        plaintext: str,
        algorithm: str = "SHA-512",
        *,
        salt: bytes | None = None,
        spin_count: int = 100_000,
    ) -> None:
        """Hash ``plaintext`` and store it on the reservation-password slot."""
        h, s = hash_password_sha512(
            plaintext,
            salt=salt,
            spin_count=spin_count,
            algorithm=algorithm,
        )
        self.algorithm_name = algorithm
        self.hash_value = h
        self.salt_value = s
        self.spin_count = spin_count
        self.reservation_password = hash_password(plaintext)

    def check_reservation_password(self, plaintext: str) -> bool:
        """Return ``True`` iff ``plaintext`` matches the stored hash."""
        if (
            self.hash_value is None
            or self.salt_value is None
            or self.algorithm_name is None
            or self.spin_count is None
        ):
            return False
        return verify_password_sha512(
            plaintext,
            self.hash_value,
            self.salt_value,
            spin_count=self.spin_count,
            algorithm=self.algorithm_name,
        )

    def to_dict(self) -> dict[str, object]:
        """Return the patcher/writer-side flat dict (RFC-058 §10)."""
        return {
            "read_only_recommended": self.read_only_recommended,
            "user_name": self.user_name,
            "algorithm_name": self.algorithm_name,
            "hash_value": self.hash_value,
            "salt_value": self.salt_value,
            "spin_count": self.spin_count,
        }

    def __repr__(self) -> str:  # pragma: no cover - debug aid
        return (
            "FileSharing("
            f"read_only_recommended={self.read_only_recommended}, "
            f"user_name={self.user_name!r}, "
            f"reservation_password_set={self.hash_value is not None})"
        )


__all__ = ["WorkbookProtection", "FileSharing"]
