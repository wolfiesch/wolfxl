"""Shim for ``openpyxl.worksheet`` - subpackage re-exports only.

openpyxl organizes worksheet-adjacent classes (DataValidation, Table,
Hyperlink) under this package. wolfxl exposes the same module paths so
imports succeed, with stub classes that raise on instantiation.
"""

from __future__ import annotations

__all__: list[str] = []
