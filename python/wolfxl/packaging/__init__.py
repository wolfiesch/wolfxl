"""openpyxl.packaging compatibility.

Re-exports :class:`DocumentProperties` from ``core`` so users can write
``from wolfxl.packaging.core import DocumentProperties`` or the shorter
``from wolfxl.packaging import DocumentProperties``.
"""

from __future__ import annotations

from wolfxl.packaging.core import DocumentProperties

__all__ = ["DocumentProperties"]
