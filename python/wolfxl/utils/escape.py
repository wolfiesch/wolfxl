"""``openpyxl.utils.escape`` — XML-string escape helpers.

Mirrors openpyxl's ``openpyxl/utils/escape.py``.  These wrap Python's
:mod:`xml.sax.saxutils` for ergonomic call sites (``escape("<")`` →
``"&lt;"``).  The Rust XML emitter handles its own escaping internally;
this module exists for openpyxl-path import parity.

Pod 2 (RFC-060).
"""

from __future__ import annotations

import re
from xml.sax.saxutils import escape as _xml_escape, unescape as _xml_unescape


# openpyxl's regex for OOXML ``_xHHHH_`` escapes — restoring control characters
# that the writer encoded to dodge XML's "no C0 controls" rule.
_ESCAPE_RE = re.compile(r"_x([0-9A-Fa-f]{4})_")


def escape(value: str) -> str:
    """XML-escape ``&``, ``<``, ``>``."""
    return _xml_escape(value)


def unescape(value: str) -> str:
    """Reverse :func:`escape` and resolve OOXML ``_xHHHH_`` controls."""

    def _replace(match: re.Match[str]) -> str:
        return chr(int(match.group(1), 16))

    return _xml_unescape(_ESCAPE_RE.sub(_replace, value))


__all__ = ["escape", "unescape"]
