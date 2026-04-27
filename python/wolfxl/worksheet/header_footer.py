"""Header / footer classes for worksheets (RFC-055 §2.3).

Backs ``ws.header_footer``. Includes the OOXML format-code grammar
validator (``&L`` / ``&C`` / ``&R`` / ``&P`` / ``&N`` / ``&D`` /
``&T`` / ``&F`` / ``&A`` / ``&Z`` / ``&K{RRGGBB}`` / ``&"font,style"``
/ ``&NN`` / ``&B`` / ``&I`` / ``&U`` / ``&S`` / ``&X`` / ``&Y`` /
``&&``).
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any


# Recognized single-letter codes (case-sensitive in OOXML).
_FORMAT_CODES_SINGLE = set("LCRPNDTFAZGBIUSXY")

# Recognized format-code grammar regex — used for validation only;
# unknown ``&X`` codes pass through unchanged for forward compat.
# Pattern explanation:
#   &&                     -> literal ampersand
#   &"font,style"          -> font select (anything between quotes)
#   &K{RRGGBB}             -> color hex
#   &<digits>              -> font size
#   &[A-Z]                 -> single-letter code
_FORMAT_CODE_RE = re.compile(
    r"""
    &&                                  # literal &
    | &"[^"]*"                          # &"font,style"
    | &K[0-9A-Fa-f]{6}                  # &Kffrrggbb (case-insensitive hex)
    | &K\{[0-9A-Fa-f]{6}\}              # &K{RRGGBB} (alternate form)
    | &\d+                              # &NN font size
    | &[A-Za-z]                         # single-letter code
    """,
    re.VERBOSE,
)


def validate_header_footer_format(text: str) -> bool:
    """Return True iff ``text`` parses as a valid OOXML header/footer string.

    "Valid" means: every ``&`` is followed by a recognized code or is
    a ``&&`` literal escape. Unknown single-letter codes (e.g. invented
    by a later spec revision) are treated as opaque — they pass through
    unchanged for forward compat.

    The function never raises; it returns False for clearly malformed
    inputs (a trailing bare ``&`` with no code character, an unclosed
    quoted font block).
    """
    if text is None:
        return True
    i = 0
    n = len(text)
    while i < n:
        ch = text[i]
        if ch != "&":
            i += 1
            continue
        # Consume an ampersand sequence.
        if i + 1 >= n:
            return False  # trailing bare '&'
        nxt = text[i + 1]
        if nxt == "&":
            i += 2
            continue
        if nxt == '"':
            # Quoted font: &"font,style"
            close = text.find('"', i + 2)
            if close == -1:
                return False
            i = close + 1
            continue
        if nxt in ("K", "k"):
            # &Kxxxxxx or &K{xxxxxx}
            if i + 2 >= n:
                return False
            if text[i + 2] == "{":
                if i + 9 >= n or text[i + 9] != "}":
                    return False
                if not all(c in "0123456789abcdefABCDEF" for c in text[i + 3:i + 9]):
                    return False
                i += 10
                continue
            # &Kxxxxxx (6-hex form)
            if i + 8 > n:
                return False
            if not all(c in "0123456789abcdefABCDEF" for c in text[i + 2:i + 8]):
                return False
            i += 8
            continue
        if nxt.isdigit():
            # &NN (font size — one or more digits)
            j = i + 1
            while j < n and text[j].isdigit():
                j += 1
            i = j
            continue
        if nxt.isalpha():
            # Single-letter code (known or forward-compat unknown).
            i += 2
            continue
        # Bare '&' followed by punctuation/whitespace — not a valid code.
        return False
    return True


@dataclass
class HeaderFooterItem:
    """One header or footer (left / center / right segments)."""

    left: str | None = None
    center: str | None = None
    right: str | None = None

    @property
    def text(self) -> str:
        """Compose the segments back into the OOXML text form.

        Used by the emitter and by tests that compare against
        openpyxl's serialization.
        """
        parts: list[str] = []
        if self.left is not None:
            parts.append("&L" + self.left)
        if self.center is not None:
            parts.append("&C" + self.center)
        if self.right is not None:
            parts.append("&R" + self.right)
        return "".join(parts)

    def is_empty(self) -> bool:
        return self.left is None and self.center is None and self.right is None

    def to_rust_dict(self) -> dict[str, Any] | None:
        if self.is_empty():
            return None
        return {
            "left": self.left,
            "center": self.center,
            "right": self.right,
        }

    def __post_init__(self) -> None:
        for name in ("left", "center", "right"):
            v = getattr(self, name)
            if v is not None and not validate_header_footer_format(v):
                raise ValueError(
                    f"HeaderFooterItem.{name}: invalid format code in {v!r}"
                )


@dataclass
class HeaderFooter:
    """Complete CT_HeaderFooter (ECMA-376 §18.3.1.36)."""

    odd_header: HeaderFooterItem = field(default_factory=HeaderFooterItem)
    odd_footer: HeaderFooterItem = field(default_factory=HeaderFooterItem)
    even_header: HeaderFooterItem = field(default_factory=HeaderFooterItem)
    even_footer: HeaderFooterItem = field(default_factory=HeaderFooterItem)
    first_header: HeaderFooterItem = field(default_factory=HeaderFooterItem)
    first_footer: HeaderFooterItem = field(default_factory=HeaderFooterItem)
    different_odd_even: bool = False
    different_first: bool = False
    scale_with_doc: bool = True
    align_with_margins: bool = True

    # openpyxl aliases — `oddHeader` etc. (camelCase mirrors)
    @property
    def oddHeader(self) -> HeaderFooterItem:  # noqa: N802
        return self.odd_header

    @property
    def oddFooter(self) -> HeaderFooterItem:  # noqa: N802
        return self.odd_footer

    @property
    def evenHeader(self) -> HeaderFooterItem:  # noqa: N802
        return self.even_header

    @property
    def evenFooter(self) -> HeaderFooterItem:  # noqa: N802
        return self.even_footer

    @property
    def firstHeader(self) -> HeaderFooterItem:  # noqa: N802
        return self.first_header

    @property
    def firstFooter(self) -> HeaderFooterItem:  # noqa: N802
        return self.first_footer

    def is_default(self) -> bool:
        return (
            self.odd_header.is_empty()
            and self.odd_footer.is_empty()
            and self.even_header.is_empty()
            and self.even_footer.is_empty()
            and self.first_header.is_empty()
            and self.first_footer.is_empty()
            and not self.different_odd_even
            and not self.different_first
            and self.scale_with_doc
            and self.align_with_margins
        )

    def to_rust_dict(self) -> dict[str, Any]:
        return {
            "odd_header": self.odd_header.to_rust_dict(),
            "odd_footer": self.odd_footer.to_rust_dict(),
            "even_header": self.even_header.to_rust_dict(),
            "even_footer": self.even_footer.to_rust_dict(),
            "first_header": self.first_header.to_rust_dict(),
            "first_footer": self.first_footer.to_rust_dict(),
            "different_odd_even": bool(self.different_odd_even),
            "different_first": bool(self.different_first),
            "scale_with_doc": bool(self.scale_with_doc),
            "align_with_margins": bool(self.align_with_margins),
        }


__all__ = [
    "HeaderFooter",
    "HeaderFooterItem",
    "validate_header_footer_format",
]
