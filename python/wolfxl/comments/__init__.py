"""openpyxl.comments compatibility.

T1 makes ``Comment`` a real, mutable dataclass. Construction works in any
mode; attaching to a cell via ``cell.comment = Comment(...)`` works in
write mode (T1 PR4). Read access — ``cell.comment.text`` — works on any
file opened in read or modify mode.
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass
class Comment:
    """A cell comment (note).

    openpyxl keeps comments mutable — users commonly do
    ``cell.comment.text = "updated"`` after attaching. Width/height are
    preserved on round-trip but not authored from Python; wolfxl stores
    them as pass-throughs.
    """

    text: str
    author: str | None = None
    height: int | None = None
    width: int | None = None

    @property
    def content(self) -> str:
        return self.text

    @content.setter
    def content(self, value: str) -> None:
        self.text = value


__all__ = ["Comment"]
