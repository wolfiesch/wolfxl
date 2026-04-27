"""``openpyxl.comments.comments`` — re-export shim for :class:`Comment`.

Pod 2 (RFC-060).
"""

from __future__ import annotations

from wolfxl._compat import _make_stub
from wolfxl.comments import Comment


CommentSheet = _make_stub(
    "CommentSheet",
    "Wolfxl serializes comment sheets internally; CommentSheet is not "
    "exposed for direct construction.",
)


__all__ = ["Comment", "CommentSheet"]
