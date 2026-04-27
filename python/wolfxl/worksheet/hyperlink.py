"""openpyxl.worksheet.hyperlink compatibility.

T1 makes ``Hyperlink`` a real dataclass. Reads work from any file; writes
via ``cell.hyperlink = Hyperlink(...)`` land in write mode (T1 PR4).
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass
class Hyperlink:
    """A cell hyperlink.

    ``target`` holds an external URL (http:// / mailto: / file://).
    ``location`` holds an internal reference (``Sheet1!A1``) for intra-
    workbook links. ``display`` is the visible text that overrides the
    cell value, ``tooltip`` the screen-tip on hover. ``id`` is the rel id
    assigned by the writer — read-only from Python.
    """

    ref: str | None = None
    target: str | None = None
    location: str | None = None
    tooltip: str | None = None
    display: str | None = None
    id: str | None = None


class HyperlinkList(list):
    """``openpyxl``-shaped list container for :class:`Hyperlink` entries.

    openpyxl exposes ``HyperlinkList`` as the type backing
    ``ws.hyperlinks``.  Wolfxl tracks hyperlinks on individual cell
    proxies, so this container is used purely to satisfy
    ``isinstance(ws.hyperlinks, HyperlinkList)`` migrations — it
    behaves as a plain :class:`list` of hyperlinks.

    Pod 2 (RFC-060).
    """


__all__ = ["Hyperlink", "HyperlinkList"]
