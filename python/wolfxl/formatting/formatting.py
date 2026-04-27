"""``openpyxl.formatting.formatting`` — re-export shim.

Surfaces :class:`ConditionalFormatting` + :class:`ConditionalFormattingList`
under the openpyxl-shaped path.  The canonical home is
:mod:`wolfxl.formatting` (the package init).

Pod 2 (RFC-060 §2.4).
"""

from __future__ import annotations

from wolfxl.formatting import ConditionalFormatting, ConditionalFormattingList

__all__ = ["ConditionalFormatting", "ConditionalFormattingList"]
