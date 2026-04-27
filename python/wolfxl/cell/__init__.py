"""openpyxl-compatible ``wolfxl.cell`` namespace.

Re-exports the rich-text shims from :mod:`wolfxl.cell.rich_text`
plus the cell-class compatibility shims (:class:`MergedCell`,
:class:`WriteOnlyCell`) and :class:`IllegalCharacterError`.  The
package shape mirrors openpyxl's ``openpyxl.cell`` so existing
code that imports from ``openpyxl.cell.*`` can be redirected to
``wolfxl.cell.*`` with one-line import swaps.

Pod 2 (RFC-060) owns the openpyxl-shaped path shims (e.g.
``wolfxl.cell.cell``) — this module exposes the classes at their
natural locations and Pod 2 builds the import paths on top.
"""

from wolfxl.cell import rich_text
from wolfxl.cell._merged import MergedCell
from wolfxl.cell._write_only import WriteOnlyCell
from wolfxl.utils.exceptions import IllegalCharacterError

__all__ = [
    "IllegalCharacterError",
    "MergedCell",
    "WriteOnlyCell",
    "rich_text",
]
