"""``openpyxl.worksheet.merge`` — merge-cell value types.

Wolfxl tracks merged ranges as plain A1 strings on the Worksheet
proxy; this module surfaces the openpyxl-shaped value types so user
code that constructs them by hand (or ``isinstance``-checks against
them) ports mechanically.

Pod 2 (RFC-060 §2.1).
"""

from __future__ import annotations

from wolfxl._compat import _make_stub
from wolfxl.cell._merged import MergedCell
from wolfxl.worksheet.cell_range import CellRange


class MergedCellRange(CellRange):
    """A :class:`CellRange` flagged as a merged region.

    openpyxl stores merged regions as ``MergedCellRange`` instances
    (a ``CellRange`` subclass) on ``ws.merged_cells``.  Wolfxl uses
    plain strings, but the class is exposed here so user code that
    constructs one explicitly continues to work.
    """


# Container shims — wolfxl exposes ``ws.merged_cells`` as a plain set
# of strings, so these are stubs.
MergeCell = _make_stub(
    "MergeCell",
    "openpyxl's MergeCell is wolfxl's plain A1 string in ``ws.merged_cells``; "
    "use ``ws.merge_cells(range_string)`` to register one.",
)
MergeCells = _make_stub(
    "MergeCells",
    "openpyxl's MergeCells container is wolfxl's ``ws.merged_cells`` set.",
)


__all__ = ["MergeCell", "MergeCells", "MergedCell", "MergedCellRange"]
