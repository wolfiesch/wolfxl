"""``openpyxl.worksheet.merge`` — merge-cell value types.

Wolfxl tracks merged ranges as plain A1 strings on the Worksheet
proxy; this module surfaces the openpyxl-shaped value types so user
code that constructs them by hand (or ``isinstance``-checks against
them) ports mechanically.

Sprint Π Pod-β (RFC-063) replaced the ``MergeCell`` and ``MergeCells``
stubs with real type-only proxies over the underlying ``set[str]``.
"""

from __future__ import annotations

from collections.abc import Iterable, Iterator
from dataclasses import dataclass
from typing import Any

from wolfxl.cell._merged import MergedCell
from wolfxl.worksheet.cell_range import CellRange


class MergedCellRange(CellRange):
    """A :class:`CellRange` flagged as a merged region.

    openpyxl stores merged regions as ``MergedCellRange`` instances
    (a ``CellRange`` subclass) on ``ws.merged_cells``.  Wolfxl uses
    plain strings, but the class is exposed here so user code that
    constructs one explicitly continues to work.
    """


@dataclass
class MergeCell:
    """Single merged region (CT_MergeCell §18.3.1.55).

    Wolfxl stores merged ranges as plain A1 strings; this class
    provides the openpyxl-shaped type-wrapper so user code
    constructing a ``MergeCell`` continues to work.
    """

    ref: str

    @property
    def coord(self) -> str:
        """Alias for :attr:`ref` — openpyxl spells it both ways."""
        return self.ref

    def __str__(self) -> str:  # pragma: no cover - trivial
        return self.ref


class MergeCells:
    """Container for :class:`MergeCell` entries (CT_MergeCells §18.3.1.56).

    Backed by ``ws.merged_cells`` (a plain ``set[str]``) when bound to
    a worksheet; otherwise backed by an in-memory list. The container
    surfaces ``__iter__`` / ``__len__`` / :attr:`count` / :meth:`append`
    / :meth:`remove` for openpyxl source compatibility.

    The wrapper is a *view* — mutations are mirrored back onto the
    worksheet's underlying set so the existing patcher / native-writer
    pipelines see them automatically.
    """

    __slots__ = ("worksheet", "_extra")

    def __init__(
        self,
        worksheet: Any = None,
        mergeCell: Iterable[Any] | None = None,  # noqa: N803 - openpyxl API
    ) -> None:
        self.worksheet = worksheet
        # Items added via the constructor when no worksheet is provided
        # (or that are explicitly stuffed via ``append``) live here when
        # the worksheet path isn't available.  When a worksheet *is*
        # provided, the constructor entries are folded into the
        # underlying ``ws.merged_cells`` set so the two views stay in
        # sync.
        self._extra: list[MergeCell] = []
        if mergeCell is not None:
            for item in mergeCell:
                self.append(item)

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _ws_set(self) -> set[str] | None:
        """Return the underlying ``ws._merged_ranges`` set, or ``None``."""
        ws = self.worksheet
        if ws is None:
            return None
        return getattr(ws, "_merged_ranges", None)

    @staticmethod
    def _coerce(item: Any) -> MergeCell:
        """Coerce *item* into a :class:`MergeCell`."""
        if isinstance(item, MergeCell):
            return item
        if isinstance(item, str):
            return MergeCell(ref=item)
        # Fall back to an attribute lookup so ``CellRange``-like inputs
        # (which carry a stringifiable form) work transparently.
        ref = getattr(item, "coord", None) or getattr(item, "ref", None)
        if ref is None:
            ref = str(item)
        return MergeCell(ref=str(ref))

    # ------------------------------------------------------------------
    # openpyxl-shape API
    # ------------------------------------------------------------------

    @property
    def mergeCell(self) -> list[MergeCell]:  # noqa: N802 - openpyxl API
        """Materialised list of :class:`MergeCell` entries (snapshot)."""
        return list(iter(self))

    @property
    def count(self) -> int:
        """Number of merged regions currently registered."""
        return len(self)

    def append(self, mc: Any) -> None:
        """Add a merged region (accepts a :class:`MergeCell` or A1 string)."""
        cell = self._coerce(mc)
        backing = self._ws_set()
        if backing is not None:
            backing.add(cell.ref)
        else:
            # Avoid duplicates when no worksheet is bound — match the
            # set-like semantics of the live path.
            if not any(existing.ref == cell.ref for existing in self._extra):
                self._extra.append(cell)

    def remove(self, mc: Any) -> None:
        """Remove a merged region (accepts a :class:`MergeCell` or A1 string)."""
        cell = self._coerce(mc)
        backing = self._ws_set()
        if backing is not None:
            backing.discard(cell.ref)
        else:
            self._extra = [existing for existing in self._extra if existing.ref != cell.ref]

    def __iter__(self) -> Iterator[MergeCell]:
        backing = self._ws_set()
        if backing is not None:
            # Sort for deterministic iteration — openpyxl's MultiCellRange
            # also yields refs in a canonical order.
            for ref in sorted(backing):
                yield MergeCell(ref=ref)
        else:
            yield from self._extra

    def __len__(self) -> int:
        backing = self._ws_set()
        if backing is not None:
            return len(backing)
        return len(self._extra)

    def __contains__(self, item: Any) -> bool:
        cell = self._coerce(item)
        backing = self._ws_set()
        if backing is not None:
            return cell.ref in backing
        return any(existing.ref == cell.ref for existing in self._extra)

    def __repr__(self) -> str:  # pragma: no cover - trivial
        refs = [c.ref for c in self]
        return f"MergeCells(count={self.count}, refs={refs!r})"


__all__ = ["MergeCell", "MergeCells", "MergedCell", "MergedCellRange"]
