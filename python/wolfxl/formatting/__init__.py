"""openpyxl.formatting compatibility.

Exposes ``ConditionalFormatting`` (one range + its rules) and
``ConditionalFormattingList`` (the ws.conditional_formatting container).
"""

from __future__ import annotations

from collections.abc import Iterator
from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from wolfxl.formatting.rule import Rule

if TYPE_CHECKING:
    from wolfxl._worksheet import Worksheet


@dataclass
class ConditionalFormatting:
    """All CF rules that apply to a single range.

    openpyxl groups rules by range: one ``ConditionalFormatting`` per
    distinct ``sqref``. ``cfRule`` is the legacy openpyxl alias for
    ``rules`` — both return the same list.
    """

    sqref: str
    rules: list[Rule] = field(default_factory=list)

    @property
    def cfRule(self) -> list[Rule]:  # noqa: N802 - openpyxl alias
        return self.rules


class ConditionalFormattingList:
    """Container for a worksheet's conditional formatting entries.

    Iterates ``ConditionalFormatting`` objects. In write mode, users
    attach new CF rules via ``ws.conditional_formatting.add(range, rule)``
    — that lands in PR5. Reads work in any mode.
    """

    __slots__ = ("_entries", "_ws")

    def __init__(self, ws: Worksheet | None = None) -> None:
        self._entries: list[ConditionalFormatting] = []
        self._ws = ws

    def __iter__(self) -> Iterator[ConditionalFormatting]:
        return iter(self._entries)

    def __len__(self) -> int:
        return len(self._entries)

    def __bool__(self) -> bool:
        return bool(self._entries)

    def _append_entry(self, entry: ConditionalFormatting) -> None:
        """Internal: used by the lazy reader to populate the container."""
        self._entries.append(entry)

    def add(self, range_string: str, rule: Rule) -> None:
        """Attach a new conditional-formatting rule.

        Landing in PR5 — this raises with a T1.5 pointer in modify mode
        and queues a pending CF in write mode.
        """
        ws = self._ws
        if ws is None:
            raise RuntimeError("ConditionalFormattingList is not attached to a worksheet")
        wb = ws._workbook  # noqa: SLF001
        if wb._rust_writer is None:  # noqa: SLF001
            raise NotImplementedError(
                "Adding conditional formatting rules to existing files is a T1.5 follow-up. "
                "Write mode (Workbook() + save) is supported."
            )
        # Find or create the CF entry for this range in our container.
        for entry in self._entries:
            if entry.sqref == range_string:
                entry.rules.append(rule)
                break
        else:
            self._entries.append(ConditionalFormatting(sqref=range_string, rules=[rule]))
        ws._pending_conditional_formats.append((range_string, rule))  # noqa: SLF001


__all__ = [
    "ConditionalFormatting",
    "ConditionalFormattingList",
]
