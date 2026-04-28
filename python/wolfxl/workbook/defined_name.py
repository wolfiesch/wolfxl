"""openpyxl.workbook.defined_name compatibility.

T1 makes ``DefinedName`` a real dataclass.  Read access comes through
``wb.defined_names`` — that returns a :class:`DefinedNameDict` whose
values are ``DefinedName`` objects instead of bare strings.

Breaking change from T0: callers that did ``wb.defined_names["X"]`` and
expected a string must now do ``wb.defined_names["X"].value`` (or
``.attr_text`` for openpyxl parity). See CHANGELOG for the 1-line
migration.

RFC-021 — modify-mode defined-name mutation — accepts ``attr_text`` as
an alias for ``value`` to match openpyxl's canonical attribute name
(`openpyxl.workbook.defined_name.DefinedName.attr_text`). Supplying
both with conflicting values raises ``TypeError``; supplying neither
also raises ``TypeError``.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from wolfxl.workbook import DefinedNameDict


class DefinedName:
    """A workbook-scoped or sheet-scoped name range.

    ``value`` holds the ``refers_to`` expression (``Sheet1!$A$1:$A$10``
    or an external reference). ``localSheetId`` is ``None`` for
    workbook-scoped names or the 0-based sheet index for sheet-scoped
    ones. ``hidden=True`` marks internal names Excel uses for print
    areas and table ranges.

    ``attr_text`` is an openpyxl-compat alias for ``value`` accepted on
    construction; ``DefinedName(name="X", attr_text="Sheet1!A1")`` is
    equivalent to ``DefinedName(name="X", value="Sheet1!A1")``.
    """

    __slots__ = ("name", "value", "comment", "localSheetId", "hidden")

    def __init__(
        self,
        name: str,
        value: str | None = None,
        comment: str | None = None,
        localSheetId: int | None = None,  # noqa: N803 - openpyxl public API
        hidden: bool = False,
        *,
        attr_text: str | None = None,
    ) -> None:
        if value is None and attr_text is None:
            raise TypeError(
                "DefinedName requires 'value' (or its openpyxl alias 'attr_text')"
            )
        if value is not None and attr_text is not None and value != attr_text:
            raise TypeError(
                "DefinedName: pass either 'value' or 'attr_text', not both with conflicting values"
            )
        self.name: str = name
        self.value: str = value if value is not None else attr_text  # type: ignore[assignment]
        self.comment: str | None = comment
        self.localSheetId: int | None = localSheetId  # noqa: N815
        self.hidden: bool = hidden

    @property
    def attr_text(self) -> str:
        """openpyxl alias for ``.value``."""
        return self.value

    @attr_text.setter
    def attr_text(self, v: str) -> None:
        self.value = v

    def __repr__(self) -> str:
        bits = [f"name={self.name!r}", f"value={self.value!r}"]
        if self.localSheetId is not None:
            bits.append(f"localSheetId={self.localSheetId}")
        if self.hidden:
            bits.append("hidden=True")
        if self.comment is not None:
            bits.append(f"comment={self.comment!r}")
        return f"DefinedName({', '.join(bits)})"

    def __eq__(self, other: Any) -> bool:
        if not isinstance(other, DefinedName):
            return NotImplemented
        return (
            self.name == other.name
            and self.value == other.value
            and self.comment == other.comment
            and self.localSheetId == other.localSheetId
            and self.hidden == other.hidden
        )

    def __hash__(self) -> int:
        return hash((self.name, self.value, self.localSheetId, self.hidden))


class DefinedNameList(list):
    """openpyxl-shaped list-of-:class:`DefinedName`.

    openpyxl 3.0 used a ``DefinedNameList`` — newer releases moved to
    a dict.  Wolfxl exposes both shapes for import-compat parity:
    :class:`DefinedNameList` here, :class:`~wolfxl.workbook.DefinedNameDict`
    in :mod:`wolfxl.workbook` (the real container backing
    ``wb.defined_names``).

    Pod 2 (RFC-060 §2.6).
    """


def __getattr__(name: str):  # type: ignore[no-untyped-def]
    """Lazy re-export to dodge a circular ``wolfxl.workbook`` import."""
    if name == "DefinedNameDict":
        from wolfxl.workbook import DefinedNameDict as _D
        return _D
    raise AttributeError(name)


__all__ = ["DefinedName", "DefinedNameDict", "DefinedNameList"]
