"""``openpyxl.utils.formulas`` — Excel function name catalog.

openpyxl exposes a frozen ``FORMULAE`` set of every Excel-recognised
function name (``"SUM"``, ``"VLOOKUP"``, ``"XLOOKUP"``, ...) so callers
can validate user-supplied formula strings against the canonical list.

Wolfxl's calc engine has its own function registry under
:mod:`wolfxl.calc._functions`; this module exposes that catalogue under
the openpyxl-shaped name.  Names are uppercased to match openpyxl.

Pod 2 (RFC-060).
"""

from __future__ import annotations

try:
    from wolfxl.calc._functions import _BUILTINS as _FUNCTIONS
    FORMULAE: frozenset[str] = frozenset(name.upper() for name in _FUNCTIONS)
except Exception:  # pragma: no cover — defensive: calc engine optional at import.
    FORMULAE = frozenset()


__all__ = ["FORMULAE"]
