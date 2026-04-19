"""openpyxl-shape coordinate utilities.

These wrap WolfXL's existing primitives in ``wolfxl._utils`` while presenting
the openpyxl API contract — error messages, value bounds, and tuple shapes
match openpyxl 3.1.x.
"""

from __future__ import annotations

import re
from functools import cache
from string import ascii_uppercase

# Verbatim from openpyxl/utils/cell.py — keep in sync if openpyxl changes the
# pattern. The whole point of this module is bug-for-bug compatibility.
_RANGE_EXPR = r"""
[$]?(?P<min_col>[A-Za-z]{1,3})?
[$]?(?P<min_row>\d+)?
(:[$]?(?P<max_col>[A-Za-z]{1,3})?
[$]?(?P<max_row>\d+)?)?
"""
_ABSOLUTE_RE = re.compile("^" + _RANGE_EXPR + "$", re.VERBOSE)
_COORD_RE = re.compile(r"^[$]?([A-Za-z]{1,3})[$]?(\d+)$")

_DECIMAL_TO_ALPHA = [""] + list(ascii_uppercase)
_ALPHA_TO_DECIMAL = {letter: pos for pos, letter in enumerate(ascii_uppercase, 1)}
_POWERS = (1, 26, 676)


@cache
def get_column_letter(col_idx: int) -> str:
    """1-based column index → letter. Capped at 18278 (ZZZ) per openpyxl."""
    if not 1 <= col_idx <= 18278:
        raise ValueError(f"Invalid column index {col_idx}")

    if col_idx < 26:
        return _DECIMAL_TO_ALPHA[col_idx]

    result: list[str] = []
    while col_idx:
        col_idx, remainder = divmod(col_idx, 26)
        result.insert(0, _DECIMAL_TO_ALPHA[remainder])
        if not remainder:
            col_idx -= 1
            result.insert(0, "Z")
    return "".join(result)


@cache
def column_index_from_string(col: str) -> int:
    """Column letter → 1-based index. Accepts up to 3 letters (A..ZZZ)."""
    error_msg = f"'{col}' is not a valid column name. Column names are from A to ZZZ"
    if len(col) > 3:
        raise ValueError(error_msg)
    idx = 0
    for letter, power in zip(reversed(col.upper()), _POWERS):
        try:
            pos = _ALPHA_TO_DECIMAL[letter]
        except KeyError as exc:
            raise ValueError(error_msg) from exc
        idx += pos * power
    if not 0 < idx < 18279:
        raise ValueError(error_msg)
    return idx


def range_boundaries(range_string: str) -> tuple[int | None, int | None, int | None, int | None]:
    """Parse ``'A1:D10'`` / ``'$A$1:$D$10'`` → ``(min_col, min_row, max_col, max_row)``.

    Matches openpyxl's contract: degenerate single-cell refs return identical
    min/max; ``'A:A'`` returns ``(1, None, 1, None)``; ``'1:1'`` returns
    ``(None, 1, None, 1)``.
    """
    msg = f"{range_string} is not a valid coordinate or range"
    m = _ABSOLUTE_RE.match(range_string)
    if not m:
        raise ValueError(msg)

    min_col, min_row, sep, max_col, max_row = m.groups()

    if sep:
        cols = (min_col, max_col)
        rows = (min_row, max_row)
        # Mixed-validity check from openpyxl: either every coord is set, or
        # cols-only (whole-column ref), or rows-only (whole-row ref).
        if not (
            all(cols + rows)
            or (all(cols) and not any(rows))
            or (all(rows) and not any(cols))
        ):
            raise ValueError(msg)

    min_col_i = column_index_from_string(min_col) if min_col is not None else None
    min_row_i = int(min_row) if min_row is not None else None
    max_col_i = column_index_from_string(max_col) if max_col is not None else min_col_i
    max_row_i = int(max_row) if max_row is not None else min_row_i

    return min_col_i, min_row_i, max_col_i, max_row_i


def coordinate_to_tuple(coordinate: str) -> tuple[int, int]:
    """``'B3'`` → ``(3, 2)``. Returns ``(row, col)``, both 1-based."""
    m = _COORD_RE.match(coordinate)
    if not m:
        raise ValueError(f"Invalid cell coordinates ({coordinate})")
    col_letter, row_str = m.groups()
    return int(row_str), column_index_from_string(col_letter)
