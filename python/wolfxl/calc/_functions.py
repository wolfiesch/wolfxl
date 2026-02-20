"""Function whitelist and builtin implementations for formula evaluation."""

from __future__ import annotations

import fnmatch
import math
import re
from dataclasses import dataclass
from typing import Any, Callable


# ---------------------------------------------------------------------------
# RangeValue: shape-aware 2D range container
# ---------------------------------------------------------------------------


@dataclass
class RangeValue:
    """A resolved cell range that preserves 2D shape metadata.

    Iterable and sized for backward compat with functions that expect lists.
    """

    values: list[Any]
    n_rows: int
    n_cols: int

    def get(self, row: int, col: int) -> Any:
        """Get value at 1-based (row, col) position."""
        if row < 1 or row > self.n_rows or col < 1 or col > self.n_cols:
            return None
        idx = (row - 1) * self.n_cols + (col - 1)
        return self.values[idx] if idx < len(self.values) else None

    def column(self, col: int) -> list[Any]:
        """Extract a 1-based column as a list."""
        if col < 1 or col > self.n_cols:
            return []
        return [self.values[(r * self.n_cols) + (col - 1)]
                for r in range(self.n_rows)
                if (r * self.n_cols) + (col - 1) < len(self.values)]

    def row(self, row: int) -> list[Any]:
        """Extract a 1-based row as a list."""
        if row < 1 or row > self.n_rows:
            return []
        start = (row - 1) * self.n_cols
        return self.values[start:start + self.n_cols]

    def as_flat(self) -> list[Any]:
        """Return values as a flat list."""
        return list(self.values)

    def __iter__(self):
        return iter(self.values)

    def __len__(self):
        return len(self.values)

# ---------------------------------------------------------------------------
# Whitelist: functions the calc engine will attempt to evaluate.
# Organized by category for readability.
# ---------------------------------------------------------------------------

FUNCTION_WHITELIST_V1: dict[str, str] = {
    # Math (10)
    "SUM": "math",
    "ABS": "math",
    "ROUND": "math",
    "ROUNDUP": "math",
    "ROUNDDOWN": "math",
    "INT": "math",
    "MOD": "math",
    "POWER": "math",
    "SQRT": "math",
    "SIGN": "math",
    # Logic (5)
    "IF": "logic",
    "AND": "logic",
    "OR": "logic",
    "NOT": "logic",
    "IFERROR": "logic",
    # Lookup (7)
    "VLOOKUP": "lookup",
    "HLOOKUP": "lookup",
    "INDEX": "lookup",
    "MATCH": "lookup",
    "OFFSET": "lookup",
    "CHOOSE": "lookup",
    "XLOOKUP": "lookup",
    # Statistical (9)
    "AVERAGE": "statistical",
    "COUNT": "statistical",
    "COUNTA": "statistical",
    "COUNTIF": "statistical",
    "COUNTIFS": "statistical",
    "MIN": "statistical",
    "MAX": "statistical",
    "SUMIF": "statistical",
    "SUMIFS": "statistical",
    # Financial (7)
    "PV": "financial",
    "FV": "financial",
    "PMT": "financial",
    "NPV": "financial",
    "IRR": "financial",
    "SLN": "financial",
    "DB": "financial",
    # Text (5)
    "LEFT": "text",
    "RIGHT": "text",
    "MID": "text",
    "LEN": "text",
    "CONCATENATE": "text",
}


def is_supported(func_name: str) -> bool:
    """Check if a function name is in the evaluation whitelist."""
    return func_name.upper() in FUNCTION_WHITELIST_V1


# ---------------------------------------------------------------------------
# Builtin implementations - pure Python, no external deps.
# Each takes a list of resolved argument values.
# ---------------------------------------------------------------------------


def _coerce_numeric(values: list[Any]) -> list[float]:
    """Flatten and coerce values to floats, skipping None/str/bool."""
    result: list[float] = []
    for v in values:
        if isinstance(v, RangeValue):
            result.extend(_coerce_numeric(v.values))
        elif isinstance(v, (list, tuple)):
            result.extend(_coerce_numeric(list(v)))
        elif isinstance(v, bool):
            # In Excel, TRUE=1, FALSE=0 in numeric context
            result.append(float(v))
        elif isinstance(v, (int, float)):
            result.append(float(v))
        # Skip None, str, errors
    return result


def _builtin_sum(args: list[Any]) -> float:
    return sum(_coerce_numeric(args))


def _builtin_abs(args: list[Any]) -> float:
    if len(args) != 1:
        raise ValueError("ABS requires exactly 1 argument")
    nums = _coerce_numeric(args)
    if not nums:
        raise ValueError("ABS: non-numeric argument")
    return abs(nums[0])


def _builtin_round(args: list[Any]) -> float:
    if len(args) < 1 or len(args) > 2:
        raise ValueError("ROUND requires 1 or 2 arguments")
    nums = _coerce_numeric([args[0]])
    if not nums:
        raise ValueError("ROUND: non-numeric argument")
    digits = int(_coerce_numeric([args[1]])[0]) if len(args) > 1 else 0
    return round(nums[0], digits)


def _builtin_roundup(args: list[Any]) -> float:
    if len(args) < 1 or len(args) > 2:
        raise ValueError("ROUNDUP requires 1 or 2 arguments")
    nums = _coerce_numeric([args[0]])
    if not nums:
        raise ValueError("ROUNDUP: non-numeric argument")
    digits = int(_coerce_numeric([args[1]])[0]) if len(args) > 1 else 0
    if digits == 0:
        return float(math.ceil(nums[0]))
    factor = 10 ** digits
    return math.ceil(nums[0] * factor) / factor


def _builtin_int(args: list[Any]) -> float:
    if len(args) != 1:
        raise ValueError("INT requires exactly 1 argument")
    nums = _coerce_numeric(args)
    if not nums:
        raise ValueError("INT: non-numeric argument")
    return float(math.floor(nums[0]))


def _builtin_if(args: list[Any]) -> Any:
    if len(args) < 2 or len(args) > 3:
        raise ValueError("IF requires 2 or 3 arguments")
    condition = args[0]
    # Excel truthy: 0/False/None/"" are falsy
    truthy = bool(condition) if not isinstance(condition, (int, float)) else condition != 0
    if truthy:
        return args[1]
    return args[2] if len(args) > 2 else False


def _builtin_iferror(args: list[Any]) -> Any:
    if len(args) != 2:
        raise ValueError("IFERROR requires exactly 2 arguments")
    value = args[0]
    # If the value is an error string (e.g., "#DIV/0!"), return the fallback
    if isinstance(value, str) and value.startswith("#"):
        return args[1]
    return value


def _builtin_and(args: list[Any]) -> bool:
    if not args:
        raise ValueError("AND requires at least 1 argument")
    for a in args:
        if isinstance(a, (RangeValue, list, tuple)):
            if not all(bool(x) for x in a if x is not None):
                return False
        elif not a:
            return False
    return True


def _builtin_or(args: list[Any]) -> bool:
    if not args:
        raise ValueError("OR requires at least 1 argument")
    for a in args:
        if isinstance(a, (RangeValue, list, tuple)):
            if any(bool(x) for x in a if x is not None):
                return True
        elif a:
            return True
    return False


def _builtin_not(args: list[Any]) -> bool:
    if len(args) != 1:
        raise ValueError("NOT requires exactly 1 argument")
    return not bool(args[0])


def _builtin_count(args: list[Any]) -> float:
    """COUNT - counts numeric values only."""
    return float(len(_coerce_numeric(args)))


def _builtin_counta(args: list[Any]) -> float:
    """COUNTA - counts non-empty values."""
    count = 0
    for v in args:
        if isinstance(v, (RangeValue, list, tuple)):
            count += sum(1 for x in v if x is not None)
        elif v is not None:
            count += 1
    return float(count)


def _builtin_min(args: list[Any]) -> float:
    nums = _coerce_numeric(args)
    if not nums:
        return 0.0
    return min(nums)


def _builtin_max(args: list[Any]) -> float:
    nums = _coerce_numeric(args)
    if not nums:
        return 0.0
    return max(nums)


def _builtin_average(args: list[Any]) -> float:
    nums = _coerce_numeric(args)
    if not nums:
        raise ValueError("AVERAGE: no numeric values")
    return sum(nums) / len(nums)


# ---------------------------------------------------------------------------
# Additional math builtins
# ---------------------------------------------------------------------------


def _builtin_rounddown(args: list[Any]) -> float:
    if len(args) < 1 or len(args) > 2:
        raise ValueError("ROUNDDOWN requires 1 or 2 arguments")
    nums = _coerce_numeric([args[0]])
    if not nums:
        raise ValueError("ROUNDDOWN: non-numeric argument")
    digits = int(_coerce_numeric([args[1]])[0]) if len(args) > 1 else 0
    if digits == 0:
        return float(math.trunc(nums[0]))
    factor = 10 ** digits
    return math.trunc(nums[0] * factor) / factor


def _builtin_mod(args: list[Any]) -> float:
    if len(args) != 2:
        raise ValueError("MOD requires exactly 2 arguments")
    nums = _coerce_numeric(args)
    if len(nums) != 2:
        raise ValueError("MOD: non-numeric argument")
    if nums[1] == 0:
        raise ValueError("MOD: division by zero")
    # Excel MOD: result has the sign of the divisor
    return nums[0] - nums[1] * math.floor(nums[0] / nums[1])


def _builtin_power(args: list[Any]) -> float | str:
    if len(args) != 2:
        raise ValueError("POWER requires exactly 2 arguments")
    nums = _coerce_numeric(args)
    if len(nums) != 2:
        raise ValueError("POWER: non-numeric argument")
    # Excel returns #NUM! for negative base with fractional exponent
    if nums[0] < 0 and not float(nums[1]).is_integer():
        return "#NUM!"
    return nums[0] ** nums[1]


def _builtin_sqrt(args: list[Any]) -> float:
    if len(args) != 1:
        raise ValueError("SQRT requires exactly 1 argument")
    nums = _coerce_numeric(args)
    if not nums:
        raise ValueError("SQRT: non-numeric argument")
    if nums[0] < 0:
        raise ValueError("SQRT: negative argument")
    return math.sqrt(nums[0])


def _builtin_sign(args: list[Any]) -> float:
    if len(args) != 1:
        raise ValueError("SIGN requires exactly 1 argument")
    nums = _coerce_numeric(args)
    if not nums:
        raise ValueError("SIGN: non-numeric argument")
    if nums[0] > 0:
        return 1.0
    if nums[0] < 0:
        return -1.0
    return 0.0


# ---------------------------------------------------------------------------
# Text builtins
# ---------------------------------------------------------------------------


def _coerce_string(val: Any) -> str:
    if val is None:
        return ""
    return str(val)


def _builtin_left(args: list[Any]) -> str:
    if len(args) < 1 or len(args) > 2:
        raise ValueError("LEFT requires 1 or 2 arguments")
    text = _coerce_string(args[0])
    num_chars = int(_coerce_numeric([args[1]])[0]) if len(args) > 1 else 1
    if num_chars < 0:
        return "#VALUE!"
    return text[:num_chars]


def _builtin_right(args: list[Any]) -> str:
    if len(args) < 1 or len(args) > 2:
        raise ValueError("RIGHT requires 1 or 2 arguments")
    text = _coerce_string(args[0])
    num_chars = int(_coerce_numeric([args[1]])[0]) if len(args) > 1 else 1
    return text[-num_chars:] if num_chars > 0 else ""


def _builtin_mid(args: list[Any]) -> str:
    if len(args) != 3:
        raise ValueError("MID requires exactly 3 arguments")
    text = _coerce_string(args[0])
    start = int(_coerce_numeric([args[1]])[0])
    num_chars = int(_coerce_numeric([args[2]])[0])
    if start < 1 or num_chars < 0:
        return "#VALUE!"
    # Excel MID is 1-indexed
    return text[start - 1 : start - 1 + num_chars]


def _builtin_len(args: list[Any]) -> float:
    if len(args) != 1:
        raise ValueError("LEN requires exactly 1 argument")
    return float(len(_coerce_string(args[0])))


def _builtin_concatenate(args: list[Any]) -> str:
    if not args:
        raise ValueError("CONCATENATE requires at least 1 argument")
    return "".join(_coerce_string(a) for a in args)


# ---------------------------------------------------------------------------
# Criteria matching engine (shared by SUMIF, SUMIFS, COUNTIF, COUNTIFS)
# ---------------------------------------------------------------------------

_CRITERIA_OP_RE = re.compile(r"^(>=|<=|<>|>|<|=)(.*)$")


def _parse_criteria(criteria: Any) -> Callable[[Any], bool]:
    """Parse an Excel criteria value into a predicate function.

    Supports:
    - Numeric exact match: ``100`` matches cells equal to 100
    - String exact match (case-insensitive): ``"Sales"``
    - Operator prefix: ``">100"``, ``"<=50"``, ``"<>0"``
    - Wildcards: ``"apple*"``, ``"?pple"`` (via fnmatch)
    """
    if isinstance(criteria, (int, float)):
        target = float(criteria)
        return lambda v: isinstance(v, (int, float)) and float(v) == target

    crit_str = str(criteria)

    # Check for operator prefix
    m = _CRITERIA_OP_RE.match(crit_str)
    if m:
        op, val_str = m.group(1), m.group(2).strip()
        try:
            threshold = float(val_str)
        except (ValueError, TypeError):
            # String comparison with operator
            val_lower = val_str.lower()
            if op == ">":
                return lambda v: str(v).lower() > val_lower if v is not None else False
            if op == "<":
                return lambda v: str(v).lower() < val_lower if v is not None else False
            if op == ">=":
                return lambda v: str(v).lower() >= val_lower if v is not None else False
            if op == "<=":
                return lambda v: str(v).lower() <= val_lower if v is not None else False
            if op == "<>":
                return lambda v: str(v).lower() != val_lower if v is not None else True
            if op == "=":
                return lambda v: str(v).lower() == val_lower if v is not None else False
            return lambda v: False

        if op == ">":
            return lambda v, t=threshold: isinstance(v, (int, float)) and float(v) > t
        if op == "<":
            return lambda v, t=threshold: isinstance(v, (int, float)) and float(v) < t
        if op == ">=":
            return lambda v, t=threshold: isinstance(v, (int, float)) and float(v) >= t
        if op == "<=":
            return lambda v, t=threshold: isinstance(v, (int, float)) and float(v) <= t
        if op == "<>":
            return lambda v, t=threshold: not (isinstance(v, (int, float)) and float(v) == t)
        if op == "=":
            return lambda v, t=threshold: isinstance(v, (int, float)) and float(v) == t

    # Wildcard check (contains * or ? not escaped)
    if "*" in crit_str or "?" in crit_str:
        pattern = crit_str.lower()
        return lambda v, p=pattern: fnmatch.fnmatch(str(v).lower(), p) if v is not None else False

    # Plain string exact match (case-insensitive)
    lower = crit_str.lower()
    return lambda v, l=lower: str(v).lower() == l if v is not None else False


def _match_criteria(criteria: Any, value: Any) -> bool:
    """Convenience: check whether *value* satisfies *criteria*."""
    return _parse_criteria(criteria)(value)


# ---------------------------------------------------------------------------
# Lookup builtins (INDEX, MATCH, XLOOKUP, CHOOSE)
# ---------------------------------------------------------------------------


def _builtin_index(args: list[Any]) -> Any:
    """INDEX(array, row_num [, col_num])."""
    if len(args) < 2 or len(args) > 3:
        raise ValueError("INDEX requires 2 or 3 arguments")
    array = args[0]
    row_num = args[1]
    col_num = args[2] if len(args) > 2 else None

    # Safety net: if row_num is None (e.g. from unsupported nested func), bail
    if row_num is None:
        return None

    row_num = int(float(row_num))

    if isinstance(array, RangeValue):
        if col_num is not None:
            col_num = int(float(col_num))
            if row_num < 1 or row_num > array.n_rows or col_num < 1 or col_num > array.n_cols:
                return "#REF!"
            return array.get(row_num, col_num)
        # 1D horizontal range: row_num acts as column index
        if array.n_rows == 1:
            if row_num < 1 or row_num > array.n_cols:
                return "#REF!"
            return array.get(1, row_num)
        # 1D column or multi-col: row_num selects row, return first col
        if row_num < 1 or row_num > array.n_rows:
            return "#REF!"
        if array.n_cols == 1:
            return array.get(row_num, 1)
        # Multi-col without col_num: return first column value
        return array.get(row_num, 1)

    # Plain list fallback
    if isinstance(array, (list, tuple)):
        if row_num < 1 or row_num > len(array):
            return "#REF!"
        return array[row_num - 1]

    return None


def _builtin_match(args: list[Any]) -> Any:
    """MATCH(lookup_value, lookup_array, [match_type]).

    match_type: 0=exact, 1=largest<=, -1=smallest>=. Default 0.
    """
    if len(args) < 2 or len(args) > 3:
        raise ValueError("MATCH requires 2 or 3 arguments")
    lookup_value = args[0]
    lookup_array = args[1]
    match_type = int(float(args[2])) if len(args) > 2 and args[2] is not None else 0

    # Flatten to list
    if isinstance(lookup_array, RangeValue):
        values = lookup_array.values
    elif isinstance(lookup_array, (list, tuple)):
        values = list(lookup_array)
    else:
        return "#N/A"

    if match_type == 0:
        # Exact match - case-insensitive for strings
        for i, v in enumerate(values):
            if v is None:
                continue
            if isinstance(lookup_value, str) and isinstance(v, str):
                if lookup_value.lower() == v.lower():
                    return i + 1  # 1-based
            elif isinstance(lookup_value, (int, float)) and isinstance(v, (int, float)):
                if float(lookup_value) == float(v):
                    return i + 1
            elif lookup_value == v:
                return i + 1
        return "#N/A"

    if match_type == 1:
        # Largest value <= lookup (assumes sorted ascending)
        best_idx = None
        for i, v in enumerate(values):
            if isinstance(v, (int, float)) and isinstance(lookup_value, (int, float)):
                if float(v) <= float(lookup_value):
                    best_idx = i + 1
        return best_idx if best_idx is not None else "#N/A"

    if match_type == -1:
        # Smallest value >= lookup (assumes sorted descending)
        best_idx = None
        for i, v in enumerate(values):
            if isinstance(v, (int, float)) and isinstance(lookup_value, (int, float)):
                if float(v) >= float(lookup_value):
                    best_idx = i + 1
        return best_idx if best_idx is not None else "#N/A"

    return "#N/A"


def _builtin_xlookup(args: list[Any]) -> Any:
    """XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode]).

    Only exact match (match_mode=0, search_mode=1) is built in.
    Other modes return None to fall through to formulas lib.
    """
    if len(args) < 3 or len(args) > 6:
        raise ValueError("XLOOKUP requires 3 to 6 arguments")
    lookup_value = args[0]
    lookup_array = args[1]
    return_array = args[2]
    if_not_found = args[3] if len(args) > 3 else "#N/A"
    match_mode = int(float(args[4])) if len(args) > 4 and args[4] is not None else 0
    search_mode = int(float(args[5])) if len(args) > 5 and args[5] is not None else 1

    # Only handle exact match with forward search
    if match_mode != 0 or search_mode not in (1, -1):
        return None  # fall through to formulas lib

    # Flatten arrays
    if isinstance(lookup_array, RangeValue):
        lookup_vals = lookup_array.values
    elif isinstance(lookup_array, (list, tuple)):
        lookup_vals = list(lookup_array)
    else:
        return if_not_found

    if isinstance(return_array, RangeValue):
        return_vals = return_array.values
    elif isinstance(return_array, (list, tuple)):
        return_vals = list(return_array)
    else:
        return if_not_found

    search_range = range(len(lookup_vals)) if search_mode == 1 else range(len(lookup_vals) - 1, -1, -1)

    for i in search_range:
        v = lookup_vals[i]
        if v is None:
            continue
        matched = False
        if isinstance(lookup_value, str) and isinstance(v, str):
            matched = lookup_value.lower() == v.lower()
        elif isinstance(lookup_value, (int, float)) and isinstance(v, (int, float)):
            matched = float(lookup_value) == float(v)
        else:
            matched = lookup_value == v
        if matched:
            return return_vals[i] if i < len(return_vals) else if_not_found

    return if_not_found


def _builtin_vlookup(args: list[Any]) -> Any:
    """VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup]).

    range_lookup: FALSE (or 0) = exact match, TRUE (or 1, default) = approximate.
    Approximate match assumes the first column is sorted ascending and finds
    the largest value <= lookup_value.
    """
    if len(args) < 3 or len(args) > 4:
        raise ValueError("VLOOKUP requires 3 or 4 arguments")
    lookup_value = args[0]
    table_array = args[1]
    col_index_num = int(float(args[2]))
    range_lookup = True
    if len(args) > 3 and args[3] is not None:
        rl = args[3]
        if isinstance(rl, bool):
            range_lookup = rl
        elif isinstance(rl, (int, float)):
            range_lookup = bool(rl)
        elif isinstance(rl, str):
            range_lookup = rl.upper() != "FALSE"

    if col_index_num < 1:
        return "#VALUE!"

    if isinstance(table_array, RangeValue):
        if col_index_num > table_array.n_cols:
            return "#REF!"
        first_col = table_array.column(1)
        return_col = table_array.column(col_index_num)
    elif isinstance(table_array, (list, tuple)):
        # Flat list treated as single column
        if col_index_num > 1:
            return "#REF!"
        first_col = list(table_array)
        return_col = first_col
    else:
        return "#N/A"

    if range_lookup:
        # Approximate match: largest value <= lookup_value (sorted ascending)
        best_idx = None
        for i, v in enumerate(first_col):
            if v is None:
                continue
            if isinstance(lookup_value, (int, float)) and isinstance(v, (int, float)):
                if float(v) <= float(lookup_value):
                    best_idx = i
            elif isinstance(lookup_value, str) and isinstance(v, str):
                if v.lower() <= lookup_value.lower():
                    best_idx = i
        if best_idx is None:
            return "#N/A"
        return return_col[best_idx] if best_idx < len(return_col) else "#N/A"
    else:
        # Exact match (case-insensitive for strings)
        for i, v in enumerate(first_col):
            if v is None:
                continue
            if isinstance(lookup_value, str) and isinstance(v, str):
                if lookup_value.lower() == v.lower():
                    return return_col[i] if i < len(return_col) else "#N/A"
            elif isinstance(lookup_value, (int, float)) and isinstance(v, (int, float)):
                if float(lookup_value) == float(v):
                    return return_col[i] if i < len(return_col) else "#N/A"
            elif lookup_value == v:
                return return_col[i] if i < len(return_col) else "#N/A"
        return "#N/A"


def _builtin_hlookup(args: list[Any]) -> Any:
    """HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup]).

    Searches the first row of a table and returns a value from the specified row.
    range_lookup: FALSE (or 0) = exact match, TRUE (or 1, default) = approximate.
    """
    if len(args) < 3 or len(args) > 4:
        raise ValueError("HLOOKUP requires 3 or 4 arguments")
    lookup_value = args[0]
    table_array = args[1]
    row_index_num = int(float(args[2]))
    range_lookup = True
    if len(args) > 3 and args[3] is not None:
        rl = args[3]
        if isinstance(rl, bool):
            range_lookup = rl
        elif isinstance(rl, (int, float)):
            range_lookup = bool(rl)
        elif isinstance(rl, str):
            range_lookup = rl.upper() != "FALSE"

    if row_index_num < 1:
        return "#VALUE!"

    if isinstance(table_array, RangeValue):
        if row_index_num > table_array.n_rows:
            return "#REF!"
        first_row = table_array.row(1)
        return_row = table_array.row(row_index_num)
    elif isinstance(table_array, (list, tuple)):
        # Flat list treated as single row
        if row_index_num > 1:
            return "#REF!"
        first_row = list(table_array)
        return_row = first_row
    else:
        return "#N/A"

    if range_lookup:
        # Approximate match: largest value <= lookup_value (sorted ascending)
        best_idx = None
        for i, v in enumerate(first_row):
            if v is None:
                continue
            if isinstance(lookup_value, (int, float)) and isinstance(v, (int, float)):
                if float(v) <= float(lookup_value):
                    best_idx = i
            elif isinstance(lookup_value, str) and isinstance(v, str):
                if v.lower() <= lookup_value.lower():
                    best_idx = i
        if best_idx is None:
            return "#N/A"
        return return_row[best_idx] if best_idx < len(return_row) else "#N/A"
    else:
        # Exact match (case-insensitive for strings)
        for i, v in enumerate(first_row):
            if v is None:
                continue
            if isinstance(lookup_value, str) and isinstance(v, str):
                if lookup_value.lower() == v.lower():
                    return return_row[i] if i < len(return_row) else "#N/A"
            elif isinstance(lookup_value, (int, float)) and isinstance(v, (int, float)):
                if float(lookup_value) == float(v):
                    return return_row[i] if i < len(return_row) else "#N/A"
            elif lookup_value == v:
                return return_row[i] if i < len(return_row) else "#N/A"
        return "#N/A"


def _builtin_choose(args: list[Any]) -> Any:
    """CHOOSE(index_num, value1, value2, ...)."""
    if len(args) < 2:
        raise ValueError("CHOOSE requires at least 2 arguments")
    index_num = int(float(args[0]))
    if index_num < 1 or index_num > len(args) - 1:
        return "#VALUE!"
    return args[index_num]


# ---------------------------------------------------------------------------
# Conditional aggregation builtins (SUMIF, SUMIFS, COUNTIF, COUNTIFS)
# ---------------------------------------------------------------------------


def _builtin_sumif(args: list[Any]) -> float:
    """SUMIF(criteria_range, criteria, [sum_range])."""
    if len(args) < 2 or len(args) > 3:
        raise ValueError("SUMIF requires 2 or 3 arguments")
    criteria_range = args[0]
    criteria = args[1]
    sum_range = args[2] if len(args) > 2 else None

    # Flatten ranges
    if isinstance(criteria_range, RangeValue):
        crit_vals = criteria_range.values
    elif isinstance(criteria_range, (list, tuple)):
        crit_vals = list(criteria_range)
    else:
        crit_vals = [criteria_range]

    if sum_range is None:
        sum_vals = crit_vals
    elif isinstance(sum_range, RangeValue):
        sum_vals = sum_range.values
    elif isinstance(sum_range, (list, tuple)):
        sum_vals = list(sum_range)
    else:
        sum_vals = [sum_range]

    predicate = _parse_criteria(criteria)
    total = 0.0
    for i, cv in enumerate(crit_vals):
        if predicate(cv):
            sv = sum_vals[i] if i < len(sum_vals) else 0
            if isinstance(sv, (int, float)):
                total += float(sv)
    return total


def _builtin_sumifs(args: list[Any]) -> float:
    """SUMIFS(sum_range, criteria_range1, criteria1, ...).

    Note: sum_range is FIRST (unlike SUMIF where it's last).
    """
    if len(args) < 3 or len(args) % 2 == 0:
        raise ValueError("SUMIFS requires sum_range + pairs of (criteria_range, criteria)")
    sum_range = args[0]

    # Flatten sum_range
    if isinstance(sum_range, RangeValue):
        sum_vals = sum_range.values
    elif isinstance(sum_range, (list, tuple)):
        sum_vals = list(sum_range)
    else:
        sum_vals = [sum_range]

    # Build predicate pairs
    predicates: list[tuple[list[Any], Callable[[Any], bool]]] = []
    for j in range(1, len(args), 2):
        crit_range = args[j]
        criteria = args[j + 1]
        if isinstance(crit_range, RangeValue):
            cv = crit_range.values
        elif isinstance(crit_range, (list, tuple)):
            cv = list(crit_range)
        else:
            cv = [crit_range]
        predicates.append((cv, _parse_criteria(criteria)))

    total = 0.0
    for i in range(len(sum_vals)):
        if all(pred(cv[i]) if i < len(cv) else False for cv, pred in predicates):
            sv = sum_vals[i]
            if isinstance(sv, (int, float)):
                total += float(sv)
    return total


def _builtin_countif(args: list[Any]) -> float:
    """COUNTIF(range, criteria)."""
    if len(args) != 2:
        raise ValueError("COUNTIF requires exactly 2 arguments")
    count_range = args[0]
    criteria = args[1]

    if isinstance(count_range, RangeValue):
        values = count_range.values
    elif isinstance(count_range, (list, tuple)):
        values = list(count_range)
    else:
        values = [count_range]

    predicate = _parse_criteria(criteria)
    return float(sum(1 for v in values if predicate(v)))


def _builtin_countifs(args: list[Any]) -> float:
    """COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2, ...])."""
    if len(args) < 2 or len(args) % 2 != 0:
        raise ValueError("COUNTIFS requires pairs of (criteria_range, criteria)")

    # Build predicate pairs
    predicates: list[tuple[list[Any], Callable[[Any], bool]]] = []
    for j in range(0, len(args), 2):
        crit_range = args[j]
        criteria = args[j + 1]
        if isinstance(crit_range, RangeValue):
            cv = crit_range.values
        elif isinstance(crit_range, (list, tuple)):
            cv = list(crit_range)
        else:
            cv = [crit_range]
        predicates.append((cv, _parse_criteria(criteria)))

    # Length of first criteria range determines row count
    n = len(predicates[0][0]) if predicates else 0
    count = 0
    for i in range(n):
        if all(pred(cv[i]) if i < len(cv) else False for cv, pred in predicates):
            count += 1
    return float(count)


# ---------------------------------------------------------------------------
# Registry
# ---------------------------------------------------------------------------

_BUILTINS: dict[str, Callable[[list[Any]], Any]] = {
    "SUM": _builtin_sum,
    "ABS": _builtin_abs,
    "ROUND": _builtin_round,
    "ROUNDUP": _builtin_roundup,
    "ROUNDDOWN": _builtin_rounddown,
    "INT": _builtin_int,
    "MOD": _builtin_mod,
    "POWER": _builtin_power,
    "SQRT": _builtin_sqrt,
    "SIGN": _builtin_sign,
    "IF": _builtin_if,
    "IFERROR": _builtin_iferror,
    "AND": _builtin_and,
    "OR": _builtin_or,
    "NOT": _builtin_not,
    "COUNT": _builtin_count,
    "COUNTA": _builtin_counta,
    "MIN": _builtin_min,
    "MAX": _builtin_max,
    "AVERAGE": _builtin_average,
    "LEFT": _builtin_left,
    "RIGHT": _builtin_right,
    "MID": _builtin_mid,
    "LEN": _builtin_len,
    "CONCATENATE": _builtin_concatenate,
    "INDEX": _builtin_index,
    "MATCH": _builtin_match,
    "VLOOKUP": _builtin_vlookup,
    "HLOOKUP": _builtin_hlookup,
    "XLOOKUP": _builtin_xlookup,
    "CHOOSE": _builtin_choose,
    "SUMIF": _builtin_sumif,
    "SUMIFS": _builtin_sumifs,
    "COUNTIF": _builtin_countif,
    "COUNTIFS": _builtin_countifs,
}


class FunctionRegistry:
    """Registry of callable function implementations.

    Starts with builtins and can be extended with custom functions.
    """

    def __init__(self) -> None:
        self._functions: dict[str, Callable[[list[Any]], Any]] = dict(_BUILTINS)

    def register(self, name: str, func: Callable[[list[Any]], Any]) -> None:
        self._functions[name.upper()] = func

    def get(self, name: str) -> Callable[[list[Any]], Any] | None:
        return self._functions.get(name.upper())

    def has(self, name: str) -> bool:
        return name.upper() in self._functions

    @property
    def supported_functions(self) -> frozenset[str]:
        return frozenset(self._functions.keys())
