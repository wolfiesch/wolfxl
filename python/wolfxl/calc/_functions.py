"""Function whitelist and builtin implementations for formula evaluation."""

from __future__ import annotations

import math
from typing import Any, Callable

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
    # Lookup (6)
    "VLOOKUP": "lookup",
    "HLOOKUP": "lookup",
    "INDEX": "lookup",
    "MATCH": "lookup",
    "OFFSET": "lookup",
    "CHOOSE": "lookup",
    # Statistical (6)
    "AVERAGE": "statistical",
    "COUNT": "statistical",
    "COUNTA": "statistical",
    "COUNTIF": "statistical",
    "MIN": "statistical",
    "MAX": "statistical",
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
        if isinstance(v, (list, tuple)):
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
        if isinstance(a, (list, tuple)):
            if not all(bool(x) for x in a if x is not None):
                return False
        elif not a:
            return False
    return True


def _builtin_or(args: list[Any]) -> bool:
    if not args:
        raise ValueError("OR requires at least 1 argument")
    for a in args:
        if isinstance(a, (list, tuple)):
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
        if isinstance(v, (list, tuple)):
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


def _builtin_power(args: list[Any]) -> float:
    if len(args) != 2:
        raise ValueError("POWER requires exactly 2 arguments")
    nums = _coerce_numeric(args)
    if len(nums) != 2:
        raise ValueError("POWER: non-numeric argument")
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
