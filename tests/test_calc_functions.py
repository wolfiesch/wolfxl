"""Tests for wolfxl.calc function registry and builtins."""

from __future__ import annotations

import pytest
from wolfxl.calc._functions import (
    _BUILTINS,
    FUNCTION_WHITELIST_V1,
    FunctionRegistry,
    is_supported,
)


class TestWhitelist:
    def test_whitelist_has_39_functions(self) -> None:
        assert len(FUNCTION_WHITELIST_V1) == 39

    def test_all_categories_represented(self) -> None:
        categories = set(FUNCTION_WHITELIST_V1.values())
        assert categories == {"math", "logic", "lookup", "statistical", "financial", "text"}

    def test_is_supported_case_insensitive(self) -> None:
        assert is_supported("sum")
        assert is_supported("SUM")
        assert is_supported("Sum")
        assert not is_supported("WEBSERVICE")
        assert not is_supported("RAND")


class TestFunctionRegistry:
    def test_builtins_registered(self) -> None:
        reg = FunctionRegistry()
        assert reg.has("SUM")
        assert reg.has("IF")
        assert reg.has("AVERAGE")

    def test_custom_registration(self) -> None:
        reg = FunctionRegistry()
        reg.register("MYFUNC", lambda args: 42)
        assert reg.has("MYFUNC")
        assert reg.get("MYFUNC")([]) == 42

    def test_case_insensitive_lookup(self) -> None:
        reg = FunctionRegistry()
        assert reg.get("sum") is reg.get("SUM")

    def test_supported_functions_property(self) -> None:
        reg = FunctionRegistry()
        funcs = reg.supported_functions
        assert isinstance(funcs, frozenset)
        assert "SUM" in funcs


class TestBuiltinSUM:
    def test_basic(self) -> None:
        fn = _BUILTINS["SUM"]
        assert fn([1, 2, 3]) == 6.0

    def test_nested_lists(self) -> None:
        fn = _BUILTINS["SUM"]
        assert fn([[1, 2], [3, 4]]) == 10.0

    def test_skip_none_and_strings(self) -> None:
        fn = _BUILTINS["SUM"]
        assert fn([1, None, "text", 3]) == 4.0

    def test_empty(self) -> None:
        fn = _BUILTINS["SUM"]
        assert fn([]) == 0.0

    def test_booleans_coerced(self) -> None:
        fn = _BUILTINS["SUM"]
        assert fn([True, False, 1]) == 2.0


class TestBuiltinABS:
    def test_positive(self) -> None:
        assert _BUILTINS["ABS"]([-5]) == 5.0

    def test_zero(self) -> None:
        assert _BUILTINS["ABS"]([0]) == 0.0

    def test_already_positive(self) -> None:
        assert _BUILTINS["ABS"]([3.14]) == 3.14

    def test_wrong_arity(self) -> None:
        with pytest.raises(ValueError, match="exactly 1"):
            _BUILTINS["ABS"]([1, 2])


class TestBuiltinROUND:
    def test_round_default_digits(self) -> None:
        assert _BUILTINS["ROUND"]([3.14159]) == 3.0

    def test_round_2_digits(self) -> None:
        assert _BUILTINS["ROUND"]([3.14159, 2]) == 3.14

    def test_round_negative_digits(self) -> None:
        assert _BUILTINS["ROUND"]([1234, -2]) == 1200.0


class TestBuiltinROUNDUP:
    def test_roundup_basic(self) -> None:
        assert _BUILTINS["ROUNDUP"]([3.2]) == 4.0

    def test_roundup_2_digits(self) -> None:
        assert _BUILTINS["ROUNDUP"]([3.141, 2]) == 3.15


class TestBuiltinINT:
    def test_positive(self) -> None:
        assert _BUILTINS["INT"]([3.7]) == 3.0

    def test_negative(self) -> None:
        # Excel INT floors toward negative infinity
        assert _BUILTINS["INT"]([-3.2]) == -4.0


class TestBuiltinIF:
    def test_true_branch(self) -> None:
        assert _BUILTINS["IF"]([True, "yes", "no"]) == "yes"

    def test_false_branch(self) -> None:
        assert _BUILTINS["IF"]([False, "yes", "no"]) == "no"

    def test_numeric_condition(self) -> None:
        assert _BUILTINS["IF"]([1, "yes", "no"]) == "yes"
        assert _BUILTINS["IF"]([0, "yes", "no"]) == "no"

    def test_missing_false_branch(self) -> None:
        assert _BUILTINS["IF"]([False, "yes"]) is False


class TestBuiltinIFERROR:
    def test_no_error(self) -> None:
        assert _BUILTINS["IFERROR"]([42, 0]) == 42

    def test_error_string(self) -> None:
        assert _BUILTINS["IFERROR"](["#DIV/0!", 0]) == 0

    def test_ref_error(self) -> None:
        assert _BUILTINS["IFERROR"](["#REF!", "fallback"]) == "fallback"


class TestBuiltinLogic:
    def test_and_all_true(self) -> None:
        assert _BUILTINS["AND"]([True, True, 1]) is True

    def test_and_one_false(self) -> None:
        assert _BUILTINS["AND"]([True, False]) is False

    def test_or_one_true(self) -> None:
        assert _BUILTINS["OR"]([False, True]) is True

    def test_or_all_false(self) -> None:
        assert _BUILTINS["OR"]([False, 0, None]) is False

    def test_not(self) -> None:
        assert _BUILTINS["NOT"]([True]) is False
        assert _BUILTINS["NOT"]([False]) is True


class TestBuiltinCounting:
    def test_count_numeric(self) -> None:
        assert _BUILTINS["COUNT"]([1, "text", None, 3.5, True]) == 3.0

    def test_counta_non_empty(self) -> None:
        assert _BUILTINS["COUNTA"]([1, "text", None, 3.5]) == 3.0

    def test_count_empty(self) -> None:
        assert _BUILTINS["COUNT"]([]) == 0.0


class TestBuiltinMinMax:
    def test_min(self) -> None:
        assert _BUILTINS["MIN"]([3, 1, 4, 1, 5]) == 1.0

    def test_max(self) -> None:
        assert _BUILTINS["MAX"]([3, 1, 4, 1, 5]) == 5.0

    def test_min_empty(self) -> None:
        assert _BUILTINS["MIN"]([]) == 0.0

    def test_max_nested(self) -> None:
        assert _BUILTINS["MAX"]([[1, 2], [3, 4]]) == 4.0


class TestBuiltinAVERAGE:
    def test_basic(self) -> None:
        assert _BUILTINS["AVERAGE"]([2, 4, 6]) == 4.0

    def test_empty_raises(self) -> None:
        with pytest.raises(ValueError, match="no numeric"):
            _BUILTINS["AVERAGE"]([])

    def test_skip_non_numeric(self) -> None:
        assert _BUILTINS["AVERAGE"]([10, None, "text", 20]) == 15.0


class TestBuiltinDivisionByZero:
    """Edge case: ensure no unhandled ZeroDivisionError from builtins."""

    def test_average_single(self) -> None:
        assert _BUILTINS["AVERAGE"]([0]) == 0.0
