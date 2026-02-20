"""WorkbookEvaluator: recursive expression evaluator for Excel formulas.

Replaces fragile regex-based dispatch with a proper recursive descent
parser that handles balanced parentheses, operator precedence, and
arbitrarily nested expressions like ``=ROUND(SUM(A1:A5)*IF(B1>0,1.1,1.0),2)``.
"""

from __future__ import annotations

import logging
import re
from typing import TYPE_CHECKING, Any

from wolfxl.calc._functions import FunctionRegistry
from wolfxl.calc._graph import DependencyGraph
from wolfxl.calc._parser import expand_range
from wolfxl.calc._protocol import CellDelta, RecalcResult

if TYPE_CHECKING:
    from wolfxl._workbook import Workbook

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Expression parsing helpers
# ---------------------------------------------------------------------------


def _find_matching_paren(expr: str, start: int) -> int:
    """Index of the ``')'`` matching the ``'('`` at *expr[start]*, or -1."""
    depth = 1
    i = start + 1
    in_string = False
    while i < len(expr):
        ch = expr[i]
        if ch == '"':
            in_string = not in_string
        elif not in_string:
            if ch == '(':
                depth += 1
            elif ch == ')':
                depth -= 1
                if depth == 0:
                    return i
        i += 1
    return -1


def _match_function_call(expr: str) -> tuple[str, str] | None:
    """If *expr* is exactly ``FUNC(balanced_args)``, return ``(name, args_str)``.

    Uses balanced parenthesis matching so ``SUM(A1:A5)*2`` is NOT matched
    (there's trailing content after the close-paren).
    """
    stripped = expr.strip()
    m = re.match(r'^([A-Z][A-Z0-9_.]*)\s*\(', stripped, re.IGNORECASE)
    if not m:
        return None
    open_idx = m.end() - 1  # position of '('
    close_idx = _find_matching_paren(stripped, open_idx)
    # The close-paren must be the very last character
    if close_idx >= 0 and close_idx == len(stripped) - 1:
        return (m.group(1), stripped[open_idx + 1 : close_idx])
    return None


def _find_top_level_split(expr: str) -> tuple[str, str, str] | None:
    """Find the rightmost lowest-precedence binary operator at paren depth 0.

    Precedence (lowest to highest)::

        1. comparison   (>=, <=, <>, >, <, =)
        2. additive     (+, -)
        3. multiplicative (*, /)

    Right-to-left scan produces correct left-to-right associativity.
    Returns ``(left, op, right)`` or ``None``.
    """
    length = len(expr)

    for pass_type in ("cmp", "add", "mul"):
        depth = 0
        in_string = False
        i = length - 1
        while i > 0:
            ch = expr[i]

            # Skip string literals
            if ch == '"':
                in_string = not in_string
                i -= 1
                continue
            if in_string:
                i -= 1
                continue

            # Track parentheses (inverted for right-to-left)
            if ch == ')':
                depth += 1
                i -= 1
                continue
            if ch == '(':
                depth -= 1
                i -= 1
                continue

            if depth != 0:
                i -= 1
                continue

            matched_op: str | None = None
            op_start = i

            if pass_type == "cmp":
                # 2-char comparison operators checked first
                if i >= 1 and expr[i - 1 : i + 1] in (">=", "<=", "<>"):
                    matched_op = expr[i - 1 : i + 1]
                    op_start = i - 1
                elif ch in ('>', '<'):
                    matched_op = ch
                elif ch == '=' and not (i >= 1 and expr[i - 1] in ('>', '<', '!')):
                    matched_op = ch
            elif pass_type == "add" and ch in ('+', '-'):
                matched_op = ch
            elif pass_type == "mul" and ch in ('*', '/'):
                matched_op = ch

            if matched_op is not None:
                # Verify it's a binary operator (not unary prefix)
                if op_start <= 0:
                    i -= 1
                    continue
                # Check preceding non-space character
                j = op_start - 1
                while j >= 0 and expr[j] == ' ':
                    j -= 1
                if j < 0 or expr[j] in ('(', ',', '+', '-', '*', '/', '>', '<', '='):
                    i -= 1
                    continue

                left = expr[:op_start].strip()
                right = expr[op_start + len(matched_op) :].strip()
                if left and right:
                    return (left, matched_op, right)

            i -= 1

    return None


def _has_top_level_colon(expr: str) -> bool:
    """``True`` when *expr* contains ``:`` at paren depth 0 (range ref)."""
    depth = 0
    for ch in expr:
        if ch == '(':
            depth += 1
        elif ch == ')':
            depth -= 1
        elif ch == ':' and depth == 0:
            return True
    return False


def _binary_op(left: Any, op: str, right: Any) -> Any:
    """Evaluate an arithmetic binary operation."""
    if not isinstance(left, (int, float)) or not isinstance(right, (int, float)):
        return None
    if op == '+':
        return left + right
    if op == '-':
        return left - right
    if op == '*':
        return left * right
    if op == '/':
        return "#DIV/0!" if right == 0 else left / right
    return None


def _compare(left: Any, right: Any, op: str) -> bool:
    """Evaluate a comparison operation."""
    try:
        lf = float(left) if not isinstance(left, (int, float)) else left
        rf = float(right) if not isinstance(right, (int, float)) else right
    except (ValueError, TypeError):
        return False
    if op == '>':
        return lf > rf
    if op == '<':
        return lf < rf
    if op == '>=':
        return lf >= rf
    if op == '<=':
        return lf <= rf
    if op in ('=', '=='):
        return lf == rf
    if op in ('<>', '!='):
        return lf != rf
    return False


def _values_differ(a: Any, b: Any, tolerance: float) -> bool:
    """Check if two values differ beyond tolerance."""
    if a is None and b is None:
        return False
    if a is None or b is None:
        return True
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(float(a) - float(b)) > tolerance
    return a != b


# ---------------------------------------------------------------------------
# Evaluator
# ---------------------------------------------------------------------------


class WorkbookEvaluator:
    """Evaluates Excel formulas in a wolfxl Workbook.

    Usage::

        evaluator = WorkbookEvaluator()
        evaluator.load(workbook)
        results = evaluator.calculate()
        recalc = evaluator.recalculate({"Sheet1!A1": 42.0})
    """

    def __init__(self) -> None:
        self._cell_values: dict[str, Any] = {}
        self._graph = DependencyGraph()
        self._functions = FunctionRegistry()
        self._loaded = False

    def load(self, workbook: Workbook) -> None:
        """Scan workbook, store cell values, build dependency graph."""
        self._cell_values.clear()
        self._graph = DependencyGraph()

        for sheet_name in workbook.sheetnames:
            ws = workbook[sheet_name]
            for row in ws.iter_rows(values_only=False):
                for cell in row:
                    val = cell.value
                    cell_ref = f"{sheet_name}!{cell.coordinate}"
                    if isinstance(val, str) and val.startswith("="):
                        # Formula cell: store formula string, register in graph
                        self._cell_values[cell_ref] = val
                        self._graph.add_formula(cell_ref, val, sheet_name)
                    elif val is not None:
                        # Value cell: store the value
                        self._cell_values[cell_ref] = val

        self._loaded = True

    def calculate(self) -> dict[str, Any]:
        """Evaluate all formulas in topological order.

        Returns dict of cell_ref -> computed value for formula cells.
        """
        if not self._loaded:
            raise RuntimeError("Call load() before calculate()")

        order = self._graph.topological_order()
        results: dict[str, Any] = {}

        for cell_ref in order:
            formula = self._graph.formulas[cell_ref]
            value = self._evaluate_formula(cell_ref, formula)
            self._cell_values[cell_ref] = value
            results[cell_ref] = value

        return results

    def recalculate(
        self,
        perturbations: dict[str, float | int],
        tolerance: float = 1e-10,
    ) -> RecalcResult:
        """Perturb input cells and recompute affected formulas."""
        if not self._loaded:
            raise RuntimeError("Call load() before recalculate()")

        # Snapshot old values for delta computation
        old_values: dict[str, Any] = {}
        for cell_ref in self._graph.formulas:
            old_values[cell_ref] = self._cell_values.get(cell_ref)

        # Apply perturbations
        for cell_ref, value in perturbations.items():
            self._cell_values[cell_ref] = value

        # Find and evaluate affected cells
        affected = self._graph.affected_cells(set(perturbations.keys()))
        for cell_ref in affected:
            formula = self._graph.formulas[cell_ref]
            value = self._evaluate_formula(cell_ref, formula)
            self._cell_values[cell_ref] = value

        # Build deltas
        deltas: list[CellDelta] = []
        propagated = 0
        for cell_ref in affected:
            old_val = old_values.get(cell_ref)
            new_val = self._cell_values.get(cell_ref)
            if _values_differ(old_val, new_val, tolerance):
                propagated += 1
                deltas.append(CellDelta(
                    cell_ref=cell_ref,
                    old_value=old_val,
                    new_value=new_val,
                    formula=self._graph.formulas.get(cell_ref),
                ))

        max_depth = self._graph.max_depth(set(perturbations.keys()))

        return RecalcResult(
            perturbations=dict(perturbations),
            deltas=tuple(deltas),
            total_formula_cells=len(self._graph.formulas),
            propagated_cells=propagated,
            max_chain_depth=max_depth,
        )

    # ------------------------------------------------------------------
    # Formula evaluation (recursive descent)
    # ------------------------------------------------------------------

    def _evaluate_formula(self, cell_ref: str, formula: str) -> Any:
        """Evaluate a single formula string (starting with ``=``)."""
        body = formula.strip()
        if body.startswith('='):
            body = body[1:]
        sheet = self._sheet_from_ref(cell_ref)
        result = self._eval_expr(body.strip(), sheet)
        if result is not None:
            return result
        logger.debug("Cannot evaluate formula %r in %s", formula, cell_ref)
        return None

    def _eval_expr(self, expr: str, sheet: str) -> Any:
        """Recursively evaluate an expression (no leading ``=``).

        Dispatch order (first match wins):

        1. Binary/comparison split at top level (paren-aware, precedence-correct)
        2. Parenthesized sub-expression ``(...)``
        3. Function call ``FUNC(balanced_args)``
        4. Unary minus / plus
        5. Numeric literal
        6. String literal
        7. Boolean literal
        8. Cell reference
        """
        expr = expr.strip()
        if not expr:
            return None

        # 1. Binary split (comparison → additive → multiplicative)
        split = _find_top_level_split(expr)
        if split:
            left_str, op, right_str = split
            left_val = self._eval_expr(left_str, sheet)
            right_val = self._eval_expr(right_str, sheet)
            if op in ('+', '-', '*', '/'):
                return _binary_op(left_val, op, right_val)
            return _compare(left_val, right_val, op)

        # 2. Parenthesized sub-expression: (expr)
        if expr.startswith('('):
            close = _find_matching_paren(expr, 0)
            if close == len(expr) - 1:
                return self._eval_expr(expr[1:close], sheet)

        # 3. Function call: FUNC(balanced_args)
        func = _match_function_call(expr)
        if func:
            return self._eval_function(func[0].upper(), func[1], sheet)

        # 4. Unary minus / plus
        if expr.startswith('-'):
            val = self._eval_expr(expr[1:], sheet)
            if isinstance(val, (int, float)):
                return -val
            return val
        if expr.startswith('+'):
            return self._eval_expr(expr[1:], sheet)

        # 5. Numeric literal
        try:
            return float(expr) if '.' in expr else int(expr)
        except ValueError:
            pass

        # 6. String literal
        if len(expr) >= 2 and expr[0] == '"' and expr[-1] == '"':
            return expr[1:-1]

        # 7. Boolean
        upper = expr.upper()
        if upper == 'TRUE':
            return True
        if upper == 'FALSE':
            return False

        # 8. Cell reference
        return self._resolve_cell_ref(expr, sheet)

    # ------------------------------------------------------------------
    # Atom / argument resolution
    # ------------------------------------------------------------------

    def _resolve_cell_ref(self, expr: str, sheet: str) -> Any:
        """Resolve a cell reference string to its stored value."""
        clean = expr.strip().replace('$', '')
        if '!' in clean:
            parts = clean.split('!', 1)
            ref_sheet = parts[0].strip("'")
            ref = f"{ref_sheet}!{parts[1].upper()}"
        else:
            ref = f"{sheet}!{clean.upper()}"
        return self._cell_values.get(ref)

    def _resolve_range(self, arg: str, sheet: str) -> list[Any]:
        """Resolve a range like ``A1:A5`` to a list of cell values."""
        clean = arg.strip().replace('$', '')
        if '!' not in clean:
            range_ref = f"{sheet}!{clean.upper()}"
        else:
            parts = clean.split('!', 1)
            ref_sheet = parts[0].strip("'")
            range_ref = f"{ref_sheet}!{parts[1].upper()}"
        cells = expand_range(range_ref)
        return [self._cell_values.get(c) for c in cells]

    # ------------------------------------------------------------------
    # Function dispatch
    # ------------------------------------------------------------------

    def _eval_function(self, func_name: str, args_str: str, sheet: str) -> Any:
        """Evaluate a function call with resolved arguments."""
        func = self._functions.get(func_name)
        if func is None:
            logger.debug("Unsupported function: %s", func_name)
            return None
        args = self._parse_function_args(args_str, sheet)
        try:
            return func(args)
        except Exception as e:
            logger.debug("Error evaluating %s: %s", func_name, e)
            return None

    def _parse_function_args(self, args_str: str, sheet: str) -> list[Any]:
        """Split on commas at depth 0, resolve each argument."""
        args: list[Any] = []
        depth = 0
        current = ""

        for ch in args_str:
            if ch == '(':
                depth += 1
                current += ch
            elif ch == ')':
                depth -= 1
                current += ch
            elif ch == ',' and depth == 0:
                args.append(self._resolve_arg(current.strip(), sheet))
                current = ""
            else:
                current += ch

        if current.strip():
            args.append(self._resolve_arg(current.strip(), sheet))

        return args

    def _resolve_arg(self, arg: str, sheet: str) -> Any:
        """Resolve a single function argument.

        Range references (containing ``:`` at depth 0) return a list of
        cell values.  Everything else delegates to ``_eval_expr``.
        """
        if not arg:
            return None

        # Range reference at top level
        if _has_top_level_colon(arg) and not arg.startswith('"'):
            return self._resolve_range(arg, sheet)

        return self._eval_expr(arg, sheet)

    @staticmethod
    def _sheet_from_ref(cell_ref: str) -> str:
        """Extract sheet name from a canonical cell reference."""
        if '!' in cell_ref:
            return cell_ref.rsplit('!', 1)[0]
        return 'Sheet1'
