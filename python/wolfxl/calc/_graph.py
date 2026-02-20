"""Dependency graph for formula cells with topological ordering."""

from __future__ import annotations

from collections import deque
from typing import TYPE_CHECKING

from wolfxl.calc._parser import all_references

if TYPE_CHECKING:
    from wolfxl._workbook import Workbook


class DependencyGraph:
    """Tracks formula cell dependencies for evaluation ordering.

    All cell references use canonical "SheetName!A1" format.
    """

    __slots__ = ("dependencies", "dependents", "formulas")

    def __init__(self) -> None:
        # cell -> set of cells it reads from
        self.dependencies: dict[str, set[str]] = {}
        # cell -> set of cells that read from it (reverse edges)
        self.dependents: dict[str, set[str]] = {}
        # cell -> formula string
        self.formulas: dict[str, str] = {}

    def add_formula(self, cell_ref: str, formula: str, current_sheet: str) -> None:
        """Register a formula cell and its dependencies."""
        self.formulas[cell_ref] = formula
        refs = all_references(formula, current_sheet)

        self.dependencies[cell_ref] = set(refs)

        for ref in refs:
            if ref not in self.dependents:
                self.dependents[ref] = set()
            self.dependents[ref].add(cell_ref)

    def topological_order(self) -> list[str]:
        """Return formula cells in evaluation order (Kahn's algorithm).

        Raises ValueError if a circular reference is detected.
        """
        # Only consider formula cells
        formula_cells = set(self.formulas.keys())
        if not formula_cells:
            return []

        # Compute in-degrees within formula cells only
        in_degree: dict[str, int] = {}
        for cell in formula_cells:
            deps = self.dependencies.get(cell, set())
            # Only count deps that are themselves formula cells
            in_degree[cell] = len(deps & formula_cells)

        # Start with formula cells that have no formula-cell dependencies
        queue: deque[str] = deque()
        for cell in formula_cells:
            if in_degree[cell] == 0:
                queue.append(cell)

        order: list[str] = []
        while queue:
            cell = queue.popleft()
            order.append(cell)
            # Reduce in-degree for dependent formula cells
            for dep in self.dependents.get(cell, set()):
                if dep in formula_cells:
                    in_degree[dep] -= 1
                    if in_degree[dep] == 0:
                        queue.append(dep)

        if len(order) != len(formula_cells):
            missing = formula_cells - set(order)
            raise ValueError(f"Circular reference detected involving: {missing}")

        return order

    def affected_cells(self, changed_cells: set[str]) -> list[str]:
        """Find all formula cells affected by changes, in evaluation order.

        Uses BFS on the dependents graph, then filters to topological order.
        """
        affected: set[str] = set()
        queue: deque[str] = deque(changed_cells)
        visited: set[str] = set(changed_cells)

        while queue:
            cell = queue.popleft()
            for dep in self.dependents.get(cell, set()):
                if dep not in visited:
                    visited.add(dep)
                    queue.append(dep)
                    if dep in self.formulas:
                        affected.add(dep)

        # Return in topological order
        full_order = self.topological_order()
        return [c for c in full_order if c in affected]

    def max_depth(self, roots: set[str]) -> int:
        """Longest dependency chain from root cells through formula cells."""
        if not roots:
            return 0

        depth: dict[str, int] = {r: 0 for r in roots}
        queue: deque[str] = deque(roots)
        max_d = 0

        while queue:
            cell = queue.popleft()
            current_depth = depth[cell]
            for dep in self.dependents.get(cell, set()):
                if dep in self.formulas:
                    new_depth = current_depth + 1
                    if dep not in depth or new_depth > depth[dep]:
                        depth[dep] = new_depth
                        max_d = max(max_d, new_depth)
                        queue.append(dep)

        return max_d

    @classmethod
    def from_workbook(cls, workbook: Workbook) -> DependencyGraph:
        """Build a dependency graph by scanning all sheets for formula cells."""
        graph = cls()

        for sheet_name in workbook.sheetnames:
            ws = workbook[sheet_name]
            for row in ws.iter_rows(values_only=False):
                for cell in row:
                    val = cell.value
                    if isinstance(val, str) and val.startswith("="):
                        cell_ref = f"{sheet_name}!{cell.coordinate}"
                        graph.add_formula(cell_ref, val, sheet_name)

        return graph
