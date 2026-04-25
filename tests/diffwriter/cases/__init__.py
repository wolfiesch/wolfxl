"""Hand-authored cases for the differential writer harness.

Each module exports ``CASES: list[tuple[str, Callable[[Workbook], None]]]``.
The harness driver in ``tests/diffwriter/test_cases.py`` discovers every
case across this package, parameterizes a pytest test per case, and runs
the 3-layer diff between oracle and native outputs.

Cases are scoped tightly: each one exercises one OOXML construct. Keep
them small (~30 LOC) — comprehensive feature coverage comes from the count
of cases, not the depth of any single one.
"""
