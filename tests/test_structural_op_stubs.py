"""Structural-op stub tests.

Pin the contract for the structural ops that openpyxl exposes but wolfxl
has not yet implemented. Until the corresponding RFCs ship (RFC-030 / 031 /
034 / 035), each call must raise ``NotImplementedError`` with a message
that points at the right RFC. The point is to give users a discoverable
roadmap entry instead of an ``AttributeError``.

RFC-036 (``Workbook.move_sheet``) and RFC-034 (``Worksheet.move_range``)
shipped in WolfXL 1.1; their tests live in ``test_move_sheet_modify.py``
and ``test_move_range_modify.py``. Only ``copy_worksheet`` remains
stubbed at the workbook level.
"""

from __future__ import annotations

import pytest

import wolfxl

# All worksheet-level structural-op stubs have shipped (RFC-030, 031, 034).
# This list intentionally stays empty — adding entries here marks "not
# yet implemented" surfaces; removing them marks "shipped".
WORKSHEET_STUBS: list[tuple[str, tuple, str]] = []


WORKBOOK_STUBS = [
    ("copy_worksheet", "RFC-035"),
]


def _fresh_active() -> tuple[wolfxl.Workbook, wolfxl.Worksheet]:
    """Return ``(workbook, active_sheet)`` with a non-None active sheet."""
    wb = wolfxl.Workbook()
    ws = wb.active
    assert ws is not None
    return wb, ws


@pytest.mark.parametrize(("method", "args", "rfc"), WORKSHEET_STUBS)
def test_worksheet_stub_raises_with_rfc_pointer(
    method: str, args: tuple, rfc: str
) -> None:
    _wb, ws = _fresh_active()
    fn = getattr(ws, method)
    with pytest.raises(NotImplementedError, match=rfc):
        fn(*args)


def test_workbook_copy_worksheet_stub() -> None:
    wb, ws = _fresh_active()
    with pytest.raises(NotImplementedError, match="RFC-035"):
        wb.copy_worksheet(ws)


@pytest.mark.parametrize(("method", "_rfc"), WORKBOOK_STUBS)
def test_workbook_stubs_mention_workaround(method: str, _rfc: str) -> None:
    """Every stub message should include the openpyxl workaround pointer."""
    wb, ws = _fresh_active()
    fn = getattr(wb, method)
    args: tuple = (ws,) if method == "copy_worksheet" else (ws, 1)
    with pytest.raises(NotImplementedError, match="wolfxl.load_workbook"):
        fn(*args)


@pytest.mark.parametrize(("method", "args", "_rfc"), WORKSHEET_STUBS)
def test_worksheet_stubs_mention_workaround(
    method: str, args: tuple, _rfc: str
) -> None:
    _wb, ws = _fresh_active()
    fn = getattr(ws, method)
    with pytest.raises(NotImplementedError, match="wolfxl.load_workbook"):
        fn(*args)
