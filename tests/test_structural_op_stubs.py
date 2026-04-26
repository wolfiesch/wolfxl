"""Structural-op stub tests.

Pin the contract for the 7 structural ops that openpyxl exposes but wolfxl
has not yet implemented. Until the corresponding RFCs ship (RFC-030 / 031 /
034 / 035 / 036), each call must raise ``NotImplementedError`` with a message
that points at the right RFC. The point is to give users a discoverable
roadmap entry instead of an ``AttributeError``.
"""

from __future__ import annotations

import pytest

import wolfxl

WORKSHEET_STUBS = [
    # RFC-030 ships in this branch — insert_rows / delete_rows now work
    # in modify mode. Remaining stubs are tracked by their RFCs.
    ("insert_cols", (2,), "RFC-031"),
    ("delete_cols", (2,), "RFC-031"),
    ("move_range", ("A1:B2",), "RFC-034"),
]


WORKBOOK_STUBS = [
    ("copy_worksheet", "RFC-035"),
    ("move_sheet", "RFC-036"),
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


def test_workbook_move_sheet_stub() -> None:
    wb, ws = _fresh_active()
    with pytest.raises(NotImplementedError, match="RFC-036"):
        wb.move_sheet(ws, 1)


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
