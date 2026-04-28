"""W4E.P1 regression: NativeWorkbook is consumed-on-save.

The fix: ``self.saved = true`` is set *before* the emit/write so a
panic in ``emit_xlsx`` or ``fs::write`` cannot leave the workbook in
a state where a retry would re-emit on partially-mutated data.

These tests assert the consumed-on-save contract holds for the native
pyclass — the sole write-mode backend as of W5.
"""
from __future__ import annotations

from pathlib import Path

import pytest


def test_second_save_raises_already_saved_native(tmp_path: Path) -> None:
    """A successful save() consumes the workbook; a second save() must
    raise (no silent overwrite, no retry path)."""
    import wolfxl

    wb = wolfxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.cell(row=1, column=1, value="hello")
    out1 = tmp_path / "first.xlsx"
    wb.save(str(out1))
    assert out1.exists() and out1.stat().st_size > 0

    out2 = tmp_path / "second.xlsx"
    with pytest.raises(Exception) as exc_info:
        wb.save(str(out2))
    assert "already saved" in str(exc_info.value).lower(), (
        f"expected 'already saved' in error, got: {exc_info.value!r}"
    )
    assert not out2.exists(), "second save() must not write any output"


def test_failed_save_still_consumes_workbook_native(tmp_path: Path) -> None:
    """W4E.P1 invariant: even if the first save() fails (e.g. unwritable
    path), the workbook is still marked saved. Retrying after a partial
    failure could re-emit on partially-mutated state — the lifecycle is
    designed to fail loudly instead.
    """
    import wolfxl

    wb = wolfxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.cell(row=1, column=1, value="hello")

    # /dev/full on Linux always-fails-write; on macOS we use a directory
    # path that doesn't exist. Either yields a write error from fs::write.
    bad_path = str(tmp_path / "no_such_dir" / "bad.xlsx")
    with pytest.raises(Exception):
        wb.save(bad_path)

    # Second save() — even to a valid path — must still raise
    # 'already saved'. Crucial: retry-on-failure is forbidden.
    good_path = tmp_path / "good.xlsx"
    with pytest.raises(Exception) as exc_info:
        wb.save(str(good_path))
    assert "already saved" in str(exc_info.value).lower(), (
        f"after failed first save, second save did not raise "
        f"'already saved': got {exc_info.value!r}"
    )
    assert not good_path.exists(), "second save() must not write any output"
