from __future__ import annotations

from pathlib import Path

import pytest

import wolfxl


def test_failed_native_save_preserves_existing_target(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    target = tmp_path / "target.xlsx"

    original = wolfxl.Workbook()
    original.active["A1"] = "original"
    original.save(target)
    before = target.read_bytes()

    replacement = wolfxl.Workbook()
    replacement.active["A1"] = "replacement"
    monkeypatch.setenv("WOLFXL_MAX_ZIP_ENTRY_BYTES", "1")

    with pytest.raises(Exception, match="too large"):
        replacement.save(target)

    assert target.read_bytes() == before
