from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

import wolfxl


def _write_zip(path: Path, entries: dict[str, bytes]) -> None:
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, data in entries.items():
            zf.writestr(name, data)


def test_rejects_oversized_zip_entry(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    path = tmp_path / "oversized.xlsx"
    _write_zip(path, {"xl/workbook.xml": b"x" * 128})
    monkeypatch.setenv("WOLFXL_MAX_ZIP_ENTRY_BYTES", "64")

    with pytest.raises(Exception, match="too large"):
        wolfxl.load_workbook(path)


def test_rejects_zip_bomb_ratio(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    path = tmp_path / "ratio.xlsx"
    _write_zip(path, {"xl/workbook.xml": b"x" * 20_000})
    monkeypatch.setenv("WOLFXL_MAX_ZIP_ENTRY_BYTES", "1000000")
    monkeypatch.setenv("WOLFXL_MAX_ZIP_COMPRESSION_RATIO", "2")

    with pytest.raises(Exception, match="compression ratio"):
        wolfxl.load_workbook(path)


def test_rejects_unsafe_part_path(tmp_path: Path) -> None:
    path = tmp_path / "unsafe.xlsx"
    _write_zip(path, {"../xl/workbook.xml": b"<workbook/>"})

    with pytest.raises(Exception, match="unsafe OOXML package part path"):
        wolfxl.load_workbook(path)
