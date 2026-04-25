"""W4E.P4 regression: ``contentStatus`` round-trips through both backends.

Before W4E.P4 the native writer silently dropped ``contentStatus`` because
``DocProperties`` had no field for it. Oracle wired it via
``rust_xlsxwriter``'s ``set_status`` so the dual-backend diff showed a
divergence on every workbook that set a status. The fix adds
``DocProperties.content_status`` and emits ``<cp:contentStatus>`` between
``<cp:category>`` and ``<dcterms:created>`` per the OOXML core-properties
schema.
"""
from __future__ import annotations

import os
import zipfile
from pathlib import Path

import pytest


def _save_with_backend(backend: str, tmp_path: Path, status: str) -> Path:
    monkey_env = {**os.environ, "WOLFXL_WRITER": backend, "WOLFXL_TEST_EPOCH": "0"}
    out = tmp_path / f"out_{backend}.xlsx"
    import subprocess
    script = f"""
import wolfxl
wb = wolfxl.Workbook()
wb._rust_writer.set_properties({{"contentStatus": {status!r}}})
wb.save({str(out)!r})
"""
    subprocess.run(
        ["python", "-c", script],
        env=monkey_env,
        check=True,
        capture_output=True,
        text=True,
    )
    return out


@pytest.mark.parametrize("backend", ["oracle", "native"])
def test_content_status_emitted_in_core_props(
    backend: str, tmp_path: Path,
) -> None:
    """Both backends must emit ``<cp:contentStatus>Draft</cp:contentStatus>``
    in ``docProps/core.xml`` when ``contentStatus`` is set."""
    out = _save_with_backend(backend, tmp_path, "Draft")
    with zipfile.ZipFile(out) as zf:
        core = zf.read("docProps/core.xml").decode("utf-8")
    assert "<cp:contentStatus>Draft</cp:contentStatus>" in core, (
        f"backend={backend} core.xml missing contentStatus:\n{core[:500]}"
    )


@pytest.mark.parametrize("backend", ["oracle", "native"])
def test_content_status_omitted_when_unset(
    backend: str, tmp_path: Path, monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Without ``contentStatus`` in the props dict, neither backend should
    emit the element. Guards against accidental empty element emission."""
    monkeypatch.setenv("WOLFXL_WRITER", backend)
    monkeypatch.setenv("WOLFXL_TEST_EPOCH", "0")
    import wolfxl
    wb = wolfxl.Workbook()
    wb._rust_writer.set_properties({"title": "no status"})
    out = tmp_path / "out.xlsx"
    wb.save(str(out))
    with zipfile.ZipFile(out) as zf:
        core = zf.read("docProps/core.xml").decode("utf-8")
    assert "<cp:contentStatus" not in core, (
        f"backend={backend} core.xml has unexpected contentStatus:\n{core[:500]}"
    )
