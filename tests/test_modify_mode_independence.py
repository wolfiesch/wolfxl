"""Modify mode is independent of the write-mode backend.

After the W5 rip-out, the only write-mode backend is ``NativeWorkbook``.
The modify-mode patcher (``XlsxPatcher``) has its own ZIP-rewrite path
and shares zero code with the writer. This test pins that property: if
a future change introduces a fall-through from the patcher to the
writer, the silent coupling is caught.

Two checks remain after rip-out:

1. **Source-level**: ``src/wolfxl/`` contains no ``rust_xlsxwriter``
   references. The reference is gone, but a grep enforces the absence
   in case anything reintroduces it.

2. **T1.5 raise-consistency**: T1.5-deferred features (rewriting doc
   properties, adding defined names to an existing file) raise
   ``NotImplementedError`` with a "T1.5" hint — not silent fall-through
   to the writer.
"""
from __future__ import annotations

import subprocess
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
FIXTURE = REPO / "tests" / "fixtures" / "minimal.xlsx"


def test_xlsxpatcher_has_no_rust_xlsxwriter_references() -> None:
    """``src/wolfxl/`` must never import or call ``rust_xlsxwriter``.

    The W5 rip-out removed the dependency entirely. If a future commit
    reintroduces a coupling here (e.g. a debug helper that imports the
    crate), this grep catches it.
    """
    target = REPO / "src" / "wolfxl"
    result = subprocess.run(
        ["grep", "-rln", "rust_xlsxwriter", str(target)],
        capture_output=True,
        text=True,
        check=False,
    )
    # grep returns 1 when no matches — that's the success path.
    assert result.returncode == 1, (
        "src/wolfxl/ has rust_xlsxwriter references — Wave 5 rip-out "
        "would silently break modify mode. Files with refs:\n"
        f"{result.stdout}"
    )


def test_modify_mode_t15_features_raise_with_pointer(tmp_path: Path) -> None:
    """T1.5-deferred modify-mode operations must raise
    ``NotImplementedError`` with "T1.5" in the message — never silent
    fall-through to the writer backend.

    Currently tracks two paths:
    - rewriting workbook properties on an existing file
    - adding defined names to an existing file

    If a new T1.5-deferred feature lands, append it here so the
    raise-contract is enforced at CI time.
    """
    if not FIXTURE.exists():
        pytest.skip("hermetic fixture missing")

    import wolfxl
    from wolfxl.workbook.defined_name import DefinedName

    # Path 1 — mutating wb.properties dirties properties; save raises.
    out_props = tmp_path / "props.xlsx"
    wb = wolfxl.load_workbook(str(FIXTURE), modify=True)
    wb.properties.title = "T1.5 probe"
    with pytest.raises(NotImplementedError, match=r"T1\.5"):
        wb.save(str(out_props))

    # Path 2 — adding a defined name queues a pending entry; save raises.
    out_dn = tmp_path / "dn.xlsx"
    wb2 = wolfxl.load_workbook(str(FIXTURE), modify=True)
    wb2.defined_names["probe"] = DefinedName(name="probe", value="Sheet1!$A$1")
    with pytest.raises(NotImplementedError, match=r"T1\.5"):
        wb2.save(str(out_dn))
