"""W4F pre-rip-out invariant: modify mode must remain functional after
``rust_xlsxwriter`` is removed in Wave 5.

The migration plan tracks one risk above all others for the rip-out: that
``XlsxPatcher`` (modify mode) might have a hidden coupling to the writer
backend — a direct call, a fall-through path, or env-var-conditional
behavior — that would silently break when ``RustXlsxWriterBook`` is
deleted.

The audit (2026-04-24) found no such coupling: ``src/wolfxl/`` imports
zero ``rust_xlsxwriter::`` symbols, T1.5-deferred features raise
``NotImplementedError`` instead of falling back to the writer, and
``WOLFXL_WRITER`` is read only by ``_backend.make_writer()`` (write mode
only). This test encodes that finding so any future change that breaks
the invariant fails CI loudly.

Three checks here:

1. **Source-level**: ``src/wolfxl/`` contains no ``rust_xlsxwriter``
   references. A grep of the canonical import shape catches future
   accidental coupling.

2. **Cross-backend invariance**: a modify-mode round-trip should produce
   byte-identical output regardless of ``WOLFXL_WRITER``, because the
   patcher is the active engine and it doesn't read the env var.

3. **T1.5 raise-consistency**: T1.5-deferred features (rewriting doc
   properties, adding defined names to an existing file) raise
   ``NotImplementedError`` with a "T1.5" hint — not silent fall-through
   to the writer.
"""
from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
FIXTURE = REPO / "tests" / "fixtures" / "minimal.xlsx"


def test_xlsxpatcher_has_no_rust_xlsxwriter_references() -> None:
    """``src/wolfxl/`` must never import or call ``rust_xlsxwriter``.

    Wave 5 will delete the rust_xlsxwriter dependency from Cargo.toml and
    every reference in the codebase. If a future commit reintroduces a
    coupling here (e.g. a debug helper that imports the crate), this
    grep catches it before the rip-out lands.
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


@pytest.mark.skipif(
    not FIXTURE.exists(),
    reason="hermetic fixture missing — regenerate via tests/fixtures/build.py",
)
def test_modify_mode_invariant_under_writer_backend(tmp_path: Path) -> None:
    """Modify-mode output must be byte-identical regardless of
    ``WOLFXL_WRITER``. The patcher path doesn't read the env var, so any
    divergence here means a coupling we missed.

    We use subprocess + ``WOLFXL_TEST_EPOCH=0`` so ZIP mtimes are
    deterministic and the byte comparison is meaningful. The script
    inside is intentionally minimal: load the fixture in modify mode,
    overwrite one cell, save.
    """
    paths: dict[str, bytes] = {}
    for backend in ("oracle", "native"):
        out = tmp_path / f"out_{backend}.xlsx"
        env = {
            **os.environ,
            "WOLFXL_WRITER": backend,
            "WOLFXL_TEST_EPOCH": "0",
        }
        script = (
            "from wolfxl import load_workbook\n"
            f"wb = load_workbook(r'{FIXTURE}', modify=True)\n"
            "wb.active['A1'] = 'modified'\n"
            f"wb.save(r'{out}')\n"
        )
        subprocess.run(
            [sys.executable, "-c", script],
            env=env,
            check=True,
            capture_output=True,
        )
        paths[backend] = out.read_bytes()

    assert paths["oracle"] == paths["native"], (
        "modify-mode bytes differ between WOLFXL_WRITER=oracle and "
        "WOLFXL_WRITER=native — the patcher should be invariant of the "
        "writer selection. Wave 5 rip-out would expose this drift."
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
