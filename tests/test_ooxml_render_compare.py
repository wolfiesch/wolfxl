from __future__ import annotations

import importlib.util
import sys
from pathlib import Path
from types import ModuleType

import openpyxl


def _load_render_module() -> ModuleType:
    script = (
        Path(__file__).resolve().parents[1] / "scripts" / "run_ooxml_render_compare.py"
    )
    spec = importlib.util.spec_from_file_location("run_ooxml_render_compare", script)
    assert spec is not None
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


render_module = _load_render_module()


def _make_fixture(path: Path) -> None:
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet["A1"] = "ok"
    workbook.save(path)


def test_render_compare_skips_when_renderer_tools_missing(
    tmp_path: Path, monkeypatch
) -> None:
    fixture_dir = tmp_path / "fixtures"
    output_dir = tmp_path / "out"
    fixture_dir.mkdir()
    _make_fixture(fixture_dir / "simple.xlsx")
    monkeypatch.setattr(
        render_module.run_ooxml_app_smoke,
        "_find_libreoffice",
        lambda: None,
    )

    report = render_module.run_render_compare(fixture_dir, output_dir, timeout=1)

    assert report["result_count"] == 1
    assert report["failure_count"] == 0
    assert report["results"][0]["status"] == "skipped"
    assert "soffice not found" in report["results"][0]["message"]


def test_render_compare_can_discover_recursive_fixture_trees(
    tmp_path: Path, monkeypatch
) -> None:
    fixture_dir = tmp_path / "fixtures"
    output_dir = tmp_path / "out"
    nested_dir = fixture_dir / "nested" / "deep"
    nested_dir.mkdir(parents=True)
    _make_fixture(nested_dir / "simple.xlsx")

    monkeypatch.setattr(
        render_module.run_ooxml_app_smoke,
        "_find_libreoffice",
        lambda: None,
    )

    report = render_module.run_render_compare(
        fixture_dir,
        output_dir,
        timeout=1,
        recursive=True,
    )

    assert report["recursive"] is True
    assert report["result_count"] == 1
    result = report["results"][0]
    assert result["fixture"] == "nested/deep/simple.xlsx"
    assert result["status"] == "skipped"


def test_render_compare_reports_rmse_threshold_failure(tmp_path: Path, monkeypatch) -> None:
    fixture_dir = tmp_path / "fixtures"
    output_dir = tmp_path / "out"
    fixture_dir.mkdir()
    _make_fixture(fixture_dir / "simple.xlsx")
    before_pdf = tmp_path / "before.pdf"
    after_pdf = tmp_path / "after.pdf"
    before_pdf.write_bytes(b"%PDF-before")
    after_pdf.write_bytes(b"%PDF-after")
    before_page = tmp_path / "before-1.png"
    after_page = tmp_path / "after-1.png"
    before_page.write_bytes(b"before")
    after_page.write_bytes(b"after")

    monkeypatch.setattr(
        render_module.run_ooxml_app_smoke, "_find_libreoffice", lambda: "soffice"
    )
    monkeypatch.setattr(render_module.shutil, "which", lambda name: name)
    monkeypatch.setattr(
        render_module,
        "_export_pdf",
        lambda _soffice, src, _outdir, _timeout: before_pdf
        if src.name.startswith("before-")
        else after_pdf,
    )
    monkeypatch.setattr(
        render_module,
        "_rasterize_pdf",
        lambda _pdftoppm, pdf, _prefix, _density, _timeout: [before_page]
        if pdf == before_pdf
        else [after_page],
    )
    monkeypatch.setattr(
        render_module,
        "_normalized_rmse",
        lambda _compare_cmd, _before_page, _after_page, _timeout: 0.25,
    )

    report = render_module.run_render_compare(
        fixture_dir,
        output_dir,
        timeout=1,
        max_normalized_rmse=0.01,
    )

    assert report["failure_count"] == 1
    result = report["results"][0]
    assert result["status"] == "failed"
    assert result["max_normalized_rmse"] == 0.25
    assert "render drift above threshold" in result["message"]


def test_normalized_rmse_parses_imagemagick_metric(tmp_path: Path, monkeypatch) -> None:
    before_page = tmp_path / "before.png"
    after_page = tmp_path / "after.png"
    before_page.write_bytes(b"before")
    after_page.write_bytes(b"after")

    class Completed:
        returncode = 1
        stderr = "123.4 (0.001883)"

    monkeypatch.setattr(
        render_module.subprocess,
        "run",
        lambda *args, **kwargs: Completed(),
    )

    assert (
        render_module._normalized_rmse(("compare",), before_page, after_page, timeout=1)
        == 0.001883
    )
