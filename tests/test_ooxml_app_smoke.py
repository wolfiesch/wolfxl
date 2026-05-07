from __future__ import annotations

import importlib.util
import sys
from pathlib import Path
from types import ModuleType

import openpyxl


def _load_smoke_module() -> ModuleType:
    script = Path(__file__).resolve().parents[1] / "scripts" / "run_ooxml_app_smoke.py"
    spec = importlib.util.spec_from_file_location("run_ooxml_app_smoke", script)
    assert spec is not None
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


smoke_module = _load_smoke_module()


def _make_fixture(path: Path) -> None:
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet["A1"] = "ok"
    workbook.save(path)


def test_validate_xlsx_rejects_missing_output(tmp_path: Path) -> None:
    ok, message = smoke_module._validate_xlsx(tmp_path / "missing.xlsx")

    assert not ok
    assert "no output file" in message


def test_validate_xlsx_accepts_basic_workbook(tmp_path: Path) -> None:
    fixture = tmp_path / "simple.xlsx"
    _make_fixture(fixture)

    ok, message = smoke_module._validate_xlsx(fixture)

    assert ok
    assert message == "ok"


def test_smoke_skips_libreoffice_when_missing(tmp_path: Path, monkeypatch) -> None:
    fixture = tmp_path / "simple.xlsx"
    _make_fixture(fixture)
    monkeypatch.setattr(smoke_module, "_find_libreoffice", lambda: None)

    result = smoke_module._smoke_libreoffice(fixture, tmp_path / "out", timeout=1)

    assert result.status == "skipped"
    assert result.app == "libreoffice"


def test_smoke_excel_rejects_unrelated_active_workbook(
    tmp_path: Path, monkeypatch
) -> None:
    fixture = tmp_path / "simple.xlsx"
    _make_fixture(fixture)
    excel_app = tmp_path / "Microsoft Excel.app"
    excel_app.mkdir()
    monkeypatch.setattr(smoke_module, "EXCEL_APP", str(excel_app))
    monkeypatch.setattr(
        smoke_module,
        "_open_excel_with_finder_and_close",
        lambda _src, _timeout: "Book1",
    )

    result = smoke_module._smoke_excel(fixture, tmp_path / "out", timeout=1)

    assert result.status == "failed"
    assert "opened 'Book1', expected 'simple.xlsx'" in result.message


def test_smoke_excel_accepts_expected_active_workbook(
    tmp_path: Path, monkeypatch
) -> None:
    fixture = tmp_path / "simple.xlsx"
    _make_fixture(fixture)
    excel_app = tmp_path / "Microsoft Excel.app"
    excel_app.mkdir()
    monkeypatch.setattr(smoke_module, "EXCEL_APP", str(excel_app))
    monkeypatch.setattr(
        smoke_module,
        "_open_excel_with_finder_and_close",
        lambda src, _timeout: src.name,
    )

    result = smoke_module._smoke_excel(fixture, tmp_path / "out", timeout=1)

    assert result.status == "passed"
    assert result.message == "opened and closed in Microsoft Excel: simple.xlsx"


def test_excel_active_workbook_name_treats_missing_value_as_none(monkeypatch) -> None:
    class FakeProc:
        stdout = "missing value\n"

    monkeypatch.setattr(
        smoke_module.subprocess,
        "run",
        lambda *_args, **_kwargs: FakeProc(),
    )

    assert smoke_module._excel_active_workbook_name() is None


def test_run_smoke_reports_failure_count(tmp_path: Path, monkeypatch) -> None:
    fixture_dir = tmp_path / "fixtures"
    output_dir = tmp_path / "out"
    fixture_dir.mkdir()
    fixture = fixture_dir / "simple.xlsx"
    _make_fixture(fixture)

    def fake_smoke(src: Path, _output_dir: Path, _timeout: int):
        return smoke_module.AppSmokeResult(
            fixture=src.name,
            mutation="source",
            app="libreoffice",
            status="failed",
            output=None,
            message="simulated failure",
        )

    monkeypatch.setattr(smoke_module, "_smoke_libreoffice", fake_smoke)

    report = smoke_module.run_smoke(
        fixture_dir,
        output_dir,
        apps=("libreoffice",),
        timeout=1,
    )

    assert report["result_count"] == 1
    assert report["failure_count"] == 1
    assert report["mutations"] == ["source"]
    assert (output_dir / "app-smoke-report.json").is_file()


def test_run_smoke_can_apply_mutation_before_app_smoke(
    tmp_path: Path, monkeypatch
) -> None:
    fixture_dir = tmp_path / "fixtures"
    output_dir = tmp_path / "out"
    fixture_dir.mkdir()
    fixture = fixture_dir / "simple.xlsx"
    _make_fixture(fixture)
    seen_sources: list[Path] = []

    def fake_smoke(src: Path, _output_dir: Path, _timeout: int):
        seen_sources.append(src)
        return smoke_module.AppSmokeResult(
            fixture=src.name,
            mutation="source",
            app="libreoffice",
            status="passed",
            output=str(src),
            message="ok",
        )

    monkeypatch.setattr(smoke_module, "_smoke_libreoffice", fake_smoke)

    report = smoke_module.run_smoke(
        fixture_dir,
        output_dir,
        apps=("libreoffice",),
        timeout=1,
        mutations=("marker_cell",),
    )

    assert report["result_count"] == 1
    assert report["failure_count"] == 0
    assert report["mutations"] == ["marker_cell"]
    result = report["results"][0]
    assert result["fixture"] == "simple.xlsx"
    assert result["mutation"] == "marker_cell"
    assert seen_sources[0].name == "after-simple.xlsx"
    assert seen_sources[0].is_file()
