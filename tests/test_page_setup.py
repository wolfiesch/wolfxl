"""RFC-055 §2.1 — PageSetup tests (Sprint Ο Pod 1A)."""

from __future__ import annotations

import pytest

from wolfxl import Workbook
from wolfxl.worksheet.page_setup import PageMargins, PageSetup


class TestPageSetupDefaults:
    def test_defaults_are_construction_defaults(self):
        ps = PageSetup()
        assert ps.orientation == "default"
        assert ps.paperSize is None
        assert ps.scale is None

    def test_is_default_returns_true_for_fresh(self):
        assert PageSetup().is_default()

    def test_is_default_returns_false_after_mutation(self):
        ps = PageSetup()
        ps.orientation = "landscape"
        # Re-validate after mutation by reconstructing:
        ps2 = PageSetup(orientation="landscape")
        assert not ps2.is_default()


class TestPageSetupValidation:
    def test_invalid_orientation_raises(self):
        with pytest.raises(ValueError, match="orientation"):
            PageSetup(orientation="sideways")

    def test_scale_out_of_range_raises_low(self):
        with pytest.raises(ValueError, match="scale"):
            PageSetup(scale=5)

    def test_scale_out_of_range_raises_high(self):
        with pytest.raises(ValueError, match="scale"):
            PageSetup(scale=500)

    def test_scale_in_range_accepted(self):
        ps = PageSetup(scale=200)
        assert ps.scale == 200

    def test_invalid_cell_comments_raises(self):
        with pytest.raises(ValueError, match="cellComments"):
            PageSetup(cellComments="invalid")

    def test_invalid_errors_raises(self):
        with pytest.raises(ValueError, match="errors"):
            PageSetup(errors="oops")


class TestPageSetupAliases:
    def test_paper_size_alias(self):
        ps = PageSetup(paperSize=9)
        assert ps.paper_size == 9
        ps.paper_size = 1
        assert ps.paperSize == 1

    def test_fit_to_width_alias(self):
        ps = PageSetup()
        ps.fit_to_width = 1
        assert ps.fitToWidth == 1
        assert ps.fit_to_width == 1

    def test_fit_to_height_alias(self):
        ps = PageSetup()
        ps.fit_to_height = 2
        assert ps.fitToHeight == 2


class TestPageSetupRustDict:
    def test_default_orientation_emits_none(self):
        d = PageSetup().to_rust_dict()
        assert d["orientation"] is None

    def test_explicit_orientation_emits_string(self):
        d = PageSetup(orientation="landscape").to_rust_dict()
        assert d["orientation"] == "landscape"

    def test_paper_size_round_trip(self):
        d = PageSetup(paperSize=9).to_rust_dict()
        assert d["paper_size"] == 9


class TestWorksheetPageSetupAccessor:
    def test_lazy_access_returns_default_instance(self):
        wb = Workbook()
        ws = wb.active
        ps = ws.page_setup
        assert isinstance(ps, PageSetup)
        assert ps.is_default()

    def test_assignment_replaces_instance(self):
        wb = Workbook()
        ws = wb.active
        new_ps = PageSetup(orientation="landscape", scale=150)
        ws.page_setup = new_ps
        assert ws.page_setup is new_ps
        assert ws.page_setup.orientation == "landscape"

    def test_mutation_persists_across_access(self):
        wb = Workbook()
        ws = wb.active
        ws.page_setup.orientation = "portrait"
        assert ws.page_setup.orientation == "portrait"


class TestPageMarginsBasics:
    def test_defaults(self):
        pm = PageMargins()
        assert pm.left == 0.7
        assert pm.right == 0.7
        assert pm.top == 0.75
        assert pm.bottom == 0.75
        assert pm.header == 0.3
        assert pm.footer == 0.3

    def test_mutation(self):
        pm = PageMargins(top=1.0, bottom=1.0)
        assert pm.top == 1.0
        assert pm.bottom == 1.0
        assert not pm.is_default()

    def test_to_rust_dict(self):
        pm = PageMargins(top=1.5, footer=0.5)
        d = pm.to_rust_dict()
        assert d["top"] == 1.5
        assert d["footer"] == 0.5
        assert d["left"] == 0.7
