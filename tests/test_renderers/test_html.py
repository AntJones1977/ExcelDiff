"""Tests for HTML report renderer."""

from __future__ import annotations

from pathlib import Path

from exceldiff.comparison.engine import compare
from exceldiff.renderers.html import render


class TestHtmlRenderer:
    def test_generates_html_file(self, base_path, modified_formulas_path, tmp_path):
        diff = compare(str(base_path), str(modified_formulas_path))
        out = str(tmp_path / "report.html")
        result_path = render(diff, out)
        assert Path(result_path).exists()

    def test_html_contains_formula_diffs(self, base_path, modified_formulas_path, tmp_path):
        diff = compare(str(base_path), str(modified_formulas_path))
        out = str(tmp_path / "report.html")
        render(diff, out)
        html = Path(out).read_text(encoding="utf-8")
        assert "Formula Changes" in html
        assert "SUM" in html

    def test_html_contains_sheet_names(self, base_path, modified_structure_path, tmp_path):
        diff = compare(str(base_path), str(modified_structure_path))
        out = str(tmp_path / "report.html")
        render(diff, out)
        html = Path(out).read_text(encoding="utf-8")
        assert "Summary" in html
        assert "Config" in html
        assert "Audit" in html

    def test_html_identical_shows_no_diff(self, base_path, identical_path, tmp_path):
        diff = compare(str(base_path), str(identical_path))
        out = str(tmp_path / "report.html")
        render(diff, out)
        html = Path(out).read_text(encoding="utf-8")
        assert "No differences found" in html

    def test_html_is_self_contained(self, base_path, modified_formulas_path, tmp_path):
        diff = compare(str(base_path), str(modified_formulas_path))
        out = str(tmp_path / "report.html")
        render(diff, out)
        html = Path(out).read_text(encoding="utf-8")
        # CSS should be inlined, not linked
        assert "<style>" in html
        assert "--color-primary" in html
