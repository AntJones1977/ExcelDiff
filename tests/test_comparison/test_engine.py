"""Integration tests for the comparison engine."""

from __future__ import annotations

from exceldiff.comparison.engine import compare
from exceldiff.comparison.models import ChangeType


class TestCompareIdentical:
    def test_identical_has_no_differences(self, base_path, identical_path):
        result = compare(str(base_path), str(identical_path))
        assert not result.has_differences

    def test_identical_zero_formula_diffs(self, base_path, identical_path):
        result = compare(str(base_path), str(identical_path))
        assert result.total_formula_diffs == 0

    def test_identical_all_sheets_unchanged(self, base_path, identical_path):
        result = compare(str(base_path), str(identical_path))
        for sd in result.sheet_diffs:
            assert sd.change_type == ChangeType.UNCHANGED


class TestCompareFormulas:
    def test_detects_formula_changes(self, base_path, modified_formulas_path):
        result = compare(str(base_path), str(modified_formulas_path))
        assert result.has_differences
        assert result.total_formula_diffs > 0

    def test_summary_formula_change_count(self, base_path, modified_formulas_path):
        result = compare(str(base_path), str(modified_formulas_path))
        summary = next(sd for sd in result.sheet_diffs if sd.sheet_name == "Summary")
        # B2 changed from SUM to SUMIF, B3 changed from AVERAGE to AVERAGEIF
        assert summary.formula_diff_count >= 2

    def test_data_markup_formula_changed(self, base_path, modified_formulas_path):
        result = compare(str(base_path), str(modified_formulas_path))
        data = next(sd for sd in result.sheet_diffs if sd.sheet_name == "Data")
        # C2-C10 all changed from *1.2 to *1.25
        assert data.formula_diff_count >= 9

    def test_new_formula_column_detected(self, base_path, modified_formulas_path):
        result = compare(str(base_path), str(modified_formulas_path))
        data = next(sd for sd in result.sheet_diffs if sd.sheet_name == "Data")
        # D2 is a new formula cell
        d2_diffs = [cd for cd in data.cell_diffs if cd.coordinate == "D2"]
        assert len(d2_diffs) == 1
        assert d2_diffs[0].change_type == ChangeType.ADDED

    def test_config_unchanged(self, base_path, modified_formulas_path):
        result = compare(str(base_path), str(modified_formulas_path))
        config = next(sd for sd in result.sheet_diffs if sd.sheet_name == "Config")
        assert config.change_type == ChangeType.UNCHANGED


class TestCompareStructure:
    def test_detects_removed_sheet(self, base_path, modified_structure_path):
        result = compare(str(base_path), str(modified_structure_path))
        assert "Config" in result.sheets_only_in_left

    def test_detects_added_sheet(self, base_path, modified_structure_path):
        result = compare(str(base_path), str(modified_structure_path))
        assert "Audit" in result.sheets_only_in_right

    def test_detects_renamed_sheet(self, base_path, modified_structure_path):
        result = compare(str(base_path), str(modified_structure_path))
        # "Config" removed, "Settings" added
        assert "Config" in result.sheets_only_in_left
        assert "Settings" in result.sheets_only_in_right

    def test_detects_removed_merge(self, base_path, modified_structure_path):
        result = compare(str(base_path), str(modified_structure_path))
        summary = next(sd for sd in result.sheet_diffs if sd.sheet_name == "Summary")
        merge_diffs = [d for d in summary.structural_diffs if d.property_name == "merged_cells"]
        assert any(d.change_type == ChangeType.REMOVED for d in merge_diffs)

    def test_detects_freeze_change(self, base_path, modified_structure_path):
        result = compare(str(base_path), str(modified_structure_path))
        summary = next(sd for sd in result.sheet_diffs if sd.sheet_name == "Summary")
        freeze_diffs = [d for d in summary.structural_diffs if d.property_name == "freeze_panes"]
        assert len(freeze_diffs) == 1
        assert freeze_diffs[0].left_value == "A2"
        assert freeze_diffs[0].right_value == "A3"

    def test_detects_column_width_change(self, base_path, modified_structure_path):
        result = compare(str(base_path), str(modified_structure_path))
        summary = next(sd for sd in result.sheet_diffs if sd.sheet_name == "Summary")
        width_diffs = [d for d in summary.structural_diffs if "column_width" in d.property_name]
        assert len(width_diffs) >= 1


class TestCompareFormatting:
    def test_detects_style_changes(self, base_path, modified_formatting_path):
        result = compare(str(base_path), str(modified_formatting_path))
        assert result.has_differences
        assert result.total_style_diffs > 0

    def test_no_formula_diffs_on_formatting_only(self, base_path, modified_formatting_path):
        result = compare(str(base_path), str(modified_formatting_path))
        assert result.total_formula_diffs == 0

    def test_ignore_style_flag(self, base_path, modified_formatting_path):
        result = compare(
            str(base_path), str(modified_formatting_path), ignore_style=True
        )
        assert result.total_style_diffs == 0


class TestCompareOptions:
    def test_sheet_filter(self, base_path, modified_formulas_path):
        result = compare(
            str(base_path), str(modified_formulas_path),
            sheet_filter=("Summary",)
        )
        # Only Summary should be compared
        sheet_names = [sd.sheet_name for sd in result.sheet_diffs]
        assert "Summary" in sheet_names
        assert "Data" not in sheet_names

    def test_formulas_only_skips_value_diffs(self, base_path, modified_formulas_path):
        result = compare(
            str(base_path), str(modified_formulas_path),
            formulas_only=True
        )
        # Should still detect formula changes
        assert result.total_formula_diffs > 0
