"""Tests for workbook parsing."""

from __future__ import annotations

from exceldiff.parsing.models import CellKind
from exceldiff.parsing.workbook_parser import parse_workbook


class TestParseWorkbook:
    def test_parse_sheet_names(self, base_path):
        wb = parse_workbook(str(base_path))
        assert wb.sheet_names == ("Summary", "Data", "Config")

    def test_parse_sheet_count(self, base_path):
        wb = parse_workbook(str(base_path))
        assert len(wb.sheets) == 3

    def test_parse_formulas_not_values(self, base_path):
        wb = parse_workbook(str(base_path))
        cell = wb.sheets["Summary"].cells["B2"]
        assert cell.kind == CellKind.FORMULA
        assert cell.formula == "=SUM(Data!B2:B10)"

    def test_parse_value_cells(self, base_path):
        wb = parse_workbook(str(base_path))
        cell = wb.sheets["Config"].cells["B1"]
        assert cell.kind == CellKind.VALUE
        assert cell.value == 0.2

    def test_parse_string_cells(self, base_path):
        wb = parse_workbook(str(base_path))
        cell = wb.sheets["Data"].cells["A2"]
        assert cell.kind == CellKind.VALUE
        assert cell.value == "Item 1"

    def test_parse_merged_cells(self, base_path):
        wb = parse_workbook(str(base_path))
        assert "A6:B6" in wb.sheets["Summary"].merged_cells

    def test_parse_freeze_panes(self, base_path):
        wb = parse_workbook(str(base_path))
        assert wb.sheets["Summary"].freeze_panes == "A2"

    def test_parse_column_widths(self, base_path):
        wb = parse_workbook(str(base_path))
        widths = wb.sheets["Summary"].column_widths
        assert "A" in widths
        assert widths["A"] == 20

    def test_parse_comment(self, base_path):
        wb = parse_workbook(str(base_path))
        cell = wb.sheets["Summary"].cells["A1"]
        assert cell.comment == "Header column"

    def test_parse_empty_workbook(self, empty_path):
        wb = parse_workbook(str(empty_path))
        assert wb.sheet_names == ("Sheet",)
        assert len(wb.sheets["Sheet"].cells) == 0

    def test_parse_all_data_formulas(self, base_path):
        wb = parse_workbook(str(base_path))
        data_sheet = wb.sheets["Data"]
        # C2 through C10 should all be formulas
        for i in range(2, 11):
            cell = data_sheet.cells[f"C{i}"]
            assert cell.kind == CellKind.FORMULA
            assert f"B{i}*1.2" in cell.formula
