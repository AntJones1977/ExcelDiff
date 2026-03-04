"""Top-level comparison orchestrator."""

from __future__ import annotations

from exceldiff.comparison.models import WorkbookDiff
from exceldiff.comparison.workbook_comparator import compare_workbooks
from exceldiff.parsing.workbook_parser import parse_workbook


def compare(
    left_path: str,
    right_path: str,
    ignore_style: bool = False,
    formulas_only: bool = False,
    sheet_filter: tuple[str, ...] | None = None,
) -> WorkbookDiff:
    """Main entry point: parse both workbooks and compare them.

    Args:
        left_path: Path to the first (baseline) .xlsx file.
        right_path: Path to the second (comparison) .xlsx file.
        ignore_style: Skip style/formatting comparison entirely.
        formulas_only: Only compare formulas, skip values and other properties.
        sheet_filter: If provided, only compare these sheet names.

    Returns:
        A WorkbookDiff containing all differences found.
    """
    left = parse_workbook(left_path)
    right = parse_workbook(right_path)

    return compare_workbooks(
        left,
        right,
        ignore_style=ignore_style,
        formulas_only=formulas_only,
        sheet_filter=sheet_filter,
    )
