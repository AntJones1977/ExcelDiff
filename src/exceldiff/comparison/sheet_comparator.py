"""Sheet-level structural comparison."""

from __future__ import annotations

from exceldiff.comparison.cell_comparator import compare_cells
from exceldiff.comparison.models import ChangeType, PropertyDiff, SheetDiff
from exceldiff.parsing.models import SheetData


def compare_sheets(
    left: SheetData,
    right: SheetData,
    ignore_style: bool = False,
    formulas_only: bool = False,
) -> SheetDiff:
    """Compare two SheetData snapshots and return a SheetDiff."""
    structural_diffs: list[PropertyDiff] = []

    # Dimensions
    if left.dimensions != right.dimensions:
        structural_diffs.append(PropertyDiff(
            "dimensions", left.dimensions, right.dimensions, ChangeType.MODIFIED
        ))

    # Merged cells (set comparison)
    left_merged = set(left.merged_cells)
    right_merged = set(right.merged_cells)
    for removed in sorted(left_merged - right_merged):
        structural_diffs.append(PropertyDiff(
            "merged_cells", removed, None, ChangeType.REMOVED
        ))
    for added in sorted(right_merged - left_merged):
        structural_diffs.append(PropertyDiff(
            "merged_cells", None, added, ChangeType.ADDED
        ))

    # Freeze panes
    if left.freeze_panes != right.freeze_panes:
        structural_diffs.append(PropertyDiff(
            "freeze_panes", left.freeze_panes, right.freeze_panes, ChangeType.MODIFIED
        ))

    # Column widths
    _compare_dicts(left.column_widths, right.column_widths, "column_width", structural_diffs)

    # Row heights
    _compare_dicts(left.row_heights, right.row_heights, "row_height", structural_diffs)

    # Auto filter
    if left.auto_filter != right.auto_filter:
        structural_diffs.append(PropertyDiff(
            "auto_filter", left.auto_filter, right.auto_filter, ChangeType.MODIFIED
        ))

    # Print area
    if left.print_area != right.print_area:
        structural_diffs.append(PropertyDiff(
            "print_area", left.print_area, right.print_area, ChangeType.MODIFIED
        ))

    # Sheet protection
    if left.sheet_protection != right.sheet_protection:
        structural_diffs.append(PropertyDiff(
            "sheet_protection", left.sheet_protection, right.sheet_protection, ChangeType.MODIFIED
        ))

    # Data validations
    if left.data_validations != right.data_validations:
        structural_diffs.append(PropertyDiff(
            "data_validations",
            f"{len(left.data_validations)} rules",
            f"{len(right.data_validations)} rules",
            ChangeType.MODIFIED,
        ))

    # Conditional formatting
    if left.conditional_formats != right.conditional_formats:
        structural_diffs.append(PropertyDiff(
            "conditional_formatting",
            f"{len(left.conditional_formats)} rules",
            f"{len(right.conditional_formats)} rules",
            ChangeType.MODIFIED,
        ))

    # Cell-level comparison
    cell_diffs = compare_cells(left.cells, right.cells, ignore_style, formulas_only)

    change_type = (
        ChangeType.MODIFIED
        if structural_diffs or cell_diffs
        else ChangeType.UNCHANGED
    )

    return SheetDiff(
        sheet_name=left.name,
        change_type=change_type,
        structural_diffs=tuple(structural_diffs),
        cell_diffs=tuple(cell_diffs),
    )


def _compare_dicts(
    left: dict,
    right: dict,
    prefix: str,
    diffs: list[PropertyDiff],
) -> None:
    """Compare two dictionaries key-by-key, appending PropertyDiff for differences."""
    all_keys = sorted(set(left.keys()) | set(right.keys()), key=str)
    for key in all_keys:
        left_val = left.get(key)
        right_val = right.get(key)
        if left_val is None and right_val is not None:
            diffs.append(PropertyDiff(f"{prefix}[{key}]", None, right_val, ChangeType.ADDED))
        elif left_val is not None and right_val is None:
            diffs.append(PropertyDiff(f"{prefix}[{key}]", left_val, None, ChangeType.REMOVED))
        elif left_val != right_val:
            diffs.append(PropertyDiff(f"{prefix}[{key}]", left_val, right_val, ChangeType.MODIFIED))
