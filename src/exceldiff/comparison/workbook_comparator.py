"""Workbook-level comparison — sheet set operations and defined names."""

from __future__ import annotations

from exceldiff.comparison.models import ChangeType, PropertyDiff, SheetDiff, WorkbookDiff
from exceldiff.comparison.sheet_comparator import compare_sheets
from exceldiff.parsing.models import WorkbookData


def compare_workbooks(
    left: WorkbookData,
    right: WorkbookData,
    ignore_style: bool = False,
    formulas_only: bool = False,
    sheet_filter: tuple[str, ...] | None = None,
) -> WorkbookDiff:
    """Compare two WorkbookData snapshots and return a WorkbookDiff."""
    left_names = set(left.sheet_names)
    right_names = set(right.sheet_names)

    only_in_left = sorted(left_names - right_names)
    only_in_right = sorted(right_names - left_names)
    in_both = [name for name in left.sheet_names if name in right_names]

    # Apply sheet filter if provided
    if sheet_filter:
        filter_set = set(sheet_filter)
        only_in_left = [n for n in only_in_left if n in filter_set]
        only_in_right = [n for n in only_in_right if n in filter_set]
        in_both = [n for n in in_both if n in filter_set]

    # Build sheet diffs
    sheet_diffs: list[SheetDiff] = []

    # Sheets only in left (removed)
    for name in only_in_left:
        sheet_diffs.append(SheetDiff(
            sheet_name=name,
            change_type=ChangeType.REMOVED,
        ))

    # Sheets only in right (added)
    for name in only_in_right:
        sheet_diffs.append(SheetDiff(
            sheet_name=name,
            change_type=ChangeType.ADDED,
        ))

    # Sheets in both — compare
    for name in in_both:
        diff = compare_sheets(
            left.sheets[name],
            right.sheets[name],
            ignore_style=ignore_style,
            formulas_only=formulas_only,
        )
        sheet_diffs.append(diff)

    # Compare defined names
    defined_name_diffs = _compare_defined_names(left, right)

    return WorkbookDiff(
        left_path=left.file_path,
        right_path=right.file_path,
        sheet_diffs=tuple(sheet_diffs),
        defined_name_diffs=tuple(defined_name_diffs),
        sheets_only_in_left=tuple(only_in_left),
        sheets_only_in_right=tuple(only_in_right),
        sheets_in_both=tuple(in_both),
    )


def _compare_defined_names(
    left: WorkbookData,
    right: WorkbookData,
) -> list[PropertyDiff]:
    """Compare defined names between two workbooks."""
    diffs: list[PropertyDiff] = []

    left_names = {dn.name: dn for dn in left.defined_names}
    right_names = {dn.name: dn for dn in right.defined_names}

    all_names = sorted(set(left_names.keys()) | set(right_names.keys()))

    for name in all_names:
        left_dn = left_names.get(name)
        right_dn = right_names.get(name)

        if left_dn is None and right_dn is not None:
            diffs.append(PropertyDiff(
                f"defined_name[{name}]", None, right_dn.value, ChangeType.ADDED
            ))
        elif left_dn is not None and right_dn is None:
            diffs.append(PropertyDiff(
                f"defined_name[{name}]", left_dn.value, None, ChangeType.REMOVED
            ))
        elif left_dn is not None and right_dn is not None and left_dn != right_dn:
            diffs.append(PropertyDiff(
                f"defined_name[{name}]", left_dn.value, right_dn.value, ChangeType.MODIFIED
            ))

    return diffs
