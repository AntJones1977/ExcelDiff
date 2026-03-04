"""Data models for diff results between two workbooks."""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any


class ChangeType(Enum):
    """Classification of a change between two workbooks."""
    ADDED = "added"
    REMOVED = "removed"
    MODIFIED = "modified"
    UNCHANGED = "unchanged"


@dataclass(frozen=True)
class PropertyDiff:
    """A single property that differs between two objects."""
    property_name: str
    left_value: Any
    right_value: Any
    change_type: ChangeType


@dataclass(frozen=True)
class CellDiff:
    """Diff result for a single cell."""
    coordinate: str
    change_type: ChangeType
    formula_changed: bool = False
    style_changed: bool = False
    properties: tuple[PropertyDiff, ...] = ()


@dataclass(frozen=True)
class SheetDiff:
    """Diff result for a single sheet."""
    sheet_name: str
    change_type: ChangeType
    structural_diffs: tuple[PropertyDiff, ...] = ()
    cell_diffs: tuple[CellDiff, ...] = ()

    @property
    def formula_diff_count(self) -> int:
        return sum(1 for cd in self.cell_diffs if cd.formula_changed)

    @property
    def style_diff_count(self) -> int:
        return sum(1 for cd in self.cell_diffs if cd.style_changed)

    @property
    def total_diff_count(self) -> int:
        return len(self.cell_diffs)


@dataclass(frozen=True)
class WorkbookDiff:
    """Top-level diff result for two workbooks."""
    left_path: str
    right_path: str
    sheet_diffs: tuple[SheetDiff, ...] = ()
    defined_name_diffs: tuple[PropertyDiff, ...] = ()
    sheets_only_in_left: tuple[str, ...] = ()
    sheets_only_in_right: tuple[str, ...] = ()
    sheets_in_both: tuple[str, ...] = ()

    @property
    def has_differences(self) -> bool:
        return (
            len(self.sheets_only_in_left) > 0
            or len(self.sheets_only_in_right) > 0
            or any(sd.change_type != ChangeType.UNCHANGED for sd in self.sheet_diffs)
            or len(self.defined_name_diffs) > 0
        )

    @property
    def total_formula_diffs(self) -> int:
        return sum(sd.formula_diff_count for sd in self.sheet_diffs)

    @property
    def total_cell_diffs(self) -> int:
        return sum(sd.total_diff_count for sd in self.sheet_diffs)

    @property
    def total_style_diffs(self) -> int:
        return sum(sd.style_diff_count for sd in self.sheet_diffs)
