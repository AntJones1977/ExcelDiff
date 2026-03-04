"""Data models for parsed Excel workbook snapshots.

All models are frozen dataclasses, completely decoupled from openpyxl.
They represent an immutable snapshot of the workbook's structure and formulae.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Optional


class CellKind(Enum):
    """What kind of content a cell holds."""
    FORMULA = "formula"
    VALUE = "value"
    EMPTY = "empty"
    ARRAY_FORMULA = "array_formula"


@dataclass(frozen=True)
class CellStyleData:
    """Normalised snapshot of a cell's formatting."""
    number_format: str = "General"
    font_name: Optional[str] = None
    font_size: Optional[float] = None
    font_bold: bool = False
    font_italic: bool = False
    font_colour: Optional[str] = None
    fill_colour: Optional[str] = None
    fill_pattern: Optional[str] = None
    border_left: Optional[str] = None
    border_right: Optional[str] = None
    border_top: Optional[str] = None
    border_bottom: Optional[str] = None
    horizontal_alignment: Optional[str] = None
    vertical_alignment: Optional[str] = None
    wrap_text: bool = False


@dataclass(frozen=True)
class CellData:
    """Complete snapshot of a single cell."""
    coordinate: str
    kind: CellKind
    formula: Optional[str] = None
    value: object = None
    data_type: str = "n"
    style: CellStyleData = field(default_factory=CellStyleData)
    comment: Optional[str] = None
    hyperlink: Optional[str] = None


@dataclass(frozen=True)
class DataValidationData:
    """Snapshot of a data validation rule."""
    type: Optional[str] = None
    formula1: Optional[str] = None
    formula2: Optional[str] = None
    cells: tuple[str, ...] = ()


@dataclass(frozen=True)
class ConditionalFormatData:
    """Snapshot of a conditional formatting rule."""
    rule_type: str = ""
    formula: Optional[str] = None
    cells: str = ""
    priority: int = 0


@dataclass(frozen=True)
class SheetData:
    """Complete snapshot of a single worksheet."""
    name: str
    dimensions: str
    max_row: int
    max_column: int
    merged_cells: tuple[str, ...] = ()
    freeze_panes: Optional[str] = None
    column_widths: dict[str, float] = field(default_factory=dict)
    row_heights: dict[int, float] = field(default_factory=dict)
    data_validations: tuple[DataValidationData, ...] = ()
    conditional_formats: tuple[ConditionalFormatData, ...] = ()
    cells: dict[str, CellData] = field(default_factory=dict)
    auto_filter: Optional[str] = None
    print_area: Optional[str] = None
    sheet_protection: bool = False


@dataclass(frozen=True)
class DefinedNameData:
    """Snapshot of a workbook-level defined name / named range."""
    name: str
    value: str
    scope: Optional[str] = None


@dataclass(frozen=True)
class WorkbookData:
    """Complete snapshot of an entire workbook."""
    file_path: str
    sheet_names: tuple[str, ...] = ()
    sheets: dict[str, SheetData] = field(default_factory=dict)
    defined_names: tuple[DefinedNameData, ...] = ()
