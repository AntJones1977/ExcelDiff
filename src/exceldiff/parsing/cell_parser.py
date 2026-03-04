"""Extract cell-level data from openpyxl cells."""

from __future__ import annotations

from typing import Optional

from openpyxl.cell import Cell
from openpyxl.styles import Font, PatternFill, Border, Alignment
from openpyxl.worksheet.formula import ArrayFormula

from exceldiff.parsing.models import CellData, CellKind, CellStyleData


# Default openpyxl style objects for comparison
_DEFAULT_FONT = Font()
_DEFAULT_FILL = PatternFill()
_DEFAULT_BORDER = Border()
_DEFAULT_ALIGNMENT = Alignment()


def parse_cell(cell: Cell) -> Optional[CellData]:
    """Parse a single openpyxl cell into a CellData snapshot.

    Returns None for truly empty cells (no value, no non-default style, no comment).
    """
    value = cell.value
    has_style = _has_non_default_style(cell)
    has_comment = cell.comment is not None

    # Determine cell kind and formula
    if value is None:
        if has_style or has_comment:
            kind = CellKind.EMPTY
        else:
            return None
        formula = None
        raw_value = None
    elif isinstance(value, ArrayFormula):
        kind = CellKind.ARRAY_FORMULA
        formula = value.text
        raw_value = None
    elif isinstance(value, str) and len(value) > 1 and value.startswith("="):
        kind = CellKind.FORMULA
        formula = value
        raw_value = None
    else:
        kind = CellKind.VALUE
        formula = None
        raw_value = value

    style = _extract_style(cell)
    comment_text = cell.comment.text if cell.comment else None
    hyperlink_target = cell.hyperlink.target if cell.hyperlink else None

    return CellData(
        coordinate=cell.coordinate,
        kind=kind,
        formula=formula,
        value=raw_value,
        data_type=cell.data_type,
        style=style,
        comment=comment_text,
        hyperlink=hyperlink_target,
    )


def _has_non_default_style(cell: Cell) -> bool:
    """Check if a cell has any non-default styling."""
    return (
        cell.font != _DEFAULT_FONT
        or cell.fill != _DEFAULT_FILL
        or cell.border != _DEFAULT_BORDER
        or cell.alignment != _DEFAULT_ALIGNMENT
        or cell.number_format != "General"
    )


def _extract_style(cell: Cell) -> CellStyleData:
    """Extract style properties into a normalised CellStyleData."""
    font = cell.font
    fill = cell.fill
    border = cell.border
    alignment = cell.alignment

    return CellStyleData(
        number_format=cell.number_format or "General",
        font_name=font.name,
        font_size=font.size,
        font_bold=font.bold or False,
        font_italic=font.italic or False,
        font_colour=_colour_to_hex(font.color),
        fill_colour=_colour_to_hex(fill.fgColor) if fill.patternType else None,
        fill_pattern=fill.patternType,
        border_left=border.left.style if border.left and border.left.style else None,
        border_right=border.right.style if border.right and border.right.style else None,
        border_top=border.top.style if border.top and border.top.style else None,
        border_bottom=border.bottom.style if border.bottom and border.bottom.style else None,
        horizontal_alignment=alignment.horizontal,
        vertical_alignment=alignment.vertical,
        wrap_text=alignment.wrap_text or False,
    )


def _colour_to_hex(colour) -> Optional[str]:
    """Convert an openpyxl colour to a hex string, or None."""
    if colour is None:
        return None
    if colour.type == "rgb" and colour.rgb and colour.rgb != "00000000":
        return str(colour.rgb)
    if colour.type == "theme":
        return f"theme:{colour.theme}+{colour.tint}"
    if colour.type == "indexed" and colour.indexed is not None:
        return f"indexed:{colour.indexed}"
    return None
