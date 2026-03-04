"""Cell-level comparison — the core of the formula diff engine."""

from __future__ import annotations

from dataclasses import fields

from exceldiff.comparison.models import CellDiff, ChangeType, PropertyDiff
from exceldiff.parsing.models import CellData, CellKind, CellStyleData


def compare_cells(
    left_cells: dict[str, CellData],
    right_cells: dict[str, CellData],
    ignore_style: bool = False,
    formulas_only: bool = False,
) -> list[CellDiff]:
    """Compare two dictionaries of {coordinate: CellData}.

    Returns only cells that have differences.
    """
    diffs: list[CellDiff] = []
    all_coords = sorted(set(left_cells.keys()) | set(right_cells.keys()))

    for coord in all_coords:
        left_cell = left_cells.get(coord)
        right_cell = right_cells.get(coord)

        if left_cell is None and right_cell is not None:
            diffs.append(CellDiff(
                coordinate=coord,
                change_type=ChangeType.ADDED,
                formula_changed=right_cell.kind in (CellKind.FORMULA, CellKind.ARRAY_FORMULA),
            ))
            continue

        if left_cell is not None and right_cell is None:
            diffs.append(CellDiff(
                coordinate=coord,
                change_type=ChangeType.REMOVED,
                formula_changed=left_cell.kind in (CellKind.FORMULA, CellKind.ARRAY_FORMULA),
            ))
            continue

        # Both exist — compare properties
        assert left_cell is not None and right_cell is not None

        if left_cell == right_cell:
            continue

        diff = _compare_single_cell(left_cell, right_cell, ignore_style, formulas_only)
        if diff is not None:
            diffs.append(diff)

    return diffs


def _compare_single_cell(
    left: CellData,
    right: CellData,
    ignore_style: bool,
    formulas_only: bool,
) -> CellDiff | None:
    """Compare two CellData objects and return a CellDiff if different."""
    properties: list[PropertyDiff] = []
    formula_changed = False
    style_changed = False

    # Formula comparison (normalise to uppercase for case-insensitive comparison)
    left_formula = left.formula.upper() if left.formula else None
    right_formula = right.formula.upper() if right.formula else None
    if left_formula != right_formula:
        formula_changed = True
        properties.append(PropertyDiff(
            "formula", left.formula, right.formula, ChangeType.MODIFIED
        ))

    # Cell kind changed (formula ↔ value)
    if left.kind != right.kind:
        formula_changed = True
        properties.append(PropertyDiff(
            "cell_kind", left.kind.value, right.kind.value, ChangeType.MODIFIED
        ))

    if not formulas_only:
        # Value comparison (for non-formula cells)
        if left.kind == CellKind.VALUE and right.kind == CellKind.VALUE:
            if left.value != right.value:
                properties.append(PropertyDiff(
                    "value", left.value, right.value, ChangeType.MODIFIED
                ))

        # Data type
        if left.data_type != right.data_type:
            properties.append(PropertyDiff(
                "data_type", left.data_type, right.data_type, ChangeType.MODIFIED
            ))

        # Comment
        if left.comment != right.comment:
            properties.append(PropertyDiff(
                "comment", left.comment, right.comment, ChangeType.MODIFIED
            ))

        # Hyperlink
        if left.hyperlink != right.hyperlink:
            properties.append(PropertyDiff(
                "hyperlink", left.hyperlink, right.hyperlink, ChangeType.MODIFIED
            ))

        # Style comparison
        if not ignore_style and left.style != right.style:
            style_changed = True
            _compare_styles(left.style, right.style, properties)

    if not properties:
        return None

    return CellDiff(
        coordinate=left.coordinate,
        change_type=ChangeType.MODIFIED,
        formula_changed=formula_changed,
        style_changed=style_changed,
        properties=tuple(properties),
    )


def _compare_styles(
    left: CellStyleData,
    right: CellStyleData,
    properties: list[PropertyDiff],
) -> None:
    """Compare individual style properties and append diffs."""
    for field_obj in fields(CellStyleData):
        left_val = getattr(left, field_obj.name)
        right_val = getattr(right, field_obj.name)
        if left_val != right_val:
            properties.append(PropertyDiff(
                f"style.{field_obj.name}",
                left_val,
                right_val,
                ChangeType.MODIFIED,
            ))
