"""Rich console output renderer for WorkbookDiff results."""

from __future__ import annotations

from pathlib import Path

from rich.console import Console
from rich.panel import Panel
from rich.table import Table
from rich.text import Text

from exceldiff.comparison.models import ChangeType, CellDiff, SheetDiff, WorkbookDiff


# Colour scheme
_COLOURS = {
    ChangeType.ADDED: "green",
    ChangeType.REMOVED: "red",
    ChangeType.MODIFIED: "yellow",
    ChangeType.UNCHANGED: "dim",
}


def render(diff: WorkbookDiff, verbose: bool = False) -> None:
    """Render a WorkbookDiff to the terminal using Rich."""
    console = Console()

    # Summary panel
    _render_summary(console, diff)

    if not diff.has_differences:
        console.print("\n[bold green]No differences found.[/bold green]\n")
        return

    # Defined name diffs
    if diff.defined_name_diffs:
        _render_defined_names(console, diff)

    # Sheet-level summary table
    _render_sheet_summary(console, diff)

    # Per-sheet details
    for sheet_diff in diff.sheet_diffs:
        if sheet_diff.change_type == ChangeType.UNCHANGED:
            continue
        _render_sheet_detail(console, sheet_diff, verbose)


def _render_summary(console: Console, diff: WorkbookDiff) -> None:
    """Render the top-level summary panel."""
    left_name = Path(diff.left_path).name
    right_name = Path(diff.right_path).name

    summary_lines = [
        f"[bold]Left:[/bold]  {left_name}",
        f"[bold]Right:[/bold] {right_name}",
        "",
        f"Sheets compared: [bold]{len(diff.sheets_in_both)}[/bold]    "
        f"Added: [green]{len(diff.sheets_only_in_right)}[/green]    "
        f"Removed: [red]{len(diff.sheets_only_in_left)}[/red]",
        f"Formula diffs: [bold yellow]{diff.total_formula_diffs}[/bold yellow]    "
        f"Style diffs: [dim]{diff.total_style_diffs}[/dim]    "
        f"Total cell diffs: [bold]{diff.total_cell_diffs}[/bold]",
    ]

    panel = Panel(
        "\n".join(summary_lines),
        title="[bold]ExcelDiff Report[/bold]",
        border_style="blue",
        padding=(1, 2),
    )
    console.print(panel)


def _render_defined_names(console: Console, diff: WorkbookDiff) -> None:
    """Render defined name differences."""
    table = Table(title="Defined Name Changes", border_style="blue")
    table.add_column("Name", style="bold")
    table.add_column("Left Value")
    table.add_column("Right Value")
    table.add_column("Change", justify="center")

    for dn_diff in diff.defined_name_diffs:
        colour = _COLOURS[dn_diff.change_type]
        table.add_row(
            dn_diff.property_name,
            str(dn_diff.left_value or "-"),
            str(dn_diff.right_value or "-"),
            Text(dn_diff.change_type.value, style=colour),
        )

    console.print(table)
    console.print()


def _render_sheet_summary(console: Console, diff: WorkbookDiff) -> None:
    """Render a summary table of all sheets."""
    table = Table(title="Sheet Summary", border_style="blue")
    table.add_column("Sheet Name", style="bold")
    table.add_column("Status", justify="center")
    table.add_column("Formula Diffs", justify="right")
    table.add_column("Style Diffs", justify="right")
    table.add_column("Structural Diffs", justify="right")

    for sd in diff.sheet_diffs:
        colour = _COLOURS[sd.change_type]
        table.add_row(
            sd.sheet_name,
            Text(sd.change_type.value, style=colour),
            str(sd.formula_diff_count) if sd.change_type == ChangeType.MODIFIED else "-",
            str(sd.style_diff_count) if sd.change_type == ChangeType.MODIFIED else "-",
            str(len(sd.structural_diffs)) if sd.change_type == ChangeType.MODIFIED else "-",
        )

    console.print(table)
    console.print()


def _render_sheet_detail(console: Console, sd: SheetDiff, verbose: bool) -> None:
    """Render detailed diff for a single sheet."""
    colour = _COLOURS[sd.change_type]
    console.rule(f"[{colour} bold]Sheet: \"{sd.sheet_name}\" [{sd.change_type.value}][/{colour} bold]")

    if sd.change_type in (ChangeType.ADDED, ChangeType.REMOVED):
        console.print(f"  Sheet was [bold {colour}]{sd.change_type.value}[/bold {colour}]\n")
        return

    # Structural diffs
    if sd.structural_diffs:
        table = Table(title="Structural Changes", border_style="cyan")
        table.add_column("Property", style="bold")
        table.add_column("Left")
        table.add_column("Right")
        table.add_column("Change", justify="center")

        for prop in sd.structural_diffs:
            prop_colour = _COLOURS[prop.change_type]
            table.add_row(
                prop.property_name,
                str(prop.left_value) if prop.left_value is not None else "-",
                str(prop.right_value) if prop.right_value is not None else "-",
                Text(prop.change_type.value, style=prop_colour),
            )
        console.print(table)

    # Formula diffs
    formula_diffs = [cd for cd in sd.cell_diffs if cd.formula_changed]
    if formula_diffs:
        table = Table(title="Formula Changes", border_style="yellow")
        table.add_column("Cell", style="bold cyan")
        table.add_column("Left Formula")
        table.add_column("Right Formula")
        table.add_column("Change", justify="center")

        for cd in formula_diffs:
            _add_cell_formula_row(table, cd)

        console.print(table)

    # Style diffs (only in verbose mode)
    style_diffs = [cd for cd in sd.cell_diffs if cd.style_changed and not cd.formula_changed]
    if verbose and style_diffs:
        table = Table(title="Style Changes (verbose)", border_style="dim")
        table.add_column("Cell", style="bold")
        table.add_column("Property")
        table.add_column("Left")
        table.add_column("Right")

        for cd in style_diffs:
            for prop in cd.properties:
                if prop.property_name.startswith("style."):
                    table.add_row(
                        cd.coordinate,
                        prop.property_name.removeprefix("style."),
                        str(prop.left_value),
                        str(prop.right_value),
                    )
        console.print(table)

    # Non-formula, non-style diffs
    other_diffs = [
        cd for cd in sd.cell_diffs
        if not cd.formula_changed and not cd.style_changed
    ]
    if other_diffs:
        table = Table(title="Other Cell Changes", border_style="blue")
        table.add_column("Cell", style="bold")
        table.add_column("Property")
        table.add_column("Left")
        table.add_column("Right")
        table.add_column("Change", justify="center")

        for cd in other_diffs:
            for prop in cd.properties:
                prop_colour = _COLOURS[prop.change_type]
                table.add_row(
                    cd.coordinate,
                    prop.property_name,
                    str(prop.left_value) if prop.left_value is not None else "-",
                    str(prop.right_value) if prop.right_value is not None else "-",
                    Text(prop.change_type.value, style=prop_colour),
                )
        console.print(table)

    console.print()


def _add_cell_formula_row(table: Table, cd: CellDiff) -> None:
    """Add a formula diff row to a table."""
    colour = _COLOURS[cd.change_type]
    left_formula = "-"
    right_formula = "-"

    for prop in cd.properties:
        if prop.property_name == "formula":
            left_formula = str(prop.left_value) if prop.left_value else "-"
            right_formula = str(prop.right_value) if prop.right_value else "-"
            break
        if prop.property_name == "cell_kind":
            left_formula = f"[{prop.left_value}]"
            right_formula = f"[{prop.right_value}]"

    table.add_row(
        cd.coordinate,
        Text(left_formula, style="red" if cd.change_type != ChangeType.ADDED else "dim"),
        Text(right_formula, style="green" if cd.change_type != ChangeType.REMOVED else "dim"),
        Text(cd.change_type.value, style=colour),
    )
