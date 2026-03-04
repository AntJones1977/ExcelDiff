"""CLI entry point for ExcelDiff."""

from __future__ import annotations

import sys

import click

from exceldiff import __version__


@click.group()
@click.version_option(version=__version__, prog_name="exceldiff")
def cli():
    """ExcelDiff - Compare Excel workbook structure and formulae."""
    pass


@cli.command()
@click.argument("left", type=click.Path(exists=True))
@click.argument("right", type=click.Path(exists=True))
@click.option(
    "--format", "-f",
    type=click.Choice(["terminal", "html", "excel"]),
    default="terminal",
    help="Output format.",
)
@click.option(
    "--output", "-o",
    type=click.Path(),
    default=None,
    help="Output file path (for html/excel formats).",
)
@click.option(
    "--formulas-only",
    is_flag=True,
    help="Only compare formulas, skip values and other properties.",
)
@click.option(
    "--ignore-style",
    is_flag=True,
    help="Skip style/formatting comparison entirely.",
)
@click.option(
    "--sheets", "-s",
    multiple=True,
    help="Compare only specific sheets (by name). Can be repeated.",
)
@click.option(
    "--verbose", "-v",
    is_flag=True,
    help="Show all details including style diffs.",
)
def diff(left, right, format, output, formulas_only, ignore_style, sheets, verbose):
    """Compare two Excel workbooks (.xlsx files)."""
    from exceldiff.comparison.engine import compare

    sheet_filter = tuple(sheets) if sheets else None

    result = compare(
        left,
        right,
        ignore_style=ignore_style,
        formulas_only=formulas_only,
        sheet_filter=sheet_filter,
    )

    if format == "terminal":
        from exceldiff.renderers.terminal import render
        render(result, verbose=verbose)
    elif format == "html":
        click.echo("HTML renderer not yet implemented (Phase 2).")
    elif format == "excel":
        click.echo("Excel renderer not yet implemented (Phase 3).")

    sys.exit(1 if result.has_differences else 0)


@cli.command()
@click.argument("workbook", type=click.Path(exists=True))
def inspect(workbook):
    """Inspect a single workbook's structure and formulas."""
    from rich.console import Console
    from rich.panel import Panel
    from rich.table import Table

    from exceldiff.parsing.workbook_parser import parse_workbook

    console = Console()
    wb = parse_workbook(workbook)

    # Summary
    console.print(Panel(
        f"[bold]File:[/bold] {wb.file_path}\n"
        f"[bold]Sheets:[/bold] {len(wb.sheet_names)}\n"
        f"[bold]Defined Names:[/bold] {len(wb.defined_names)}",
        title="[bold]Workbook Inspector[/bold]",
        border_style="blue",
    ))

    # Per-sheet summary
    table = Table(title="Sheets", border_style="blue")
    table.add_column("Sheet Name", style="bold")
    table.add_column("Dimensions")
    table.add_column("Cells", justify="right")
    table.add_column("Formulas", justify="right")
    table.add_column("Merged Regions", justify="right")
    table.add_column("Freeze Panes")

    for name in wb.sheet_names:
        sheet = wb.sheets[name]
        formula_count = sum(
            1 for c in sheet.cells.values()
            if c.kind.value in ("formula", "array_formula")
        )
        table.add_row(
            name,
            sheet.dimensions,
            str(len(sheet.cells)),
            str(formula_count),
            str(len(sheet.merged_cells)),
            sheet.freeze_panes or "-",
        )

    console.print(table)


if __name__ == "__main__":
    cli()
