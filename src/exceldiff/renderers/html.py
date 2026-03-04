"""HTML report renderer using Jinja2 templates."""

from __future__ import annotations

from pathlib import Path

from jinja2 import Environment, FileSystemLoader

from exceldiff import __version__
from exceldiff.comparison.models import WorkbookDiff

_TEMPLATES_DIR = Path(__file__).parent / "templates"


def render(diff: WorkbookDiff, output_path: str = "exceldiff_report.html") -> str:
    """Render a WorkbookDiff to a standalone HTML report file.

    Args:
        diff: The comparison result to render.
        output_path: Path for the output HTML file.

    Returns:
        The absolute path to the generated report.
    """
    env = Environment(
        loader=FileSystemLoader(str(_TEMPLATES_DIR)),
        autoescape=True,
    )
    env.filters["basename"] = lambda p: Path(p).name

    # Load CSS inline for a self-contained report
    css_path = _TEMPLATES_DIR / "style.css"
    css = css_path.read_text(encoding="utf-8")

    template = env.get_template("report.html")
    html = template.render(
        diff=diff,
        css=css,
        version=__version__,
    )

    out = Path(output_path)
    out.write_text(html, encoding="utf-8")
    return str(out.resolve())
