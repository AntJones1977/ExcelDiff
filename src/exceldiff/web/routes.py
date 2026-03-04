"""Flask routes for ExcelDiff web UI."""

from __future__ import annotations

import os
import uuid
from pathlib import Path

from flask import (
    Blueprint,
    current_app,
    flash,
    redirect,
    render_template,
    request,
    send_file,
    url_for,
)

from exceldiff.comparison.engine import compare
from exceldiff.comparison.models import WorkbookDiff

bp = Blueprint("main", __name__)

# In-memory store for comparison results (keyed by job_id)
_results: dict[str, WorkbookDiff] = {}
_file_paths: dict[str, tuple[str, str]] = {}


@bp.route("/")
def index():
    """Upload form."""
    return render_template("upload.html")


@bp.route("/compare", methods=["POST"])
def compare_files():
    """Handle file uploads and run comparison."""
    left_file = request.files.get("left")
    right_file = request.files.get("right")

    if not left_file or not left_file.filename:
        flash("Please select a left (baseline) file.", "error")
        return redirect(url_for("main.index"))
    if not right_file or not right_file.filename:
        flash("Please select a right (comparison) file.", "error")
        return redirect(url_for("main.index"))

    if not left_file.filename.endswith(".xlsx") or not right_file.filename.endswith(".xlsx"):
        flash("Only .xlsx files are supported.", "error")
        return redirect(url_for("main.index"))

    # Save uploaded files
    upload_dir = current_app.config["UPLOAD_FOLDER"]
    job_id = str(uuid.uuid4())[:8]
    job_dir = os.path.join(upload_dir, job_id)
    os.makedirs(job_dir, exist_ok=True)

    left_path = os.path.join(job_dir, left_file.filename)
    right_path = os.path.join(job_dir, right_file.filename)
    left_file.save(left_path)
    right_file.save(right_path)

    # Run comparison
    ignore_style = "ignore_style" in request.form
    formulas_only = "formulas_only" in request.form

    try:
        result = compare(
            left_path,
            right_path,
            ignore_style=ignore_style,
            formulas_only=formulas_only,
        )
    except Exception as e:
        flash(f"Comparison failed: {e}", "error")
        return redirect(url_for("main.index"))

    _results[job_id] = result
    _file_paths[job_id] = (left_path, right_path)

    return redirect(url_for("main.results", job_id=job_id))


@bp.route("/results/<job_id>")
def results(job_id: str):
    """Display comparison results."""
    diff = _results.get(job_id)
    if diff is None:
        flash("Results not found. Please run a new comparison.", "error")
        return redirect(url_for("main.index"))
    return render_template("results.html", diff=diff, job_id=job_id)


@bp.route("/results/<job_id>/sheet/<sheet_name>")
def sheet_detail(job_id: str, sheet_name: str):
    """Display detailed diff for a single sheet."""
    diff = _results.get(job_id)
    if diff is None:
        flash("Results not found.", "error")
        return redirect(url_for("main.index"))

    sheet_diff = next(
        (sd for sd in diff.sheet_diffs if sd.sheet_name == sheet_name), None
    )
    if sheet_diff is None:
        flash(f"Sheet '{sheet_name}' not found.", "error")
        return redirect(url_for("main.results", job_id=job_id))

    return render_template(
        "sheet_detail.html",
        diff=diff,
        sheet_diff=sheet_diff,
        job_id=job_id,
    )


@bp.route("/results/<job_id>/download/<fmt>")
def download(job_id: str, fmt: str):
    """Download results as HTML or annotated Excel."""
    diff = _results.get(job_id)
    if diff is None:
        flash("Results not found.", "error")
        return redirect(url_for("main.index"))

    upload_dir = current_app.config["UPLOAD_FOLDER"]
    job_dir = os.path.join(upload_dir, job_id)

    if fmt == "html":
        from exceldiff.renderers.html import render as render_html
        out_path = os.path.join(job_dir, "exceldiff_report.html")
        render_html(diff, out_path)
        return send_file(out_path, as_attachment=True, download_name="exceldiff_report.html")

    elif fmt == "excel":
        from exceldiff.renderers.excel import render as render_excel
        out_path = os.path.join(job_dir, "exceldiff_annotated.xlsx")
        render_excel(diff, out_path)
        return send_file(out_path, as_attachment=True, download_name="exceldiff_annotated.xlsx")

    flash("Invalid format.", "error")
    return redirect(url_for("main.results", job_id=job_id))
