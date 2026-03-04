"""Flask application factory for ExcelDiff web UI."""

from __future__ import annotations

import tempfile
from pathlib import Path

from flask import Flask


def create_app() -> Flask:
    """Create and configure the Flask application."""
    app = Flask(__name__)
    app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16 MB max upload
    app.config["UPLOAD_FOLDER"] = tempfile.mkdtemp(prefix="exceldiff_")
    app.secret_key = "exceldiff-local-dev"

    app.jinja_env.filters["basename"] = lambda p: Path(p).name

    from exceldiff.web.routes import bp
    app.register_blueprint(bp)

    return app
