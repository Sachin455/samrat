"""Transaction statements web app.

Account holders log in with a 9-digit user ID (created by the admin), see their
latest transactions on a personal dashboard and download statements as PDF.
The admin portal (/admin) adds user IDs, uploads Excel files of transactions
for those IDs, and can view, add, edit or delete any transaction.
"""

import os
import secrets
from datetime import date

from flask import Flask, flash, redirect, request, url_for

from core import (
    DOCUMENT_TYPES,
    INLINE_TYPES,
    check_csrf,
    close_db,
    csrf_token,
    document_ext,
    format_money,
    format_size,
    init_db,
)

import admin_views
import user_views

MAX_UPLOAD_MB = 25


def create_app(config=None):
    app = Flask(__name__, instance_relative_config=True)
    app.config.update(
        SECRET_KEY=os.environ.get("SECRET_KEY") or secrets.token_hex(32),
        DATABASE=os.environ.get("DATABASE") or os.path.join(app.instance_path, "bank.db"),
        MAX_CONTENT_LENGTH=MAX_UPLOAD_MB * 1024 * 1024,
        SESSION_COOKIE_HTTPONLY=True,
        SESSION_COOKIE_SAMESITE="Lax",
    )
    if config:
        app.config.update(config)
    app.config.setdefault("DOCUMENTS_DIR", os.environ.get("DOCUMENTS_DIR")
                          or os.path.join(os.path.dirname(os.path.abspath(app.config["DATABASE"])), "documents"))
    os.makedirs(os.path.dirname(os.path.abspath(app.config["DATABASE"])), exist_ok=True)

    app.teardown_appcontext(close_db)
    with app.app_context():
        init_db()

    app.register_blueprint(user_views.bp)
    app.register_blueprint(admin_views.bp)
    app.jinja_env.filters["money"] = format_money
    app.jinja_env.filters["filesize"] = format_size
    app.jinja_env.globals["can_preview"] = lambda name: document_ext(name) in INLINE_TYPES
    app.jinja_env.globals["document_accept"] = ",".join(sorted(DOCUMENT_TYPES))

    @app.before_request
    def protect_forms():
        if request.method == "POST":
            check_csrf()

    @app.context_processor
    def inject():
        return {
            "current_user": user_views.current_user(),
            "current_admin": admin_views.current_admin(),
            "csrf_token": csrf_token,
            "today": date.today().isoformat(),
        }

    @app.errorhandler(413)
    def too_large(_e):
        flash(f"Upload is too large (limit {MAX_UPLOAD_MB} MB in total).", "error")
        return redirect(url_for("admin.upload"))

    return app


if __name__ == "__main__":
    create_app().run(debug=os.environ.get("FLASK_DEBUG") == "1", host="127.0.0.1",
                     port=int(os.environ.get("PORT", 5000)))
