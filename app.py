"""Transaction statements web app.

Account holders log in with a 9-digit user ID (created by the admin), see their
latest transactions on a personal dashboard and download statements as PDF.
The admin portal (/admin) adds user IDs, uploads Excel files of transactions
for those IDs, and can view, add, edit or delete any transaction.
"""

import os
import secrets

from flask import Flask, flash, redirect, request, url_for

from core import check_csrf, close_db, csrf_token, format_money, get_db, init_db

import admin_views
import user_views


def create_app(config=None):
    app = Flask(__name__, instance_relative_config=True)
    app.config.update(
        SECRET_KEY=os.environ.get("SECRET_KEY") or secrets.token_hex(32),
        DATABASE=os.environ.get("DATABASE") or os.path.join(app.instance_path, "bank.db"),
        MAX_CONTENT_LENGTH=5 * 1024 * 1024,
        SESSION_COOKIE_HTTPONLY=True,
        SESSION_COOKIE_SAMESITE="Lax",
    )
    if config:
        app.config.update(config)
    os.makedirs(os.path.dirname(os.path.abspath(app.config["DATABASE"])), exist_ok=True)

    app.teardown_appcontext(close_db)
    with app.app_context():
        init_db()

    app.register_blueprint(user_views.bp)
    app.register_blueprint(admin_views.bp)
    app.jinja_env.filters["money"] = format_money

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
        }

    @app.errorhandler(413)
    def too_large(_e):
        flash("File is too large (limit 5 MB).", "error")
        return redirect(url_for("admin.upload"))

    return app


if __name__ == "__main__":
    create_app().run(debug=os.environ.get("FLASK_DEBUG") == "1", host="127.0.0.1",
                     port=int(os.environ.get("PORT", 5000)))
