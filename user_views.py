"""Pages for account holders: log in, change password, dashboard, transactions, PDF."""

from datetime import datetime
from functools import wraps

from flask import (
    Blueprint,
    flash,
    g,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
from werkzeug.security import check_password_hash

from core import (
    LATEST_COUNT,
    TXN_ORDER,
    TXN_SELECT,
    build_pdf,
    get_db,
    hash_password,
    normalize_user_id,
    now_iso,
    safe_next,
    txn_summary,
    validate_new_password,
)

bp = Blueprint("user", __name__)


def current_user():
    uid = session.get("uid")
    if uid is None:
        return None
    if "user" not in g:
        g.user = get_db().execute(
            "SELECT id, user_id, must_change_password FROM users WHERE id = ?", (uid,)
        ).fetchone()
    return g.user


def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        user = current_user()
        if user is None:
            session.pop("uid", None)
            return redirect(url_for("user.login", next=request.path))
        if user["must_change_password"] and request.endpoint != "user.change_password":
            return redirect(url_for("user.change_password"))
        return view(*args, **kwargs)

    return wrapped


def date_filters():
    """Read optional ?from=YYYY-MM-DD&to=YYYY-MM-DD filters."""
    filters = {}
    for key in ("from", "to"):
        value = request.args.get(key, "").strip()
        try:
            filters[key] = datetime.strptime(value, "%Y-%m-%d").date().isoformat() if value else ""
        except ValueError:
            filters[key] = ""
    return filters


def user_txn_query(user_pk, filters):
    where, params = "t.user_pk = ?", [user_pk]
    if filters.get("from"):
        where += " AND t.txn_date >= ?"
        params.append(filters["from"])
    if filters.get("to"):
        where += " AND t.txn_date <= ?"
        params.append(filters["to"])
    return where, params


@bp.route("/")
def index():
    return redirect(url_for("user.dashboard" if current_user() else "user.login"))


@bp.route("/login", methods=["GET", "POST"])
def login():
    if current_user():
        return redirect(url_for("user.dashboard"))
    if request.method == "POST":
        raw_id = request.form.get("user_id", "").strip()
        password = request.form.get("password", "")
        user_id = normalize_user_id(raw_id)
        user = None
        if user_id:
            user = get_db().execute(
                "SELECT id, password_hash, must_change_password FROM users WHERE user_id = ?", (user_id,)
            ).fetchone()
        if user and check_password_hash(user["password_hash"], password):
            db = get_db()
            db.execute("UPDATE users SET last_login = ? WHERE id = ?", (now_iso(), user["id"]))
            db.commit()
            session["uid"] = user["id"]
            if user["must_change_password"]:
                flash("Welcome! Please set your own password before continuing.", "success")
                return redirect(url_for("user.change_password"))
            return redirect(safe_next(request.args.get("next"), url_for("user.dashboard")))
        flash("Invalid user ID or password.", "error")
        return render_template("login.html", portal="user", user_id=raw_id), 401
    return render_template("login.html", portal="user", user_id="")


@bp.route("/logout", methods=["POST"])
def logout():
    session.pop("uid", None)
    flash("You have been logged out.", "success")
    return redirect(url_for("user.login"))


@bp.route("/change-password", methods=["GET", "POST"])
@login_required
def change_password():
    user = current_user()
    if request.method == "POST":
        db = get_db()
        stored = db.execute("SELECT password_hash FROM users WHERE id = ?", (user["id"],)).fetchone()
        current = request.form.get("current_password", "")
        new = request.form.get("new_password", "")
        error = validate_new_password(
            check_password_hash(stored["password_hash"], current), new, request.form.get("confirm", ""), current
        )
        if error:
            flash(error, "error")
            return render_template("change_password.html", portal="user", forced=user["must_change_password"]), 400
        db.execute(
            "UPDATE users SET password_hash = ?, must_change_password = 0 WHERE id = ?",
            (hash_password(new), user["id"]),
        )
        db.commit()
        flash("Your password has been changed.", "success")
        return redirect(url_for("user.dashboard"))
    return render_template("change_password.html", portal="user", forced=user["must_change_password"])


@bp.route("/dashboard")
@login_required
def dashboard():
    user = current_user()
    latest = get_db().execute(
        f"{TXN_SELECT} WHERE t.user_pk = ? {TXN_ORDER} LIMIT ?", (user["id"], LATEST_COUNT)
    ).fetchall()
    return render_template(
        "user/dashboard.html", latest=latest, summary=txn_summary("t.user_pk = ?", (user["id"],)),
        latest_count=LATEST_COUNT,
    )


@bp.route("/transactions")
@login_required
def transactions():
    user = current_user()
    filters = date_filters()
    where, params = user_txn_query(user["id"], filters)
    rows = get_db().execute(f"{TXN_SELECT} WHERE {where} {TXN_ORDER}", params).fetchall()
    return render_template(
        "user/transactions.html", rows=rows, filters=filters, summary=txn_summary(where, params)
    )


@bp.route("/transactions.pdf")
@login_required
def transactions_pdf():
    user = current_user()
    filters = date_filters()
    where, params = user_txn_query(user["id"], filters)
    rows = get_db().execute(f"{TXN_SELECT} WHERE {where} {TXN_ORDER}", params).fetchall()
    period = ""
    if filters["from"] or filters["to"]:
        period = f" | Period: {filters['from'] or 'start'} to {filters['to'] or 'today'}"
    subtitle = f"User ID: {user['user_id']}{period} | Generated: {datetime.now():%Y-%m-%d %H:%M}"
    pdf = build_pdf("Transaction Statement", subtitle, rows)
    return send_file(pdf, as_attachment=True, download_name=f"statement_{user['user_id']}.pdf",
                     mimetype="application/pdf")
