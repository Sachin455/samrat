"""Pages for account holders: log in, change password, dashboard, transactions, PDF."""

from datetime import datetime
from functools import wraps

from flask import (
    Blueprint,
    abort,
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
    PERSONAL_FIELDS,
    PROFILE_FIELDS,
    PROFILE_TABLES,
    TXN_ORDER,
    TXN_SELECT,
    statement_pdf,
    get_db,
    get_profile,
    hash_password,
    normalize_user_id,
    now_iso,
    safe_next,
    save_profile,
    send_document,
    txn_summary,
    validate_new_password,
    validate_profile,
)

bp = Blueprint("user", __name__)


def current_user():
    uid = session.get("uid")
    if uid is None:
        return None
    if "user" not in g:
        g.user = get_db().execute(
            """SELECT u.id, u.user_id, u.must_change_password, p.full_name, pp.full_name AS person_name
               FROM users u
               LEFT JOIN profiles p ON p.user_pk = u.id
               LEFT JOIN personal_profiles pp ON pp.user_pk = u.id
               WHERE u.id = ?""",
            (uid,),
        ).fetchone()
    return g.user


def login_required(view):
    """Require a logged-in user who has set a password and filled in their profile."""

    @wraps(view)
    def wrapped(*args, **kwargs):
        user = current_user()
        if user is None:
            session.pop("uid", None)
            return redirect(url_for("user.login", next=request.path))
        if user["must_change_password"] and request.endpoint != "user.change_password":
            return redirect(url_for("user.change_password"))
        if (not user["must_change_password"] and user["full_name"] is None
                and request.endpoint not in ("user.profile", "user.change_password")):
            return redirect(url_for("user.profile"))
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
        if user["full_name"] is None:
            flash("Password saved. Now please fill in your organization's profile.", "success")
            return redirect(url_for("user.profile"))
        flash("Your password has been changed.", "success")
        return redirect(url_for("user.dashboard"))
    return render_template("change_password.html", portal="user", forced=user["must_change_password"])


def profile_form(kind, heading, intro="", first_time=False):
    """Show and save one of the current user's profile forms.

    Returns a response to send, or None once the form has been saved.
    """
    user = current_user()
    page = {"portal": "user", "fields": PROFILE_TABLES[kind][1], "heading": heading, "intro": intro,
            "first_time": first_time, "cancel_url": url_for("user.dashboard")}
    if request.method == "POST":
        values, errors = validate_profile(request.form, kind)
        if errors:
            flash("Please correct the highlighted fields.", "error")
            return render_template("profile_form.html", values=values, errors=errors, **page), 400
        save_profile(user["id"], values, kind)
        get_db().commit()
        return None
    existing = get_profile(user["id"], kind)
    return render_template("profile_form.html", values=dict(existing) if existing else {}, errors={}, **page)


@bp.route("/profile", methods=["GET", "POST"])
@login_required
def profile():
    """Organization profile: required on first login, editable afterwards."""
    first_time = current_user()["full_name"] is None
    if first_time:
        response = profile_form(
            "organization", "Tell us about your organization",
            "All fields are required. Your dashboard opens once this is saved, "
            "and you can edit these details any time later.", first_time=True)
    else:
        response = profile_form("organization", "Organization profile")
    if response is not None:
        return response
    flash("Profile saved. Welcome to your dashboard!" if first_time else "Organization profile updated.", "success")
    return redirect(url_for("user.dashboard"))


@bp.route("/me", methods=["GET", "POST"])
@login_required
def personal_profile():
    """The account holder's own details, editable any time."""
    response = profile_form("personal", "My profile", "Your own details as the person using this account.")
    if response is not None:
        return response
    flash("Your profile has been updated.", "success")
    return redirect(url_for("user.dashboard"))


@bp.route("/dashboard")
@login_required
def dashboard():
    user = current_user()
    latest = get_db().execute(
        f"{TXN_SELECT} WHERE t.user_pk = ? {TXN_ORDER} LIMIT ?", (user["id"], LATEST_COUNT)
    ).fetchall()
    return render_template(
        "user/dashboard.html", latest=latest, summary=txn_summary("t.user_pk = ?", (user["id"],)),
        latest_count=LATEST_COUNT, profile=get_profile(user["id"]), fields=PROFILE_FIELDS,
        personal=get_profile(user["id"], "personal"), personal_fields=PERSONAL_FIELDS,
        documents=get_db().execute(
            "SELECT * FROM documents WHERE user_pk = ? ORDER BY id DESC LIMIT ?", (user["id"], LATEST_COUNT)
        ).fetchall(),
        document_count=get_db().execute(
            "SELECT COUNT(*) FROM documents WHERE user_pk = ?", (user["id"],)).fetchone()[0],
    )


@bp.route("/documents")
@login_required
def documents():
    rows = get_db().execute(
        "SELECT * FROM documents WHERE user_pk = ? ORDER BY id DESC", (current_user()["id"],)).fetchall()
    return render_template("user/documents.html", documents=rows)


def own_document_or_404(doc_id):
    doc = get_db().execute(
        "SELECT * FROM documents WHERE id = ? AND user_pk = ?", (doc_id, current_user()["id"])).fetchone()
    if doc is None:
        abort(404)
    return doc


@bp.route("/documents/<int:doc_id>")
@login_required
def download_document(doc_id):
    return send_document(own_document_or_404(doc_id))


@bp.route("/documents/<int:doc_id>/view")
@login_required
def view_document(doc_id):
    return send_document(own_document_or_404(doc_id), inline=True)


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
    pdf = statement_pdf(user["id"], user["user_id"], filters["from"], filters["to"])
    return send_file(pdf, as_attachment=True, download_name=f"statement_{user['user_id']}.pdf",
                     mimetype="application/pdf")
