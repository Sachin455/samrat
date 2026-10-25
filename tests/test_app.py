import io
import re
import sys
from datetime import date
from pathlib import Path

import pytest
from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from app import create_app  # noqa: E402
from core import DEFAULT_USER_PASSWORD, normalize_user_id, parse_amount  # noqa: E402

ADMIN_PW = "adminpass1"


@pytest.fixture
def app(tmp_path):
    return create_app({"TESTING": True, "SECRET_KEY": "test", "DATABASE": str(tmp_path / "t.db")})


@pytest.fixture
def client(app):
    return app.test_client()


def token(client):
    with client.session_transaction() as s:
        s.setdefault("csrf", "tok")
        return s["csrf"]


def post(client, path, data=None, **kw):
    return client.post(path, data={"csrf_token": token(client), **(data or {})}, **kw)


def admin_login(client):
    """Log in as the seeded admin and get past the forced password change."""
    r = post(client, "/admin/login", {"user_id": "admin", "password": "admin123"})
    if r.status_code == 401:  # default password already replaced earlier in the test
        r = post(client, "/admin/login", {"user_id": "admin", "password": ADMIN_PW})
    if r.headers["Location"].endswith("/admin/change-password"):
        post(client, "/admin/change-password",
             {"current_password": "admin123", "new_password": ADMIN_PW, "confirm": ADMIN_PW})
        return
    assert r.status_code == 302


def add_users(client, *ids):
    return post(client, "/admin/users", {"action": "add", "user_ids": "\n".join(ids)})


def user_login(client, user_id, password=DEFAULT_USER_PASSWORD):
    return post(client, "/login", {"user_id": user_id, "password": password})


PROFILE = {
    "full_name": "Sunrise Academy",
    "address": "Ward 4, Kathmandu",
    "established_date": "2005-04-14",
    "head_name": "Ram Sharma",
    "email": "office@sunrise.edu.np",
    "phone": "+977 1 4123456",
    "running_classes": "Nursery to Grade 10",
    "chairperson": "Sita Karki",
    "accountant": "Hari Thapa",
}


def onboard_user(client, user_id, password="mine123", profile=None):
    """First login: set own password, then fill in the organization profile."""
    user_login(client, user_id)
    post(client, "/change-password", {"current_password": DEFAULT_USER_PASSWORD,
                                      "new_password": password, "confirm": password})
    r = post(client, "/profile", {**PROFILE, **(profile or {})})
    assert r.headers["Location"].endswith("/dashboard")


def xlsx(rows, header=("User ID", "Date", "Description", "Type", "Amount", "Reference")):
    wb = Workbook()
    ws = wb.active
    ws.append(list(header))
    for r in rows:
        ws.append(list(r))
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def upload(client, buf, name="txns.xlsx", **fields):
    return post(client, "/admin/upload", {"file": (buf, name), **fields}, content_type="multipart/form-data")


# ----------------------------------------------------------------- admin auth


def test_admin_must_change_default_password(client):
    r = post(client, "/admin/login", {"user_id": "admin", "password": "admin123"})
    assert r.headers["Location"].endswith("/admin/change-password")
    assert client.get("/admin/users").headers["Location"].endswith("/admin/change-password")
    post(client, "/admin/change-password",
         {"current_password": "admin123", "new_password": ADMIN_PW, "confirm": ADMIN_PW})
    assert client.get("/admin/users").status_code == 200
    post(client, "/admin/logout")
    assert post(client, "/admin/login", {"user_id": "admin", "password": "admin123"}).status_code == 401
    assert post(client, "/admin/login", {"user_id": "admin", "password": ADMIN_PW}).status_code == 302


def test_admin_pages_require_admin(client):
    for path in ("/admin/", "/admin/users", "/admin/upload", "/admin/transactions", "/admin/template.xlsx"):
        assert "/admin/login" in client.get(path).headers["Location"]


def test_user_session_cannot_access_admin(client):
    admin_login(client)
    add_users(client, "100000001")
    post(client, "/admin/logout")
    user_login(client, "100000001")
    assert "/admin/login" in client.get("/admin/users").headers["Location"]


def test_post_without_csrf_rejected(client):
    assert client.post("/login", data={"user_id": "100000001", "password": "x"}).status_code == 400


# ------------------------------------------------------------------ users


def test_admin_adds_users_and_user_first_login_forces_password_change(client):
    admin_login(client)
    add_users(client, "100000001, 100000002", "12345", "100000001")
    html = client.get("/admin/users").get_data(as_text=True)
    assert "100000001" in html and "100000002" in html
    assert "Not 9-digit IDs (ignored): 12345" in html
    post(client, "/admin/logout")

    assert user_login(client, "100000001", "wrong").status_code == 401
    r = user_login(client, "100000001")
    assert r.headers["Location"].endswith("/change-password")
    assert client.get("/dashboard").headers["Location"].endswith("/change-password")

    r = post(client, "/change-password", {"current_password": DEFAULT_USER_PASSWORD,
                                          "new_password": DEFAULT_USER_PASSWORD, "confirm": DEFAULT_USER_PASSWORD})
    assert r.status_code == 400
    r = post(client, "/change-password", {"current_password": DEFAULT_USER_PASSWORD,
                                          "new_password": "mine123", "confirm": "mine123"})
    assert r.headers["Location"].endswith("/profile")
    post(client, "/profile", PROFILE)
    assert client.get("/dashboard").status_code == 200
    post(client, "/logout")
    assert user_login(client, "100000001").status_code == 401
    assert user_login(client, "100000001", "mine123").headers["Location"].endswith("/dashboard")


def test_user_can_change_password_later(client):
    admin_login(client)
    add_users(client, "100000001")
    post(client, "/admin/logout")
    user_login(client, "100000001")
    post(client, "/change-password", {"current_password": DEFAULT_USER_PASSWORD,
                                      "new_password": "first1", "confirm": "first1"})
    r = post(client, "/change-password", {"current_password": "bad", "new_password": "second2", "confirm": "second2"})
    assert b"current password is incorrect" in r.data
    post(client, "/change-password", {"current_password": "first1", "new_password": "second2", "confirm": "second2"})
    post(client, "/logout")
    assert user_login(client, "100000001", "second2").status_code == 302


def test_admin_default_password_and_reset(client):
    admin_login(client)
    post(client, "/admin/users", {"action": "default_password", "default_password": "Start#2026"})
    add_users(client, "100000001")
    post(client, "/admin/logout")
    user_login(client, "100000001", "Start#2026")
    post(client, "/change-password", {"current_password": "Start#2026", "new_password": "mine123", "confirm": "mine123"})
    post(client, "/logout")
    admin_login(client)
    post(client, "/admin/users/100000001/reset-password")
    post(client, "/admin/logout")
    assert user_login(client, "100000001", "Start#2026").headers["Location"].endswith("/change-password")


def test_register_works_without_scrypt(client, monkeypatch):
    import hashlib

    # macOS's stock Python 3.9 is built against LibreSSL and has no hashlib.scrypt.
    monkeypatch.delattr(hashlib, "scrypt", raising=False)
    admin_login(client)
    add_users(client, "100000001")
    post(client, "/admin/logout")
    assert user_login(client, "100000001").status_code == 302


# ---------------------------------------------------------------- uploads


def test_upload_routes_rows_and_dashboard_shows_latest_five(client):
    admin_login(client)
    add_users(client, "100000001", "100000002")
    rows = [("100000001", date(2026, 10, d), f"Txn day {d}", "Credit" if d % 2 else "Debit", 10 * d, f"R{d}")
            for d in range(1, 8)]
    rows += [
        (100000002, "02/10/2026", "Salary", "cr", "1,000.00", "S-1"),
        ("999999999", None, "Unknown", "Debit", 5, None),
        ("12ab", None, "Bad id", "Debit", 5, None),
        ("100000002", None, "Bad amount", "Debit", "abc", None),
    ]
    html = upload(client, xlsx(rows)).get_data(as_text=True)
    assert "Uploaded 8 transaction(s) to 2 user(s)" in html
    assert "No user with ID 999999999" in html
    assert "not a 9-digit user ID" in html and "not a valid amount" in html
    post(client, "/admin/logout")

    onboard_user(client, "100000001")
    html = client.get("/dashboard").get_data(as_text=True)
    shown = re.findall(r"Txn day (\d)", html)
    assert shown == ["7", "6", "5", "4", "3"]  # latest five, newest first
    assert "Salary" not in html
    assert "160.00" in html and "120.00" in html  # credits 10+30+50+70, debits 20+40+60
    full = client.get("/transactions").get_data(as_text=True)
    assert len(re.findall(r"Txn day \d", full)) == 7
    ranged = client.get("/transactions?from=2026-10-03&to=2026-10-04").get_data(as_text=True)
    assert re.findall(r"Txn day (\d)", ranged) == ["4", "3"]
    pdf = client.get("/transactions.pdf")
    assert pdf.mimetype == "application/pdf" and pdf.data.startswith(b"%PDF")


def test_upload_to_single_target_user_and_create_missing(client):
    admin_login(client)
    r = upload(client, xlsx([(date(2026, 1, 1), "Rent", "Debit", 500)],
                            header=("Date", "Description", "Type", "Amount")),
               target_user_id="200000001")
    assert "No user with ID 200000001" in r.get_data(as_text=True)
    r = upload(client, xlsx([(date(2026, 1, 1), "Rent", "Debit", 500)],
                            header=("Date", "Description", "Type", "Amount")),
               target_user_id="200000001", create_missing="1")
    html = r.get_data(as_text=True)
    assert "New users created:" in html and "200000001" in html
    assert "Rent" in client.get("/admin/users/200000001").get_data(as_text=True)


def test_failed_upload_does_not_create_users(client):
    admin_login(client)
    upload(client, xlsx([("300000001", None, "x", "Debit", "oops", None)]), create_missing="1")
    assert client.get("/admin/users/300000001").status_code == 404


def test_withdrawal_deposit_columns_and_leading_zero_ids(client):
    admin_login(client)
    add_users(client, "000012345")
    r = upload(client, xlsx([(12345, "Cash", 100, None), ("000012345", "Pay", None, 250)],
                            header=("User ID", "Narration", "Withdrawal", "Deposit")))
    assert "Uploaded 2 transaction(s)" in r.get_data(as_text=True)
    html = client.get("/admin/users/000012345").get_data(as_text=True)
    assert "Debit" in html and "Credit" in html and "150.00" in html  # net = 250 - 100


def test_upload_errors(client):
    admin_login(client)
    assert b"Missing required column" in upload(client, xlsx([("x",)], header=("Name",))).data
    assert b"valid .xlsx" in upload(client, io.BytesIO(b"not excel")).data
    assert upload(client, io.BytesIO(b"a,b"), name="t.csv").status_code == 400
    assert upload(client, xlsx([]), target_user_id="123").status_code == 400


def test_delete_upload(client):
    admin_login(client)
    add_users(client, "100000001")
    upload(client, xlsx([("100000001", None, "A", "Credit", 1, None)]))
    post(client, "/admin/uploads/1/delete")
    assert "A</td>" not in client.get("/admin/users/100000001").get_data(as_text=True)


# ------------------------------------------------------- admin edit access


def test_admin_add_edit_move_delete_transaction(client):
    admin_login(client)
    add_users(client, "100000001", "100000002")
    post(client, "/admin/transactions/new", {"user_id": "100000001", "txn_date": "2026-10-20",
                                             "description": "Manual", "txn_type": "Credit", "amount": "42.50",
                                             "from_user": "1"})
    html = client.get("/admin/users/100000001").get_data(as_text=True)
    assert "Manual" in html and "42.50" in html

    form = client.get("/admin/transactions/1/edit").get_data(as_text=True)
    assert 'value="42.50"' in form
    post(client, "/admin/transactions/1/edit", {"user_id": "100000002", "txn_date": "2026-10-21",
                                                "description": "Edited", "txn_type": "Debit", "amount": "40"})
    assert "Edited" not in client.get("/admin/users/100000001").get_data(as_text=True)
    assert "Edited" in client.get("/admin/users/100000002").get_data(as_text=True)
    r = post(client, "/admin/transactions/1/edit", {"user_id": "555555555", "amount": "1"})
    assert r.status_code == 400

    post(client, "/admin/transactions/1/delete")
    assert client.get("/admin/transactions/1/edit").status_code == 404


def test_admin_sees_every_users_transactions_and_pdf(client):
    admin_login(client)
    add_users(client, "100000001", "100000002")
    upload(client, xlsx([("100000001", None, "Alpha", "Credit", 1, None),
                         ("100000002", None, "Beta", "Credit", 2, None)]))
    html = client.get("/admin/transactions").get_data(as_text=True)
    assert "Alpha" in html and "Beta" in html
    html = client.get("/admin/transactions?q=100000002").get_data(as_text=True)
    assert "Beta" in html and "Alpha" not in html
    assert client.get("/admin/users/100000001/transactions.pdf").data.startswith(b"%PDF")


def test_user_only_sees_own_transactions(client):
    admin_login(client)
    add_users(client, "100000001", "100000002")
    upload(client, xlsx([("100000002", None, "Secret", "Credit", 1, None)]))
    post(client, "/admin/logout")
    onboard_user(client, "100000001")
    assert "Secret" not in client.get("/transactions").get_data(as_text=True)
    assert b"Secret" not in client.get("/transactions.pdf").data


def test_delete_user_removes_transactions(client):
    admin_login(client)
    add_users(client, "100000001")
    upload(client, xlsx([("100000001", None, "Gone", "Credit", 1, None)]))
    post(client, "/admin/users/100000001/delete")
    assert "Gone" not in client.get("/admin/transactions").get_data(as_text=True)


# ---------------------------------------------------------------- profiles


def test_profile_required_before_dashboard(client):
    admin_login(client)
    add_users(client, "100000001")
    post(client, "/admin/logout")
    user_login(client, "100000001")
    post(client, "/change-password", {"current_password": DEFAULT_USER_PASSWORD,
                                      "new_password": "mine123", "confirm": "mine123"})
    for path in ("/dashboard", "/transactions", "/transactions.pdf", "/"):
        r = client.get(path, follow_redirects=False)
        if path == "/":
            r = client.get(r.headers["Location"])
        assert r.headers["Location"].endswith("/profile"), path
    html = client.get("/profile").get_data(as_text=True)
    assert "Tell us about your organization" in html
    assert "All transactions" not in html  # nav hidden until the profile is done

    # logging out and back in still lands on the profile form
    post(client, "/logout")
    user_login(client, "100000001", "mine123")
    assert client.get("/dashboard").headers["Location"].endswith("/profile")

    post(client, "/profile", PROFILE)
    html = client.get("/dashboard").get_data(as_text=True)
    assert "Sunrise Academy" in html and "Ram Sharma" in html and "Nursery to Grade 10" in html
    assert "office@sunrise.edu.np" in html and "Hari Thapa" in html and "Sita Karki" in html


def test_profile_validation(client):
    admin_login(client)
    add_users(client, "100000001")
    post(client, "/admin/logout")
    user_login(client, "100000001")
    post(client, "/change-password", {"current_password": DEFAULT_USER_PASSWORD,
                                      "new_password": "mine123", "confirm": "mine123"})
    r = post(client, "/profile", {})
    assert r.status_code == 400
    html = r.get_data(as_text=True)
    for label in ("Full name of the organization", "Address", "Establishment date", "Head of the organization",
                  "Organization email", "Organization contact number", "Running classes",
                  "Chairperson of the organization", "Accountant of the organization"):
        assert f"{label} is required." in html

    bad = {**PROFILE, "email": "not-an-email", "phone": "12ab", "established_date": "2999-01-01"}
    html = post(client, "/profile", bad).get_data(as_text=True)
    assert "valid email" in html and "valid phone number" in html and "can&#39;t be in the future" in html
    assert 'value="Sunrise Academy"' in html  # entered values are kept
    assert client.get("/dashboard").headers["Location"].endswith("/profile")


def test_user_can_edit_profile_later(client):
    admin_login(client)
    add_users(client, "100000001")
    post(client, "/admin/logout")
    onboard_user(client, "100000001")
    assert 'value="Sunrise Academy"' in client.get("/profile").get_data(as_text=True)
    post(client, "/profile", {**PROFILE, "accountant": "Gita Rai"})
    html = client.get("/dashboard").get_data(as_text=True)
    assert "Gita Rai" in html and "Hari Thapa" not in html


def test_admin_sees_and_edits_profiles(client):
    admin_login(client)
    add_users(client, "100000001", "100000002")
    post(client, "/admin/logout")
    onboard_user(client, "100000001")
    post(client, "/logout")

    admin_login(client)
    html = client.get("/admin/users").get_data(as_text=True)
    assert "Sunrise Academy" in html and "First-time password" in html  # 100000002 never logged in
    html = client.get("/admin/users?q=sunrise").get_data(as_text=True)
    assert "/admin/users/100000001" in html and "/admin/users/100000002" not in html
    html = client.get("/admin/users/100000001").get_data(as_text=True)
    assert "Sunrise Academy" in html and "Ram Sharma" in html
    assert "hasn't filled in their profile" in client.get("/admin/users/100000002").get_data(as_text=True)

    r = post(client, "/admin/users/100000001/profile", {**PROFILE, "head_name": "New Head"})
    assert r.headers["Location"].endswith("/admin/users/100000001")
    assert "New Head" in client.get("/admin/users/100000001").get_data(as_text=True)
    assert post(client, "/admin/users/100000001/profile", {**PROFILE, "email": "x"}).status_code == 400


def test_personal_profile_optional_and_editable(client):
    admin_login(client)
    add_users(client, "100000001")
    post(client, "/admin/logout")
    onboard_user(client, "100000001")

    # not required to reach the dashboard, but suggested there
    html = client.get("/dashboard").get_data(as_text=True)
    assert "My profile" in html and "Add your own name" in html

    r = post(client, "/me", {"designation": "Principal"})
    assert r.status_code == 400 and "Full name is required." in r.get_data(as_text=True)
    r = post(client, "/me", {"full_name": "Ram Sharma", "email": "nope", "phone": "x"})
    html = r.get_data(as_text=True)
    assert "valid email" in html and "valid phone number" in html

    r = post(client, "/me", {"full_name": "Ram Sharma", "designation": "Principal",
                             "email": "ram@example.com", "phone": "9812345678"})
    assert r.headers["Location"].endswith("/dashboard")
    html = client.get("/dashboard").get_data(as_text=True)
    assert "Signed in as Ram Sharma (Principal)" in html and "Add your own name" not in html
    assert "Log out (Ram Sharma)" in html
    assert 'value="ram@example.com"' in client.get("/me").get_data(as_text=True)

    post(client, "/me", {"full_name": "Ram K. Sharma"})  # optional fields can be cleared
    html = client.get("/me").get_data(as_text=True)
    assert 'value="Ram K. Sharma"' in html and "ram@example.com" not in html

    post(client, "/logout")
    admin_login(client)
    html = client.get("/admin/users/100000001").get_data(as_text=True)
    assert "Personal profile" in html and "Ram K. Sharma" in html


def test_personal_profile_requires_onboarding_first(client):
    admin_login(client)
    add_users(client, "100000001")
    post(client, "/admin/logout")
    user_login(client, "100000001")
    assert client.get("/me").headers["Location"].endswith("/change-password")
    post(client, "/change-password", {"current_password": DEFAULT_USER_PASSWORD,
                                      "new_password": "mine123", "confirm": "mine123"})
    assert client.get("/me").headers["Location"].endswith("/profile")


# ------------------------------------------------------------------ units


@pytest.mark.parametrize("raw,cents", [
    (10, 1000), (10.5, 1050), ("1,234.56", 123456), ("$99.99", 9999), ("(5.00)", -500), ("-3", -300),
])
def test_parse_amount(raw, cents):
    assert parse_amount(raw) == cents


@pytest.mark.parametrize("raw,expected", [
    ("123456789", "123456789"), (123456789, "123456789"), (12345, "000012345"), ("123 456 789", "123456789"),
    ("12345", None), ("1234567890", None), ("abc", None), (None, None), (True, None),
])
def test_normalize_user_id(raw, expected):
    assert normalize_user_id(raw) == expected
