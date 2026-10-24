import io
import re
import sys
from datetime import date
from pathlib import Path

import pytest
from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from app import create_app, parse_amount  # noqa: E402


@pytest.fixture
def app(tmp_path):
    return create_app({"TESTING": True, "SECRET_KEY": "test", "DATABASE": str(tmp_path / "t.db")})


@pytest.fixture
def client(app):
    return app.test_client()


def token(client, path="/login"):
    html = client.get(path).get_data(as_text=True)
    return re.search(r'name="csrf_token" value="([^"]+)"', html).group(1)


def register(client, user_id, password="secret123"):
    return client.post("/register", data={
        "csrf_token": token(client, "/register"), "user_id": user_id,
        "password": password, "confirm": password,
    })


def logout(client):
    client.post("/logout", data={"csrf_token": token(client, "/dashboard")})


def login(client, user_id, password="secret123"):
    return client.post("/login", data={
        "csrf_token": token(client), "user_id": user_id, "password": password,
    })


def xlsx(rows, header=("Recipient", "Date", "Description", "Type", "Amount", "Reference")):
    wb = Workbook()
    ws = wb.active
    ws.append(list(header))
    for r in rows:
        ws.append(list(r))
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def upload(client, buf, name="txns.xlsx"):
    return client.post("/upload", data={
        "csrf_token": token(client, "/upload"), "file": (buf, name),
    }, content_type="multipart/form-data")


def test_register_requires_unique_id(client):
    assert register(client, "alice").status_code == 302
    logout(client)
    resp = register(client, "ALICE")
    assert resp.status_code == 400
    assert b"already taken" in resp.data


def test_register_works_without_scrypt(client, monkeypatch):
    import hashlib

    # macOS's stock Python 3.9 is built against LibreSSL and has no hashlib.scrypt.
    monkeypatch.delattr(hashlib, "scrypt", raising=False)
    assert register(client, "alice").status_code == 302
    logout(client)
    assert login(client, "alice").status_code == 302


def test_login_and_bad_password(client):
    register(client, "alice")
    logout(client)
    assert login(client, "alice", "wrong").status_code == 401
    assert login(client, "alice").status_code == 302
    assert client.get("/dashboard").status_code == 200


def test_pages_require_login(client):
    for path in ("/dashboard", "/upload", "/transactions.pdf", "/template.xlsx"):
        assert client.get(path).status_code == 302


def test_post_without_csrf_rejected(client):
    resp = client.post("/register", data={"user_id": "x1x", "password": "secret123", "confirm": "secret123"})
    assert resp.status_code == 400


def test_upload_routes_rows_to_recipients_and_pdf(client):
    for u in ("alice", "bob", "carol"):
        register(client, u)
        logout(client)

    login(client, "carol")
    resp = upload(client, xlsx([
        ("alice", date(2026, 10, 1), "Rent", "Debit", 1200.5, "R-1"),
        ("Bob", "02/10/2026", "Salary", "Credit", "1,000.00", "S-1"),
        ("alice", None, "Coffee", "Debit", 4, None),
        ("nobody", None, "Lost", "Debit", 5, None),
        ("bob", None, "Bad", "Debit", "abc", None),
    ]))
    html = resp.get_data(as_text=True)
    assert resp.status_code == 200
    assert "Uploaded 3 transaction(s) to 2 user(s)" in html
    assert "No user with ID &#39;nobody&#39;" in html
    assert "not a valid amount" in html
    logout(client)

    login(client, "alice")
    html = client.get("/dashboard").get_data(as_text=True)
    assert "Rent" in html and "Coffee" in html and "Salary" not in html
    assert "1,204.50" in html
    pdf = client.get("/transactions.pdf")
    assert pdf.status_code == 200 and pdf.mimetype == "application/pdf"
    assert pdf.data.startswith(b"%PDF")
    assert client.get("/transactions.pdf?sender=carol").data.startswith(b"%PDF")
    # alice cannot see carol's upload
    assert client.get("/uploads/1").status_code == 404
    logout(client)

    login(client, "bob")
    html = client.get("/dashboard").get_data(as_text=True)
    assert "Salary" in html and "2026-10-02" in html and "Rent" not in html
    logout(client)

    login(client, "carol")
    assert "Rent" in client.get("/uploads/1").get_data(as_text=True)
    assert client.get("/uploads/1.pdf").data.startswith(b"%PDF")


def test_upload_missing_columns(client):
    register(client, "alice")
    resp = upload(client, xlsx([("x",)], header=("Name",)))
    assert resp.status_code == 400
    assert b"Missing required column" in resp.data


def test_upload_rejects_non_excel(client):
    register(client, "alice")
    resp = upload(client, io.BytesIO(b"not excel"), name="t.xlsx")
    assert resp.status_code == 400
    assert b"valid .xlsx" in resp.data
    assert upload(client, io.BytesIO(b"a,b"), name="t.csv").status_code == 400


def test_template_download(client):
    register(client, "alice")
    resp = client.get("/template.xlsx")
    assert resp.status_code == 200 and resp.data[:2] == b"PK"


@pytest.mark.parametrize("raw,cents", [
    (10, 1000), (10.5, 1050), ("1,234.56", 123456), ("$99.99", 9999), ("(5.00)", -500), ("-3", -300),
])
def test_parse_amount(raw, cents):
    assert parse_amount(raw) == cents
