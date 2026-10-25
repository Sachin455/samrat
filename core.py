"""Shared pieces: database, security helpers, Excel parsing and PDF output."""

import csv
import io
import mimetypes
import os
import re
import secrets
import sqlite3
from datetime import date, datetime
from decimal import Decimal, InvalidOperation

from flask import abort, current_app, g, request, send_file, session
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Table, TableStyle
from werkzeug.security import generate_password_hash

USER_ID_RE = re.compile(r"^\d{9}$")
MIN_PASSWORD_LEN = 6
# Werkzeug defaults to scrypt, which Python builds linked against LibreSSL
# (e.g. macOS's stock Python 3.9) lack; PBKDF2 is available everywhere.
PASSWORD_HASH_METHOD = "pbkdf2:sha256"
MAX_ROWS = 5000
LATEST_COUNT = 5

DEFAULT_USER_PASSWORD = "Welcome@123"
DEFAULT_ADMIN_USERNAME = "admin"
DEFAULT_ADMIN_PASSWORD = "admin123"

SCHEMA = """
CREATE TABLE IF NOT EXISTS settings (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS admins (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT NOT NULL UNIQUE COLLATE NOCASE,
    password_hash TEXT NOT NULL,
    must_change_password INTEGER NOT NULL DEFAULT 1
);
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id TEXT NOT NULL UNIQUE,
    password_hash TEXT NOT NULL,
    must_change_password INTEGER NOT NULL DEFAULT 1,
    created_at TEXT NOT NULL,
    last_login TEXT
);
CREATE TABLE IF NOT EXISTS uploads (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    filename TEXT NOT NULL,
    target_user_id TEXT,
    row_count INTEGER NOT NULL,
    created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS transactions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_pk INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    upload_id INTEGER REFERENCES uploads(id) ON DELETE SET NULL,
    txn_date TEXT,
    description TEXT NOT NULL DEFAULT '',
    txn_type TEXT NOT NULL DEFAULT '',
    amount_cents INTEGER NOT NULL,
    reference TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL,
    updated_at TEXT
);
CREATE TABLE IF NOT EXISTS profiles (
    user_pk INTEGER PRIMARY KEY REFERENCES users(id) ON DELETE CASCADE,
    full_name TEXT NOT NULL,
    address TEXT NOT NULL,
    established_date TEXT NOT NULL,
    head_name TEXT NOT NULL,
    email TEXT NOT NULL,
    phone TEXT NOT NULL,
    running_classes TEXT NOT NULL,
    chairperson TEXT NOT NULL,
    accountant TEXT NOT NULL,
    updated_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS personal_profiles (
    user_pk INTEGER PRIMARY KEY REFERENCES users(id) ON DELETE CASCADE,
    full_name TEXT NOT NULL,
    designation TEXT NOT NULL DEFAULT '',
    email TEXT NOT NULL DEFAULT '',
    phone TEXT NOT NULL DEFAULT '',
    address TEXT NOT NULL DEFAULT '',
    updated_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS documents (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_pk INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    title TEXT NOT NULL,
    original_name TEXT NOT NULL,
    stored_name TEXT NOT NULL UNIQUE,
    size INTEGER NOT NULL,
    created_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_doc_user ON documents(user_pk, created_at);
CREATE INDEX IF NOT EXISTS idx_txn_user ON transactions(user_pk, txn_date);
CREATE INDEX IF NOT EXISTS idx_txn_upload ON transactions(upload_id);
"""

# Newest first; rows without a date go last.
TXN_ORDER = "ORDER BY t.txn_date IS NULL, t.txn_date DESC, t.id DESC"
TXN_ORDER_OLDEST_FIRST = "ORDER BY t.txn_date IS NULL, t.txn_date, t.id"  # for statements
DEBIT_SQL = "lower(t.txn_type) LIKE 'd%' OR lower(t.txn_type) LIKE 'withdraw%'"
TXN_SELECT = "SELECT t.*, u.user_id FROM transactions t JOIN users u ON u.id = t.user_pk"


# --------------------------------------------------------------------------- db


def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(current_app.config["DATABASE"])
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys = ON")
    return g.db


def close_db(_exc=None):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def init_db():
    db = get_db()
    db.executescript(SCHEMA)
    if db.execute("SELECT 1 FROM admins").fetchone() is None:
        db.execute(
            "INSERT INTO admins (username, password_hash) VALUES (?, ?)",
            (
                os.environ.get("ADMIN_USERNAME") or DEFAULT_ADMIN_USERNAME,
                hash_password(os.environ.get("ADMIN_PASSWORD") or DEFAULT_ADMIN_PASSWORD),
            ),
        )
    db.execute(
        "INSERT OR IGNORE INTO settings (key, value) VALUES ('default_password', ?)",
        (os.environ.get("DEFAULT_USER_PASSWORD") or DEFAULT_USER_PASSWORD,),
    )
    db.commit()


def get_setting(key):
    row = get_db().execute("SELECT value FROM settings WHERE key = ?", (key,)).fetchone()
    return row["value"] if row else None


def set_setting(key, value):
    get_db().execute(
        "INSERT INTO settings (key, value) VALUES (?, ?) "
        "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
        (key, value),
    )


def now_iso():
    return datetime.now().isoformat(timespec="seconds")


def hash_password(password):
    return generate_password_hash(password, method=PASSWORD_HASH_METHOD)


def create_users(user_ids):
    """Create accounts with the current default password. Returns (created, existing)."""
    db = get_db()
    password_hash = hash_password(get_setting("default_password"))
    created, existing = [], []
    for uid in dict.fromkeys(user_ids):  # de-duplicate, keep order
        try:
            db.execute(
                "INSERT INTO users (user_id, password_hash, created_at) VALUES (?, ?, ?)",
                (uid, password_hash, now_iso()),
            )
            created.append(uid)
        except sqlite3.IntegrityError:
            existing.append(uid)
    return created, existing


# (column, label, input type, max length, required) for the two profile forms.
PROFILE_FIELDS = [  # organization profile, filled in on first login
    ("full_name", "Full name of the organization", "text", 200, True),
    ("address", "Address", "textarea", 500, True),
    ("established_date", "Establishment date", "date", 10, True),
    ("head_name", "Head of the organization", "text", 120, True),
    ("email", "Organization email", "email", 120, True),
    ("phone", "Organization contact number", "tel", 30, True),
    ("running_classes", "Running classes", "textarea", 500, True),
    ("chairperson", "Chairperson of the organization", "text", 120, True),
    ("accountant", "Accountant of the organization", "text", 120, True),
]
PERSONAL_FIELDS = [  # the account holder's own details, editable any time
    ("full_name", "Full name", "text", 120, True),
    ("designation", "Designation / position", "text", 120, False),
    ("email", "Personal email", "email", 120, False),
    ("phone", "Mobile number", "tel", 30, False),
    ("address", "Personal address", "textarea", 500, False),
]
PROFILE_TABLES = {"organization": ("profiles", PROFILE_FIELDS), "personal": ("personal_profiles", PERSONAL_FIELDS)}
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
PHONE_RE = re.compile(r"^\+?[\d\s()-]+$")


def get_profile(user_pk, kind="organization"):
    table, _ = PROFILE_TABLES[kind]
    return get_db().execute(f"SELECT * FROM {table} WHERE user_pk = ?", (user_pk,)).fetchone()


def validate_profile(form, kind="organization"):
    """Return (values, errors) where errors maps field name -> message."""
    _, fields = PROFILE_TABLES[kind]
    values, errors = {}, {}
    for key, label, input_type, max_len, required in fields:
        value = form.get(key, "").strip()
        values[key] = value
        if not value:
            if required:
                errors[key] = f"{label} is required."
            continue
        if len(value) > max_len:
            errors[key] = f"{label} must be at most {max_len} characters."
        elif input_type == "email" and not EMAIL_RE.match(value):
            errors[key] = "Enter a valid email address, e.g. office@school.edu.np."
        elif input_type == "tel" and (not PHONE_RE.match(value) or not 7 <= len(re.sub(r"\D", "", value)) <= 15):
            errors[key] = "Enter a valid phone number (7 to 15 digits)."
        elif input_type == "date":
            try:
                day = datetime.strptime(value, "%Y-%m-%d").date()
            except ValueError:
                errors[key] = f"Enter the {label.lower()} as a date."
            else:
                if day > date.today():
                    errors[key] = f"{label} can't be in the future."
    return values, errors


def save_profile(user_pk, values, kind="organization"):
    table, fields = PROFILE_TABLES[kind]
    columns = [key for key, *_ in fields]
    get_db().execute(
        f"""INSERT INTO {table} (user_pk, {", ".join(columns)}, updated_at)
            VALUES (?, {", ".join("?" * len(columns))}, ?)
            ON CONFLICT(user_pk) DO UPDATE SET
            {", ".join(f"{c} = excluded.{c}" for c in columns)}, updated_at = excluded.updated_at""",
        (user_pk, *(values[c] for c in columns), now_iso()),
    )


def txn_summary(where, params):
    row = get_db().execute(
        f"""SELECT COUNT(*) AS n,
                   COALESCE(SUM(CASE WHEN {DEBIT_SQL} THEN 0 ELSE amount_cents END), 0) AS credits,
                   COALESCE(SUM(CASE WHEN {DEBIT_SQL} THEN amount_cents ELSE 0 END), 0) AS debits
            FROM transactions t WHERE {where}""",
        params,
    ).fetchone()
    return {"count": row["n"], "credits": row["credits"], "debits": row["debits"]}


# ---------------------------------------------------------------------- security


def csrf_token():
    if "csrf" not in session:
        session["csrf"] = secrets.token_urlsafe(32)
    return session["csrf"]


def check_csrf():
    token = request.form.get("csrf_token", "")
    if not token or not secrets.compare_digest(token, session.get("csrf", "")):
        abort(400, "Invalid or missing form token. Reload the page and try again.")


def safe_next(target, fallback):
    if target and target.startswith("/") and not target.startswith("//"):
        return target
    return fallback


def validate_new_password(current_ok, new, confirm, old_password):
    if not current_ok:
        return "Your current password is incorrect."
    if len(new) < MIN_PASSWORD_LEN:
        return f"New password must be at least {MIN_PASSWORD_LEN} characters."
    if new != confirm:
        return "New passwords do not match."
    if new == old_password:
        return "New password must be different from the current one."
    return None


# ------------------------------------------------------------------- formatting


def format_money(cents):
    sign = "-" if cents < 0 else ""
    cents = abs(cents)
    return f"{sign}{cents // 100:,}.{cents % 100:02d}"


def cents_to_input(cents):
    return format_money(cents).replace(",", "")


def is_debit(txn_type):
    t = (txn_type or "").strip().lower()
    return t.startswith("d") or t.startswith("withdraw")


def normalize_type(value):
    text = cell_text(value)
    low = text.lower()
    if low in ("d", "dr", "debit", "withdrawal", "withdraw"):
        return "Debit"
    if low in ("c", "cr", "credit", "deposit"):
        return "Credit"
    return text[:30]


# ------------------------------------------------------------------- parsing


def cell_text(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    if isinstance(value, datetime):
        value = value.date()
    return str(value).strip()


def normalize_user_id(value):
    """Return a 9-digit user ID string, or None if the value isn't one.

    Excel stores IDs typed as numbers without leading zeros, so numeric cells
    are zero-padded back to 9 digits.
    """
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, (int, float)) and float(value).is_integer():
        text = str(int(value)).zfill(9)
    else:
        text = re.sub(r"[\s-]", "", str(value))
    return text if USER_ID_RE.match(text) else None


def parse_amount(value):
    if value is None or value == "":
        raise ValueError("amount is empty")
    if isinstance(value, (int, float, Decimal)) and not isinstance(value, bool):
        amount = Decimal(str(value))
    else:
        text = str(value).strip().replace(",", "")
        text = re.sub(r"^[^\d\-+.(]+", "", text)  # drop currency symbols / prefixes
        negative = text.startswith("(") and text.endswith(")")
        text = text.strip("()")
        try:
            amount = Decimal(text)
        except InvalidOperation:
            raise ValueError(f"'{value}' is not a valid amount") from None
        if negative:
            amount = -amount
    if not amount.is_finite():
        raise ValueError(f"'{value}' is not a valid amount")
    return int((amount * 100).quantize(Decimal("1")))


def parse_date(value):
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y", "%d.%m.%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            continue
    return text  # keep unrecognised formats verbatim rather than rejecting the row


# Accepted header spellings (lower-cased) for each column of the upload sheet.
COLUMN_ALIASES = {
    "user": {"user id", "userid", "user", "customer id", "account id", "id", "recipient", "recipient id", "to"},
    "date": {"date", "transaction date", "txn date", "value date"},
    "description": {"description", "details", "narration", "particulars", "remarks"},
    "type": {"type", "transaction type", "txn type", "dr/cr", "debit/credit"},
    "amount": {"amount", "amt", "value"},
    "debit": {"debit", "withdrawal", "dr", "debit amount", "withdrawal amount"},
    "credit": {"credit", "deposit", "cr", "credit amount", "deposit amount"},
    "reference": {"reference", "ref", "reference no", "reference number", "cheque no", "txn id"},
}


class UploadError(Exception):
    pass


TRANSACTION_FILE_TYPES = (".xlsx", ".xlsm", ".xls", ".csv")


def read_sheet(filename, data):
    """Return the rows of the first sheet of an Excel (.xlsx/.xlsm/.xls) or CSV file as tuples."""
    ext = os.path.splitext(filename.lower())[1]
    try:
        if ext in (".xlsx", ".xlsm"):
            from openpyxl import load_workbook

            wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
            try:
                return [tuple(r) for r in wb.active.iter_rows(values_only=True)]
            finally:
                wb.close()
        if ext == ".xls":
            import xlrd

            book = xlrd.open_workbook(file_contents=data)
            sheet = book.sheet_by_index(0)
            return [
                tuple(xlrd.xldate_as_datetime(c.value, book.datemode) if c.ctype == xlrd.XL_CELL_DATE else c.value
                      for c in sheet.row(i))
                for i in range(sheet.nrows)
            ]
        if ext == ".csv":
            try:
                text = data.decode("utf-8-sig")
            except UnicodeDecodeError:
                text = data.decode("latin-1")
            return [tuple(r) for r in csv.reader(io.StringIO(text))]
    except UploadError:
        raise
    except Exception:
        raise UploadError("Could not read the file. Please check it is a valid Excel or CSV file.") from None
    raise UploadError("Transactions can only be imported from Excel (.xlsx, .xls) or CSV files.")


def parse_workbook(filename, data, need_user_column):
    """Return a list of dicts (one per non-empty row) from the first sheet, keyed by COLUMN_ALIASES."""
    rows = read_sheet(filename, data)
    header_idx = next((i for i, r in enumerate(rows) if any(cell_text(c) for c in r)), None)
    if header_idx is None:
        raise UploadError("The spreadsheet is empty.")

    columns = {}
    for idx, name in enumerate(rows[header_idx]):
        key = cell_text(name).lower()
        for field, aliases in COLUMN_ALIASES.items():
            if key in aliases and field not in columns:
                columns[field] = idx
    missing = []
    if need_user_column and "user" not in columns:
        missing.append("User ID")
    if not ({"amount", "debit", "credit"} & columns.keys()):
        missing.append("Amount (or Debit/Credit)")
    if missing:
        raise UploadError(
            "Missing required column(s): " + ", ".join(missing)
            + ". Download the template to see the expected layout."
        )

    records = []
    for line_no in range(header_idx + 2, len(rows) + 1):  # spreadsheet row numbers are 1-based
        row = rows[line_no - 1]
        if not any(cell_text(c) for c in row):
            continue
        rec = {"line": line_no}
        for field in COLUMN_ALIASES:
            idx = columns.get(field)
            rec[field] = row[idx] if idx is not None and idx < len(row) else None
        records.append(rec)
        if len(records) > MAX_ROWS:
            raise UploadError(f"Too many rows; the limit is {MAX_ROWS} per upload.")
    if not records:
        raise UploadError("The spreadsheet has a header row but no transactions.")
    return records


# ------------------------------------------------------------------ documents

# Files the admin may attach to a user. Images and PDFs open in the browser; everything else downloads.
DOCUMENT_TYPES = {
    ".pdf", ".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp", ".heic", ".tif", ".tiff",
    ".doc", ".docx", ".rtf", ".odt", ".txt",
    ".xls", ".xlsx", ".xlsm", ".csv", ".ods",
    ".ppt", ".pptx", ".odp",
}
INLINE_TYPES = {".pdf", ".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp"}
DOCUMENT_TYPES_LABEL = "images, PDF, Word, Excel, PowerPoint, text"


def documents_dir():
    path = current_app.config["DOCUMENTS_DIR"]
    os.makedirs(path, exist_ok=True)
    return path


def document_ext(filename):
    return os.path.splitext(filename.lower())[1]


def save_document(user_pk, file, title):
    """Store an uploaded file for a user. Raises UploadError for unsupported or empty files."""
    original = os.path.basename(file.filename or "").strip()[:200]
    ext = document_ext(original)
    if ext not in DOCUMENT_TYPES:
        raise UploadError(f"'{original}': this file type isn't allowed. Allowed: {DOCUMENT_TYPES_LABEL}.")
    data = file.read()
    if not data:
        raise UploadError(f"'{original}' is empty.")
    stored = secrets.token_hex(16) + ext
    with open(os.path.join(documents_dir(), stored), "wb") as fh:
        fh.write(data)
    get_db().execute(
        """INSERT INTO documents (user_pk, title, original_name, stored_name, size, created_at)
           VALUES (?, ?, ?, ?, ?, ?)""",
        (user_pk, (title or os.path.splitext(original)[0])[:200], original, stored, len(data), now_iso()),
    )
    return original


def delete_document_files(stored_names):
    for name in stored_names:
        try:
            os.remove(os.path.join(documents_dir(), name))
        except FileNotFoundError:
            pass


def send_document(doc, inline=False):
    """Send a stored document; images and PDFs may be shown in the browser, everything else downloads."""
    ext = document_ext(doc["original_name"])
    inline = inline and ext in INLINE_TYPES
    mimetype = mimetypes.guess_type(doc["original_name"])[0] or "application/octet-stream"
    response = send_file(os.path.join(documents_dir(), doc["stored_name"]), mimetype=mimetype,
                         as_attachment=not inline, download_name=doc["original_name"])
    response.headers["X-Content-Type-Options"] = "nosniff"
    return response


def format_size(size):
    for unit in ("bytes", "KB", "MB"):
        if size < 1024 or unit == "MB":
            return f"{size:.0f} {unit}" if unit == "bytes" else f"{size:.1f} {unit}"
        size /= 1024


def record_amount_and_type(rec):
    """Work out (amount_cents, type) from either an Amount column or Debit/Credit columns."""
    if cell_text(rec["amount"]):
        return parse_amount(rec["amount"]), normalize_type(rec["type"])
    if cell_text(rec["debit"]):
        return abs(parse_amount(rec["debit"])), "Debit"
    if cell_text(rec["credit"]):
        return abs(parse_amount(rec["credit"])), "Credit"
    raise ValueError("amount is empty")


# ----------------------------------------------------------------------- pdf


def escape_xml(text):
    return str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


PDF_ROWS_PER_PAGE = 25
# (label, profile column, role shown under the name) for the signature blocks at the foot of each page.
PDF_SIGNATURES = [
    ("Prepared by", "accountant", "Accountant"),
    ("Checked by", "head_name", "Head of the organization"),
    ("Approved by", "chairperson", "Chairperson"),
]
PDF_BRAND = colors.HexColor("#1f4e79")
PDF_MUTED = colors.HexColor("#555555")


class _NumberedCanvas(canvas.Canvas):
    """Canvas that knows the total page count, so pages can say "Page 2 of 5"."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._pages = []

    def showPage(self):
        self._pages.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        total = len(self._pages)
        for state in self._pages:
            self.__dict__.update(state)
            self.setFont("Helvetica", 8)
            self.setFillColor(PDF_MUTED)
            self.drawCentredString(A4[0] / 2, 8 * mm, f"Page {self._pageNumber} of {total}")
            super().showPage()
        super().save()


def build_pdf(rows, profile, details):
    """Transaction statement: organization header, 25 rows per page, signature blocks at the foot.

    `profile` is the organization profile row (or None); `details` is a one-line description
    such as the user ID, period and generation time.
    """
    page_w, page_h = A4
    margin = 14 * mm
    width = page_w - 2 * margin
    styles = getSampleStyleSheet()
    name_style = ParagraphStyle("org", parent=styles["Title"], fontName="Helvetica-Bold", fontSize=17,
                                leading=21, alignment=TA_CENTER, spaceAfter=0)
    address_style = ParagraphStyle("addr", parent=styles["Normal"], fontSize=10, leading=13, alignment=TA_CENTER)
    details_style = ParagraphStyle("details", parent=styles["Normal"], fontSize=8.5, leading=11,
                                   alignment=TA_CENTER, textColor=PDF_MUTED)
    cell = ParagraphStyle("cell", parent=styles["BodyText"], fontSize=8.5, leading=10.5)

    org_name = profile["full_name"] if profile else "Transaction Statement"
    address = ", ".join(line.strip() for line in (profile["address"] if profile else "").splitlines() if line.strip())

    def draw_page(c, _doc):
        c.saveState()
        # Header: organization name (bold), address below it in smaller type, then statement details.
        y = page_h - 12 * mm
        for text, style in ((org_name, name_style), (address, address_style), (details, details_style)):
            if not text:
                continue
            p = Paragraph(escape_xml(text), style)
            _, h = p.wrap(width, 40 * mm)
            y -= h
            p.drawOn(c, margin, y)
            y -= 1.5 * mm
        c.setStrokeColor(PDF_BRAND)
        c.setLineWidth(1)
        c.line(margin, y, page_w - margin, y)

        # Footer: Prepared by (left), Checked by (centre), Approved by (right).
        col_w = width / 3
        for i, (label, column, role) in enumerate(PDF_SIGNATURES):
            cx = margin + col_w * i + col_w / 2
            c.setStrokeColor(colors.black)
            c.setLineWidth(0.6)
            c.line(cx - 24 * mm, 32 * mm, cx + 24 * mm, 32 * mm)
            c.setFillColor(colors.black)
            c.setFont("Helvetica-Bold", 9.5)
            c.drawCentredString(cx, 27 * mm, label)
            name = (profile[column] if profile else "") or ""
            c.setFont("Helvetica", 9)
            c.drawCentredString(cx, 22.5 * mm, name[:45])
            c.setFont("Helvetica", 8)
            c.setFillColor(PDF_MUTED)
            c.drawCentredString(cx, 18.5 * mm, role)
        c.restoreState()

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=margin, rightMargin=margin,
                            topMargin=38 * mm, bottomMargin=40 * mm, title=f"{org_name} - Transaction Statement")

    header = ["S.N.", "Date", "Description", "Reference", "Debit", "Credit"]
    widths = [11 * mm, 21 * mm, None, 30 * mm, 25 * mm, 25 * mm]
    widths[2] = width - sum(w for w in widths if w)
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), PDF_BRAND),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("ALIGN", (-2, 0), (-1, -1), "RIGHT"),
        ("ALIGN", (0, 0), (0, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#b8c4d0")),
        ("TOPPADDING", (0, 0), (-1, -1), 2.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
    ]

    story = []
    credits = debits = 0
    chunks = [rows[i:i + PDF_ROWS_PER_PAGE] for i in range(0, len(rows), PDF_ROWS_PER_PAGE)] or [[]]
    for page_no, chunk in enumerate(chunks):
        data = [header]
        for i, r in enumerate(chunk, start=page_no * PDF_ROWS_PER_PAGE + 1):
            amount = format_money(r["amount_cents"])
            if is_debit(r["txn_type"]):
                debits += r["amount_cents"]
                debit, credit = amount, ""
            else:
                credits += r["amount_cents"]
                debit, credit = "", amount
            data.append([str(i), r["txn_date"] or "", Paragraph(escape_xml(r["description"] or ""), cell),
                         Paragraph(escape_xml(r["reference"] or ""), cell), debit, credit])
        page_style = list(style)
        if not chunk:
            data.append(["", "", "No transactions.", "", "", ""])
        if page_no == len(chunks) - 1:
            data.append(["", "", "Total", "", format_money(debits), format_money(credits)])
            data.append(["", "", "Net balance (credit - debit)", "", "", format_money(credits - debits)])
            page_style += [
                ("FONTNAME", (0, -2), (-1, -1), "Helvetica-Bold"),
                ("BACKGROUND", (0, -2), (-1, -1), colors.HexColor("#eef3f8")),
                ("LINEABOVE", (0, -2), (-1, -2), 1, PDF_BRAND),
            ]
        else:
            data.append(["", "", "Carried forward", "", format_money(debits), format_money(credits)])
            page_style += [("FONTNAME", (0, -1), (-1, -1), "Helvetica-Oblique"),
                           ("TEXTCOLOR", (0, -1), (-1, -1), PDF_MUTED)]
        table = Table(data, colWidths=widths)
        table.setStyle(TableStyle(page_style))
        story.append(table)
        if page_no < len(chunks) - 1:
            story.append(PageBreak())

    doc.build(story, onFirstPage=draw_page, onLaterPages=draw_page, canvasmaker=_NumberedCanvas)
    buf.seek(0)
    return buf


def statement_pdf(user_pk, user_id, date_from="", date_to=""):
    """PDF statement of one user's transactions, oldest first, optionally limited to a date range."""
    where, params = "t.user_pk = ?", [user_pk]
    if date_from:
        where += " AND t.txn_date >= ?"
        params.append(date_from)
    if date_to:
        where += " AND t.txn_date <= ?"
        params.append(date_to)
    rows = get_db().execute(f"{TXN_SELECT} WHERE {where} {TXN_ORDER_OLDEST_FIRST}", params).fetchall()
    details = f"Transaction Statement  |  User ID {user_id}"
    if date_from or date_to:
        details += f"  |  Period {date_from or 'start'} to {date_to or date.today().isoformat()}"
    details += f"  |  Generated {datetime.now():%Y-%m-%d %H:%M}"
    return build_pdf(rows, get_profile(user_pk), details)
