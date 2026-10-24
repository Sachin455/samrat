# TxnShare: transaction statements for account holders

A Flask web app with two sides:

**Admin portal** (`/admin`): a single admin account that
- adds account holders by their **9-digit user ID** (one at a time or many pasted at once),
- sets the shared **first-time password** new users log in with, and can reset any user back to it,
- **uploads Excel files** of transactions, either to one chosen user ID or routed per row by a `User ID` column
  (optionally creating accounts for IDs that don't exist yet),
- has full access to **every user's transactions**: view, filter, add, edit (including moving to another user),
  delete, delete a whole upload, and download any user's statement as PDF.

**User side** (`/`): each account holder
- logs in with their 9-digit ID and the first-time password, and is **required to choose their own password**
  on that first login (they can change it again any time from *Change password*),
- sees a **personal dashboard** with totals and their **latest 5 transactions**,
- can view all their transactions, filter by date, and **download a PDF statement**.

## Running locally

Requires Python 3.9 or newer.

```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
export SECRET_KEY="change-me"   # needed so logins survive restarts
python app.py                   # http://localhost:5000  (PORT=5001 python app.py to use another port)
```

| Portal | URL | First login |
|---|---|---|
| Admin | http://localhost:5000/admin | username `admin`, password `admin123` (you must change it on first login) |
| Users | http://localhost:5000 | their 9-digit user ID + the first-time password (default `Welcome@123`, shown and editable on the admin *Users* page) |

You can set different starting values before the first run with the `ADMIN_USERNAME`, `ADMIN_PASSWORD` and
`DEFAULT_USER_PASSWORD` environment variables.

Data is stored in SQLite at `instance/bank.db` (change it with the `DATABASE` environment variable).
For production, run it behind a WSGI server, e.g. `gunicorn "app:create_app()"`, with a fixed `SECRET_KEY` and HTTPS.

## Excel format

The first row must be a header. Column names are case-insensitive and common alternatives are accepted.
A template can be downloaded from the admin *Upload* page.

| Column      | Required | Example |
|-------------|----------|---------|
| User ID     | yes, unless you enter one user ID on the upload form | `100000001` |
| Amount      | yes (or separate `Debit`/`Credit` or `Withdrawal`/`Deposit` columns) | `2500.00`, `1,250`, `(5.00)` |
| Date        | no | Excel date, `2026-10-24`, `24/10/2026` |
| Description | no | `Salary October` |
| Type        | no | `Credit` / `Debit` (also `Cr` / `Dr`) |
| Reference   | no | `TXN-0001` |

User IDs that Excel stored as numbers (and so lost their leading zeros) are padded back to 9 digits.
Rows with an invalid user ID or amount are skipped, and the upload result lists each one with the reason.

## Code layout

- `app.py`: creates the app and wires everything together
- `core.py`: database, Excel parsing, PDF generation, shared helpers
- `user_views.py`: account-holder pages
- `admin_views.py`: admin portal
- `templates/`, `static/style.css`: the web pages and styling
- `tests/test_app.py`: tests (`pip install pytest && python -m pytest`)
