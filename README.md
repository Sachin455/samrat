# TxnShare: share bank transactions between users

A small Flask web app where:

1. **Users register** with their own unique user ID and a password (passwords are stored hashed).
2. **A user uploads an Excel (.xlsx) sheet** of bank transactions. Each row has a `Recipient`
   column naming another user's ID, and the app delivers that row to that user.
3. **Recipients see the transactions sent to them** on their dashboard and can **download them as a PDF**
   (all of them, or only the ones from one sender). Uploaders can also view each upload and download it as a PDF.

## Excel format

The first row must be a header. Column names are case-insensitive, and common alternatives are accepted
(for example `User ID` or `To` for Recipient, `Narration` for Description).

| Column      | Required | Example                                  |
|-------------|----------|------------------------------------------|
| Recipient   | yes      | `alice`                                  |
| Amount      | yes      | `2500.00`, `1,250`, `(5.00)`             |
| Date        | no       | Excel date or `2026-10-24`, `24/10/2026` |
| Description | no       | `Invoice #1001 payment`                  |
| Type        | no       | `Credit` / `Debit`                       |
| Reference   | no       | `TXN-0001`                               |

A ready-made template can be downloaded from the Upload page. Rows with an unknown recipient or an
invalid amount are skipped, and the upload result lists each skipped row with the reason.

## Running locally

```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
export SECRET_KEY="change-me"   # needed so logins survive restarts
python app.py                   # http://localhost:5000
```

Data is stored in SQLite at `instance/transactions.db` (change it with the `DATABASE` environment variable).
For production, run it behind a WSGI server, e.g. `gunicorn "app:create_app()"`, with a fixed `SECRET_KEY` and HTTPS.

## Tests

```bash
pip install pytest
python -m pytest
```
