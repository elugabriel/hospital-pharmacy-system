"""
migrate_to_sqlite.py
--------------------
Migrates ALL tables from the hospital PostgreSQL database to a local SQLite file.

Tables covered (23):
  admin_users, cashier_users, admin_audit_logs, pharmacists, billing_users,
  drugs, drug_sales, receipts, receipt_items, stock_movements,
  billing_invoice, billing_receipt, payments, users,
  hr_users, departments, staff, attendance, leaves,
  schedules, payroll, documents, shift_swap_requests

Changes vs previous version:
  - pharmacists   : removed full_name / created_by (not in app.py)
  - billing_users : removed full_name / created_by / is_active (not in app.py)
  - receipts      : removed pharmacist column (not in app.py); column is optional fallback
  - shift_swap_requests : added `notes` column (added in app.py)
  - All TABLE_SPECS column lists updated to match app.py exactly
  - copy_from_postgres uses safe_cols to skip columns absent in either side

Usage:
  python migrate_to_sqlite.py
  python migrate_to_sqlite.py --pg-url postgresql://user:pass@host:5432/dbname
  python migrate_to_sqlite.py --out hospital.db
  python migrate_to_sqlite.py --fresh   (skip PostgreSQL copy, create empty DB with default users only)
"""

import sqlite3
import argparse
import json
import sys
import os
from datetime import datetime

# ──────────────────────────────────────────────
# CONFIGURATION
# ──────────────────────────────────────────────
DEFAULT_PG_URL = "postgresql://flask_user:Olarewaju1.@localhost:5432/hospital_db"
SQLITE_FILE    = "hospital.db"

# ──────────────────────────────────────────────
# SQLite SCHEMA  (all 23 tables, aligned with app.py)
# ──────────────────────────────────────────────
SCHEMA = """
-- ── USER / AUTH TABLES ──────────────────────────────────────────────────────

-- Extra admin table (forward-compatible; not used directly in app.py routes yet)
CREATE TABLE IF NOT EXISTS admin_users (
    id                INTEGER PRIMARY KEY AUTOINCREMENT,
    username          TEXT UNIQUE NOT NULL,
    password          TEXT NOT NULL,
    full_name         TEXT,
    email             TEXT,
    role              TEXT DEFAULT 'Admin',
    is_super_admin    INTEGER DEFAULT 0,
    is_active         INTEGER DEFAULT 1,
    created_at        TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    created_by        INTEGER,
    last_login        TIMESTAMP
);

-- Extra cashier table (forward-compatible)
CREATE TABLE IF NOT EXISTS cashier_users (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    username    TEXT UNIQUE NOT NULL,
    password    TEXT NOT NULL,
    full_name   TEXT,
    email       TEXT,
    is_active   INTEGER DEFAULT 1,
    created_at  TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    created_by  INTEGER,
    last_login  TIMESTAMP
);

-- Matches app.py pharmacists table exactly
CREATE TABLE IF NOT EXISTS pharmacists (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    username    TEXT UNIQUE NOT NULL,
    password    TEXT NOT NULL,
    is_active   INTEGER DEFAULT 1,
    created_at  TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- Matches app.py billing_users table exactly
CREATE TABLE IF NOT EXISTS billing_users (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    username    TEXT UNIQUE NOT NULL,
    password    TEXT NOT NULL,
    created_at  TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- Matches app.py hr_users table exactly
CREATE TABLE IF NOT EXISTS hr_users (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    username    TEXT UNIQUE NOT NULL,
    password    TEXT NOT NULL,
    full_name   TEXT NOT NULL,
    email       TEXT,
    role        TEXT DEFAULT 'HR Staff',
    is_active   INTEGER DEFAULT 1,
    created_at  TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- Matches app.py users table exactly
CREATE TABLE IF NOT EXISTS users (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    username    TEXT UNIQUE NOT NULL,
    password    TEXT NOT NULL,
    created_at  TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- ── AUDIT ────────────────────────────────────────────────────────────────────

CREATE TABLE IF NOT EXISTS admin_audit_logs (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    admin_id    INTEGER,
    action      TEXT NOT NULL,
    details     TEXT,
    ip_address  TEXT,
    user_agent  TEXT,
    created_at  TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- ── PHARMACY TABLES ──────────────────────────────────────────────────────────

CREATE TABLE IF NOT EXISTS drugs (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    name                TEXT NOT NULL,
    strength            TEXT NOT NULL,
    unit_price          REAL NOT NULL,
    stock_quantity      INTEGER NOT NULL,
    expiry_date         DATE NOT NULL,
    low_stock_threshold INTEGER DEFAULT 20,
    created_at          TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at          TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS drug_sales (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    receipt_no      TEXT UNIQUE NOT NULL,
    patient_name    TEXT,
    patient_id      TEXT,
    items           TEXT NOT NULL,   -- stored as JSON string
    subtotal        REAL NOT NULL,
    discount        REAL DEFAULT 0.00,
    tax             REAL DEFAULT 0.00,
    grand_total     REAL NOT NULL,
    pharmacist      TEXT NOT NULL,
    created_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- Matches app.py receipts table exactly (no pharmacist column)
CREATE TABLE IF NOT EXISTS receipts (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    patient_name    TEXT,
    patient_id      TEXT,
    subtotal        REAL NOT NULL,
    discount        REAL DEFAULT 0.00,
    tax             REAL DEFAULT 0.00,
    total_amount    REAL NOT NULL,
    grand_total     REAL NOT NULL,
    created_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS receipt_items (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    receipt_id  INTEGER NOT NULL,
    drug_name   TEXT NOT NULL,
    strength    TEXT NOT NULL,
    quantity    INTEGER NOT NULL,
    unit_price  REAL NOT NULL,
    FOREIGN KEY (receipt_id) REFERENCES receipts(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS stock_movements (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    drug_id         INTEGER NOT NULL,
    movement_type   TEXT NOT NULL,
    quantity        INTEGER NOT NULL,
    user_id         INTEGER NOT NULL,
    note            TEXT,
    created_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (drug_id) REFERENCES drugs(id) ON DELETE CASCADE
);

-- ── BILLING TABLES ───────────────────────────────────────────────────────────

CREATE TABLE IF NOT EXISTS billing_invoice (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    patient_name    TEXT NOT NULL,
    service_type    TEXT NOT NULL,
    amount          REAL NOT NULL,
    status          TEXT DEFAULT 'UNPAID',
    created_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS billing_receipt (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    invoice_id      INTEGER NOT NULL,
    amount_paid     REAL NOT NULL,
    payment_method  TEXT NOT NULL,
    received_by     TEXT NOT NULL,
    payment_date    TIMESTAMP NOT NULL,
    FOREIGN KEY (invoice_id) REFERENCES billing_invoice(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS payments (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    patient_name    TEXT NOT NULL,
    service_type    TEXT NOT NULL,
    subtotal        REAL NOT NULL,
    discount        REAL DEFAULT 0.00,
    tax             REAL DEFAULT 0.00,
    grand_total     REAL NOT NULL,
    amount_paid     REAL NOT NULL,
    balance         REAL NOT NULL,
    payment_method  TEXT NOT NULL,
    status          TEXT NOT NULL,
    payment_date    DATE NOT NULL,
    recorded_by     INTEGER NOT NULL,
    created_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- ── HR TABLES ────────────────────────────────────────────────────────────────

CREATE TABLE IF NOT EXISTS departments (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    name            TEXT NOT NULL,
    code            TEXT UNIQUE NOT NULL,
    description     TEXT,
    head_of_dept    TEXT,
    status          TEXT DEFAULT 'Active',
    created_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS staff (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    staff_id            TEXT UNIQUE NOT NULL,
    first_name          TEXT NOT NULL,
    last_name           TEXT NOT NULL,
    department_id       INTEGER,
    position            TEXT NOT NULL,
    employment_type     TEXT,
    email               TEXT,
    phone               TEXT,
    hire_date           DATE NOT NULL,
    salary              REAL,
    status              TEXT DEFAULT 'Active',
    emergency_contact   TEXT,
    address             TEXT,
    created_at          TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at          TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (department_id) REFERENCES departments(id)
);

CREATE TABLE IF NOT EXISTS attendance (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    staff_id    INTEGER,
    date        DATE NOT NULL,
    check_in    TIME,
    check_out   TIME,
    status      TEXT,
    remarks     TEXT,
    recorded_by INTEGER,
    created_at  TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(staff_id, date),
    FOREIGN KEY (staff_id) REFERENCES staff(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS leaves (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    staff_id        INTEGER,
    leave_type      TEXT NOT NULL,
    start_date      DATE NOT NULL,
    end_date        DATE NOT NULL,
    days_requested  INTEGER NOT NULL,
    reason          TEXT,
    status          TEXT DEFAULT 'Pending',
    approved_by     INTEGER,
    approved_at     TIMESTAMP,
    created_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (staff_id) REFERENCES staff(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS schedules (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    staff_id        INTEGER,
    schedule_date   DATE NOT NULL,
    shift_type      TEXT,
    start_time      TIME NOT NULL,
    end_time        TIME NOT NULL,
    location        TEXT,
    notes           TEXT,
    created_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (staff_id) REFERENCES staff(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS payroll (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    staff_id        INTEGER,
    pay_period      TEXT,
    basic_salary    REAL,
    allowances      REAL,
    deductions      REAL,
    net_salary      REAL,
    status          TEXT DEFAULT 'Pending',
    payment_date    DATE,
    created_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (staff_id) REFERENCES staff(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS documents (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    staff_id        INTEGER,
    document_type   TEXT,
    document_name   TEXT,
    file_path       TEXT,
    uploaded_by     INTEGER,
    uploaded_at     TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (staff_id) REFERENCES staff(id) ON DELETE CASCADE
);

-- Matches app.py shift_swap_requests table exactly (includes `notes` column)
CREATE TABLE IF NOT EXISTS shift_swap_requests (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    schedule_id      INTEGER NOT NULL,
    from_staff_id    INTEGER NOT NULL,
    to_staff_id      INTEGER NOT NULL,
    reason           TEXT,
    status           TEXT DEFAULT 'Pending',
    requested_by     INTEGER,
    requested_at     TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    approved_by      INTEGER,
    approved_at      TIMESTAMP,
    reviewed_by      INTEGER,
    reviewed_at      TIMESTAMP,
    rejection_reason TEXT,
    notes            TEXT,
    FOREIGN KEY (schedule_id)   REFERENCES schedules(id) ON DELETE CASCADE,
    FOREIGN KEY (from_staff_id) REFERENCES staff(id) ON DELETE CASCADE,
    FOREIGN KEY (to_staff_id)   REFERENCES staff(id) ON DELETE CASCADE
);

-- ── INDEXES ──────────────────────────────────────────────────────────────────

CREATE INDEX IF NOT EXISTS idx_staff_department       ON staff(department_id);
CREATE INDEX IF NOT EXISTS idx_attendance_staff_date  ON attendance(staff_id, date);
CREATE INDEX IF NOT EXISTS idx_leaves_staff_status    ON leaves(staff_id, status);
CREATE INDEX IF NOT EXISTS idx_schedules_staff_date   ON schedules(staff_id, schedule_date);
CREATE INDEX IF NOT EXISTS idx_payroll_staff_period   ON payroll(staff_id, pay_period);
CREATE INDEX IF NOT EXISTS idx_staff_status           ON staff(status);
CREATE INDEX IF NOT EXISTS idx_attendance_date        ON attendance(date);
CREATE INDEX IF NOT EXISTS idx_leaves_status          ON leaves(status);
CREATE INDEX IF NOT EXISTS idx_shift_swap_status      ON shift_swap_requests(status);
CREATE INDEX IF NOT EXISTS idx_shift_swap_schedule    ON shift_swap_requests(schedule_id);
"""

# ──────────────────────────────────────────────
# TABLE COPY SPECS
# Each tuple: (table_name, comma-separated columns to copy from PostgreSQL)
# - Columns listed must exist in the SQLite schema above.
# - The copy function automatically skips any column absent in PostgreSQL,
#   so it is safe to list columns that may not exist in older PG databases.
# ──────────────────────────────────────────────
TABLE_SPECS = [
    # Auth / user tables
    ("admin_users",         "username, password, full_name, email, role, is_super_admin, is_active, created_at, created_by, last_login"),
    ("cashier_users",       "username, password, full_name, email, is_active, created_at, created_by, last_login"),
    # pharmacists: app.py has no full_name / created_by
    ("pharmacists",         "username, password, is_active, created_at"),
    # billing_users: app.py has only username / password / created_at
    ("billing_users",       "username, password, created_at"),
    ("hr_users",            "username, password, full_name, email, role, is_active, created_at"),
    ("users",               "username, password, created_at"),
    ("admin_audit_logs",    "admin_id, action, details, ip_address, user_agent, created_at"),
    # Pharmacy
    ("drugs",               "name, strength, unit_price, stock_quantity, expiry_date, low_stock_threshold, created_at, updated_at"),
    ("drug_sales",          "receipt_no, patient_name, patient_id, items, subtotal, discount, tax, grand_total, pharmacist, created_at"),
    # receipts: app.py has no pharmacist column; skipped automatically if absent in PG
    ("receipts",            "patient_name, patient_id, subtotal, discount, tax, total_amount, grand_total, created_at"),
    ("receipt_items",       "receipt_id, drug_name, strength, quantity, unit_price"),
    ("stock_movements",     "drug_id, movement_type, quantity, user_id, note, created_at"),
    # Billing
    ("billing_invoice",     "patient_name, service_type, amount, status, created_at"),
    ("billing_receipt",     "invoice_id, amount_paid, payment_method, received_by, payment_date"),
    ("payments",            "patient_name, service_type, subtotal, discount, tax, grand_total, amount_paid, balance, payment_method, status, payment_date, recorded_by, created_at"),
    # HR
    ("departments",         "name, code, description, head_of_dept, status, created_at"),
    ("staff",               "staff_id, first_name, last_name, department_id, position, employment_type, email, phone, hire_date, salary, status, emergency_contact, address, created_at, updated_at"),
    ("attendance",          "staff_id, date, check_in, check_out, status, remarks, recorded_by, created_at"),
    ("leaves",              "staff_id, leave_type, start_date, end_date, days_requested, reason, status, approved_by, approved_at, created_at"),
    ("schedules",           "staff_id, schedule_date, shift_type, start_time, end_time, location, notes, created_at"),
    ("payroll",             "staff_id, pay_period, basic_salary, allowances, deductions, net_salary, status, payment_date, created_at"),
    ("documents",           "staff_id, document_type, document_name, file_path, uploaded_by, uploaded_at"),
    # shift_swap_requests: includes notes column added in app.py
    ("shift_swap_requests", "schedule_id, from_staff_id, to_staff_id, reason, status, requested_by, requested_at, approved_by, approved_at, reviewed_by, reviewed_at, rejection_reason, notes"),
]


# ──────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────

def _coerce_value(val):
    """Convert PostgreSQL-specific types to SQLite-compatible values."""
    if isinstance(val, dict):
        return json.dumps(val)
    if isinstance(val, bool):
        return 1 if val else 0
    return val


def create_sqlite_schema(sqlite_conn):
    print("Creating SQLite schema...")
    sqlite_conn.executescript(SCHEMA)
    sqlite_conn.commit()
    print("  ✔ All 23 tables + indexes created.\n")


def insert_default_users(sqlite_conn):
    """Insert default users for every auth module."""
    try:
        from werkzeug.security import generate_password_hash
    except ImportError:
        print("  ⚠ werkzeug not installed — skipping default users. Run: pip install werkzeug")
        return

    # (table, username, raw_password, extra_fields_dict)
    defaults = [
        ("admin_users",   "admin",       "admin123",   {"full_name": "Super Admin",   "is_super_admin": 1, "is_active": 1}),
        ("pharmacists",   "pharmacist1", "pharma123",  {"is_active": 1}),
        ("billing_users", "billing1",    "billing123", {}),
        ("hr_users",      "hr_admin",    "hr@admin123",{"full_name": "HR Administrator", "email": "admin@hospital.com", "role": "HR Manager", "is_active": 1}),
        ("hr_users",      "hr_staff",    "hr@admin123",{"full_name": "HR Staff",         "email": "staff@hospital.com", "role": "HR Officer",  "is_active": 1}),
    ]

    cursor = sqlite_conn.cursor()
    for table, username, raw_pw, extras in defaults:
        hashed = generate_password_hash(raw_pw)
        cols   = ["username", "password"] + list(extras.keys())
        vals   = [username, hashed]       + list(extras.values())
        placeholders = ", ".join(["?" for _ in cols])
        col_str      = ", ".join(cols)
        cursor.execute(
            f"INSERT OR IGNORE INTO {table} ({col_str}) VALUES ({placeholders})",
            vals
        )
    sqlite_conn.commit()
    print("  ✔ Default users inserted.\n")
    print("    Module      | Username      | Password")
    print("    ------------|---------------|----------")
    print("    Admin       | admin         | admin123")
    print("    Pharmacy    | pharmacist1   | pharma123")
    print("    Billing     | billing1      | billing123")
    print("    HR          | hr_admin      | hr@admin123")
    print("    HR          | hr_staff      | hr@admin123\n")


def copy_from_postgres(pg_url, sqlite_conn):
    """Copy all tables from PostgreSQL into the SQLite database."""
    try:
        import psycopg2
    except ImportError:
        print("  ✗ psycopg2 not installed. Run: pip install psycopg2-binary")
        return False

    print(f"Connecting to PostgreSQL: {pg_url[:40]}...")
    try:
        pg_conn   = psycopg2.connect(pg_url)
        pg_cursor = pg_conn.cursor()
        print("  ✔ Connected.\n")
    except Exception as e:
        print(f"  ✗ Could not connect to PostgreSQL: {e}")
        return False

    sqlite_cursor = sqlite_conn.cursor()
    total_copied  = 0
    errors        = []

    for table, columns in TABLE_SPECS:
        col_list = [c.strip() for c in columns.split(",")]

        # ── Check PG table exists ──────────────────────────
        pg_cursor.execute(
            "SELECT EXISTS (SELECT 1 FROM information_schema.tables WHERE table_name = %s)",
            (table,)
        )
        if not pg_cursor.fetchone()[0]:
            print(f"  ⚠  {table:<30} not found in PostgreSQL — skipping.")
            continue

        # ── Find columns present in BOTH PG and our spec ──
        pg_cursor.execute(
            "SELECT column_name FROM information_schema.columns WHERE table_name = %s",
            (table,)
        )
        pg_cols   = {row[0] for row in pg_cursor.fetchall()}
        safe_cols = [c for c in col_list if c in pg_cols]

        if not safe_cols:
            print(f"  ⚠  {table:<30} no matching columns — skipping.")
            continue

        skipped_cols = set(col_list) - set(safe_cols)
        if skipped_cols:
            print(f"  ℹ  {table:<30} columns not in PG (will be NULL): {', '.join(sorted(skipped_cols))}")

        safe_col_str    = ", ".join(safe_cols)
        safe_placeholders = ", ".join(["?" for _ in safe_cols])

        try:
            pg_cursor.execute(f"SELECT {safe_col_str} FROM {table}")
            rows = pg_cursor.fetchall()

            converted = [
                tuple(_coerce_value(v) for v in row)
                for row in rows
            ]

            sqlite_cursor.executemany(
                f"INSERT OR IGNORE INTO {table} ({safe_col_str}) VALUES ({safe_placeholders})",
                converted
            )
            sqlite_conn.commit()
            total_copied += len(rows)
            print(f"  ✔  {table:<30} {len(rows):>6} rows copied.")

        except Exception as e:
            sqlite_conn.rollback()
            msg = f"Error copying '{table}': {e}"
            errors.append(msg)
            print(f"  ✗  {msg}")

    pg_cursor.close()
    pg_conn.close()

    print(f"\n  Total rows copied : {total_copied}")
    if errors:
        print(f"  Errors            : {len(errors)}")
        for err in errors:
            print(f"    - {err}")

    return True


def verify(sqlite_conn):
    """Print row counts for every table after migration."""
    print("── Verification ─────────────────────────────────────")
    cursor = sqlite_conn.cursor()
    grand_total = 0
    for table, _ in TABLE_SPECS:
        try:
            cursor.execute(f"SELECT COUNT(*) FROM {table}")
            count = cursor.fetchone()[0]
            grand_total += count
            status = "✔" if count > 0 else "○"
            print(f"  {status}  {table:<30} {count:>6} rows")
        except Exception as e:
            print(f"  ✗  {table:<30} ERROR: {e}")
    print(f"─────────────────────────────────────────────────────")
    print(f"     Grand total: {grand_total} rows\n")


# ──────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Migrate PostgreSQL → SQLite for the hospital management system"
    )
    parser.add_argument("--pg-url", default=DEFAULT_PG_URL,
                        help="PostgreSQL connection URL (default: uses env DATABASE_URL or built-in default)")
    parser.add_argument("--out",    default=SQLITE_FILE,
                        help=f"Output SQLite file path (default: {SQLITE_FILE})")
    parser.add_argument("--fresh",  action="store_true",
                        help="Skip PostgreSQL copy; create an empty DB with default users only")
    args = parser.parse_args()

    # Allow DATABASE_URL env var to override default
    pg_url = os.environ.get("DATABASE_URL", args.pg_url)

    print("=" * 58)
    print("  Hospital System — PostgreSQL → SQLite Migration")
    print("=" * 58)
    print(f"  Output file : {args.out}")
    print(f"  PG URL      : {pg_url[:45]}{'...' if len(pg_url) > 45 else ''}")
    print(f"  Mode        : {'fresh (no data copy)' if args.fresh else 'full migration'}")
    print()

    sqlite_conn = sqlite3.connect(args.out)
    sqlite_conn.execute("PRAGMA foreign_keys = ON")
    sqlite_conn.execute("PRAGMA journal_mode = WAL")

    create_sqlite_schema(sqlite_conn)

    if not args.fresh:
        ok = copy_from_postgres(pg_url, sqlite_conn)
        if not ok:
            print("\n⚠  PostgreSQL copy failed — creating empty DB with default users.\n")

    print("Inserting / verifying default users...")
    insert_default_users(sqlite_conn)

    verify(sqlite_conn)

    sqlite_conn.close()
    size_kb = os.path.getsize(args.out) / 1024
    print(f"✅  Done!  SQLite database saved to: {args.out}  ({size_kb:.1f} KB)")


if __name__ == "__main__":
    main()