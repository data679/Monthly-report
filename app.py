"""Executive Summary logger -- local web app.

Run:  python3 app.py        then open http://127.0.0.1:8765
Upload a Denticon Executive Summary (.pdf or .xlsx), review the numbers per
office, and save them to the log (SQLite, mirrored to CSV after every change).
"""

import base64
import csv
import getpass
import hashlib
import hmac
import json
import os
import re
import secrets
import socket
import sqlite3
import sys
import tempfile
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote

import threading
import time
from datetime import date, timedelta

import google_sheets
import np_analysis
from sheet_paste import METRICS, PasteError, candidates, normalize, parse_paste, rows_to_text, title_case
from summary_parser import ParseError, parse_report

ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
DB_PATH = DATA_DIR / "reports.db"
CSV_PATH = DATA_DIR / "office_summary.csv"
# HOST=0.0.0.0 opens the app to the office network; that needs a password (python3 app.py --set-password).
HOST, PORT = os.environ.get("HOST", "127.0.0.1"), int(os.environ.get("PORT", 8765))
PASSWORD_FILE = DATA_DIR / "app_password.json"   # salted PBKDF2 hash, never the password itself
LOCAL_ONLY_PATHS = {"/api/sql", "/api/google/key"}   # only from the computer running the app
# View-only mode, for sharing the log through a tunnel: set VIEWER_PASSWORD to require a password,
# allow only reading the log, benchmarks and changes, and refuse every change.
VIEWER_PASSWORD = os.environ.get("VIEWER_PASSWORD", "")
VIEWER_PATHS = {"/", "/index.html", "/api/log", "/api/benchmarks", "/api/changes", "/api/checks"}
VIEWER_HTML = """<style>
  body.viewer nav button[data-tab="upload"], body.viewer nav button[data-tab="google"],
  body.viewer nav button[data-tab="sql"], body.viewer a[href="/api/export.csv"],
  body.viewer #bm-panel .row:has(#bm-save), body.viewer #bm-help, body.viewer #bm-msg,
  body.viewer #bm-panel .hint, body.viewer .bm-table tr > :last-child, body.viewer .chk-act, body.viewer .help-edit { display: none; }
  .viewer-note { margin: 0 0 8px; font-size: 13px; color: var(--muted, #666); }
</style>
<script>
  document.addEventListener("DOMContentLoaded", () => {
    document.querySelector('nav').insertAdjacentHTML("beforebegin", '<p class="viewer-note">View-only link. Nothing can be changed from here.</p>');
    document.querySelector('nav button[data-tab="log"]').click();
  });
</script>
"""
MAX_UPLOAD = 25 * 1024 * 1024
MAX_SQL_ROWS = 5000
GOOGLE_DIR = DATA_DIR / "google"          # service account key (git-ignored with the rest of data/)
SYNC_EVERY = int(os.environ.get("SYNC_MINUTES", 15)) * 60

COLUMNS = [
    "id", "office", "period_from", "period_thru",
    "total_collection", "insurance_collection", "new_patients", "np_first_visit",
    "existing_patient_referrals", "total_production",
    "source_file", "logged_at",
]

DATA_COLUMNS = COLUMNS[1:]  # everything but id

SCHEMA = """
-- One row per office per month. Reports are month-to-date (the 1st through some
-- day), so a later upload for the same month replaces the earlier one.
CREATE TABLE IF NOT EXISTS office_summary (
    id                          INTEGER PRIMARY KEY,
    office                      TEXT NOT NULL,
    period_from                 TEXT NOT NULL,   -- YYYY-MM-01
    period_thru                 TEXT NOT NULL,   -- YYYY-MM-DD, same month
    total_collection            REAL NOT NULL DEFAULT 0, -- Collection total, as a positive amount
    insurance_collection        REAL NOT NULL,   -- Collection > Insurance, as a positive amount
    new_patients                INTEGER NOT NULL,-- np_first_visit - existing_patient_referrals
    np_first_visit              INTEGER NOT NULL,-- New Patients by First Visit Date
    existing_patient_referrals  INTEGER NOT NULL,-- Referrals > EXISTING PATIENT (new patients)
    total_production            REAL NOT NULL,   -- Production > Total
    source_file                 TEXT,
    logged_at                   TEXT NOT NULL,
    UNIQUE (office, period_from)
);

-- Audit trail: every row that was replaced by a newer upload or deleted.
CREATE TABLE IF NOT EXISTS office_history (
    history_id                  INTEGER PRIMARY KEY,
    action                      TEXT NOT NULL,   -- 'replaced' or 'deleted'
    action_at                   TEXT NOT NULL,
    summary_id                  INTEGER,
    office                      TEXT NOT NULL,
    period_from                 TEXT NOT NULL,
    period_thru                 TEXT NOT NULL,
    total_collection            REAL,
    insurance_collection        REAL,
    new_patients                INTEGER,
    np_first_visit              INTEGER,
    existing_patient_referrals  INTEGER,
    total_production            REAL,
    source_file                 TEXT,
    logged_at                   TEXT
);

-- Numbers pasted from spreadsheets (e.g. refunds). One value per office, month
-- and metric; each has its own as-of date since sheets update on other days.
CREATE TABLE IF NOT EXISTS office_metrics (
    id           INTEGER PRIMARY KEY,
    office       TEXT NOT NULL,
    period_from  TEXT NOT NULL,   -- YYYY-MM-01
    as_of        TEXT NOT NULL,   -- YYYY-MM-DD the sheet was updated
    metric       TEXT NOT NULL,   -- e.g. 'refunds'
    value        REAL NOT NULL,
    source       TEXT,
    logged_at    TEXT NOT NULL,
    detail       TEXT,            -- JSON: what was counted (sheet rows, lenders); never patient names
    UNIQUE (office, period_from, metric)
);

CREATE TABLE IF NOT EXISTS metric_history (
    history_id   INTEGER PRIMARY KEY,
    action       TEXT NOT NULL,   -- 'replaced' or 'deleted'
    action_at    TEXT NOT NULL,
    metric_id    INTEGER,
    office       TEXT NOT NULL,
    period_from  TEXT NOT NULL,
    as_of        TEXT NOT NULL,
    metric       TEXT NOT NULL,
    value        REAL,
    source       TEXT,
    logged_at    TEXT,
    detail       TEXT,
    replaced_by  TEXT             -- source of the value that replaced this one (for undoing a sheet)
);

-- Office Google Sheets (New Patient Analysis) read automatically, read-only.
CREATE TABLE IF NOT EXISTS sheet_sources (
    id              INTEGER PRIMARY KEY,
    spreadsheet_id  TEXT NOT NULL UNIQUE,
    url             TEXT,
    title           TEXT,
    office          TEXT,             -- as found on the sheet ("Office: ...")
    added_at        TEXT NOT NULL,
    last_sync_at    TEXT,
    last_status     TEXT              -- "ok: ..." or the error
);

-- Every change to a logged number: from a sync, an upload or a paste.
CREATE TABLE IF NOT EXISTS change_log (
    id           INTEGER PRIMARY KEY,
    changed_at   TEXT NOT NULL,
    office       TEXT NOT NULL,
    period_from  TEXT NOT NULL,
    as_of        TEXT,
    metric       TEXT NOT NULL,
    old_value    REAL,                -- NULL when the number is new
    new_value    REAL,
    via          TEXT,                -- "Google Sheets sync", "upload", "paste"
    source       TEXT
);

-- Data checks: flags someone checked and marked OK (hidden while the value stays the same),
-- and when the office was emailed about a flag.
CREATE TABLE IF NOT EXISTS check_ok (
    office       TEXT NOT NULL,
    period_from  TEXT NOT NULL,
    check_key    TEXT NOT NULL,       -- 'rate:<column>' or 'row:<sheet row>:<Prime|Subprime>:<lender>'
    value_sig    TEXT NOT NULL,       -- the flagged value(s); a changed value is flagged again
    note         TEXT NOT NULL,
    marked_at    TEXT NOT NULL,
    PRIMARY KEY (office, period_from, check_key)
);
CREATE TABLE IF NOT EXISTS check_notified (
    office       TEXT NOT NULL,
    period_from  TEXT NOT NULL,
    check_key    TEXT NOT NULL,
    value_sig    TEXT NOT NULL,
    notified_at  TEXT NOT NULL,
    PRIMARY KEY (office, period_from, check_key)
);
-- Who to email about an office's sheet.
CREATE TABLE IF NOT EXISTS office_contacts (
    office      TEXT PRIMARY KEY,
    emails      TEXT NOT NULL,        -- comma-separated
    updated_at  TEXT NOT NULL
);

-- Sheet office names mapped to the office names used in the log.
CREATE TABLE IF NOT EXISTS office_alias (
    alias   TEXT PRIMARY KEY,     -- normalized sheet name, e.g. 'gardena 2'
    office  TEXT NOT NULL
);
"""

# Metrics kept in office_metrics: columns pasted from sheets, plus numbers from other reports.
REPORT_METRICS = [
    "np_collection",      # Referral Production Listing: office total collection
    "ortho_collection",   # Daily Journal (Patient Type: Ortho): summary Total Payments
    "ortho_referrals",    # Treatment Plan Status (proc code ZREFORTH): count on "Total for office"
    "medi_cal_patients",  # New Patient Analysis workbook: rows with Insurance Type Medi-Cal
    "hmo_patients",       # New Patient Analysis workbook: rows with Insurance Type HMO
    "no_ssn_itin_patients",  # New Patient Analysis workbook: Notes say no SSN / no ITIN
    "financing_applicants",  # New Patient Analysis workbook: applied with any prime or subprime lender
    "existing_finance_patients",  # New Patient Analysis workbook: Notes say they already have a financing account
    "prime_applicants",   # New Patient Analysis workbook: applied with a prime lender (Ran = Yes)
    "prime_approved",     # New Patient Analysis workbook: approved by a prime lender ($ approved or collected)
    "prime_apps_ran",     # New Patient Analysis workbook: prime applications (each lender with Ran = Yes)
    "prime_apps_approved",  # New Patient Analysis workbook: prime applications approved ($ approved or collected)
    "prime_approved_amount",   # New Patient Analysis workbook: $ approved by prime lenders
    "prime_collected_amount",  # New Patient Analysis workbook: $ collected through prime lenders
    "prime_funded",       # New Patient Analysis workbook: patients with money collected through a prime lender
    "prime_approved_funded",  # New Patient Analysis workbook: approved prime patients who were also funded
    "subprime_applicants",  # New Patient Analysis workbook: a lender named under "Subprime Co. Ran"
    "subprime_approved",    # New Patient Analysis workbook: applied and approved by a subprime lender
    "subprime_apps_ran",    # New Patient Analysis workbook: subprime applications (each lender run)
    "subprime_apps_approved",  # New Patient Analysis workbook: subprime applications approved (each lender named)
    "subprime_approved_amount",   # New Patient Analysis workbook: $ approved by subprime lenders
    "subprime_collected_amount",  # New Patient Analysis workbook: $ collected through subprime lenders
    "subprime_funded",      # New Patient Analysis workbook: patients with money collected through a subprime lender
    "fin_approved",         # prime + subprime: patients approved by either (each patient once)
    "fin_funded",           # prime + subprime: patients funded by either (each patient once)
    "fin_approved_funded",  # prime + subprime: approved patients who were also funded
    "np_rows",              # New Patient Analysis workbook: new patient rows (denominator for per-NP numbers)
    "tx_diagnosed_amount",  # New Patient Analysis workbook: sum of "$ Treatment Diagnosed"
    "np_collection_wb",     # New Patient Analysis workbook: sum of Denticon "Collection" for these patients
    "no_next_visit",        # New Patient Analysis workbook: "Has an Appointment" = No
    "started_tx",           # New Patient Analysis workbook: paid anything (Denticon, financing or other)
    "collected_over_approved",  # data check: patients with more collected than approved (prime or subprime)
]
METRIC_NAMES = list(METRICS) + REPORT_METRICS
COUNT_METRICS = {"ortho_referrals", "medi_cal_patients", "hmo_patients", "no_ssn_itin_patients",
                 "financing_applicants", "existing_finance_patients", "prime_applicants", "prime_approved",
                 "prime_apps_ran", "prime_apps_approved", "prime_funded", "prime_approved_funded",
                 "subprime_applicants", "subprime_approved", "subprime_apps_ran", "subprime_apps_approved",
                 "subprime_funded", "fin_approved", "fin_funded", "fin_approved_funded",
                 "np_rows", "no_next_visit", "started_tx", "collected_over_approved"}
DETAIL_METRICS = ["medi_cal_patients", "hmo_patients", "no_ssn_itin_patients", "financing_applicants",
                  "existing_finance_patients", "prime_applicants", "prime_approved", "prime_apps_ran",
                  "prime_apps_approved", "prime_approved_amount", "prime_collected_amount", "prime_funded",
                  "prime_approved_funded", "subprime_applicants", "subprime_approved",
                  "subprime_apps_ran", "subprime_apps_approved", "subprime_approved_amount",
                  "subprime_collected_amount", "subprime_funded", "fin_approved", "fin_funded",
                  "fin_approved_funded", "tx_diagnosed_amount", "np_collection_wb", "no_next_visit", "started_tx",
                  "collected_over_approved"]   # keep a list of what was counted, shown on hover   # whole numbers, not money

# Calculated in the monthly_log view; NULL when an input is missing.
CALCULATED = {
    "net_collections": "ROUND(s.total_collection - m.refunds, 2)",
    "refund_rate": "ROUND(m.refunds / NULLIF(s.total_collection, 0), 4)",   # fraction, e.g. 0.0045
    "existing_pt_collection": "ROUND(s.total_collection - m.np_collection, 2)",
    # (Medi-Cal + HMO + no SSN/ITIN patients) / Executive Summary new patients, as a fraction
    "irrelevant_rate": "ROUND(1.0 * (m.medi_cal_patients + m.hmo_patients + m.no_ssn_itin_patients)"
                       " / NULLIF(s.new_patients, 0), 4)",
    # ortho referrals (Treatment Plan Status) / Executive Summary new patients, as a fraction
    "ortho_rate": "ROUND(1.0 * m.ortho_referrals / NULLIF(s.new_patients, 0), 4)",
    # financing applicants (New Patient Analysis) / Executive Summary new patients, as a fraction
    "financing_rate": "ROUND(1.0 * m.financing_applicants / NULLIF(s.new_patients, 0), 4)",
    # prime approved / prime applicants (both from the New Patient Analysis), as a fraction
    "prime_approval_rate": "ROUND(1.0 * m.prime_approved / NULLIF(m.prime_applicants, 0), 4)",
    # prime $ collected / prime $ approved, as a fraction
    "prime_collected_rate": "ROUND(m.prime_collected_amount / NULLIF(m.prime_approved_amount, 0), 4)",
    # approved prime patients who were funded / approved prime patients, as a fraction
    "prime_funded_rate": "ROUND(1.0 * m.prime_approved_funded / NULLIF(m.prime_approved, 0), 4)",
    # subprime approved / subprime applicants, as a fraction
    "subprime_approval_rate": "ROUND(1.0 * m.subprime_approved / NULLIF(m.subprime_applicants, 0), 4)",
    # subprime $ collected / subprime $ approved, as a fraction
    "subprime_collected_rate": "ROUND(m.subprime_collected_amount / NULLIF(m.subprime_approved_amount, 0), 4)",
    # prime + subprime combined
    "fin_approval_rate": "ROUND(1.0 * m.fin_approved / NULLIF(m.financing_applicants, 0), 4)",
    "fin_apps_ran": "m.prime_apps_ran + m.subprime_apps_ran",
    "fin_apps_approved": "m.prime_apps_approved + m.subprime_apps_approved",
    "fin_app_approval_rate": "ROUND(1.0 * (m.prime_apps_approved + m.subprime_apps_approved)"
                             " / NULLIF(m.prime_apps_ran + m.subprime_apps_ran, 0), 4)",
    "fin_approved_amount": "ROUND(m.prime_approved_amount + m.subprime_approved_amount, 2)",
    "fin_collected_amount": "ROUND(m.prime_collected_amount + m.subprime_collected_amount, 2)",
    "fin_collected_rate": "ROUND((m.prime_collected_amount + m.subprime_collected_amount)"
                          " / NULLIF(m.prime_approved_amount + m.subprime_approved_amount, 0), 4)",
    "fin_funded_rate": "ROUND(1.0 * m.fin_approved_funded / NULLIF(m.fin_approved, 0), 4)",
    # treatment (all from the New Patient Analysis; per NP = per new-patient row in the workbook)
    "avg_tx_dx_per_np": "ROUND(m.tx_diagnosed_amount / NULLIF(m.np_rows, 0), 2)",
    "collection_per_np": "ROUND(m.np_collection_wb / NULLIF(m.np_rows, 0), 2)",
    "collection_per_dx": "ROUND(m.np_collection_wb / NULLIF(m.tx_diagnosed_amount, 0), 4)",
    "np_loss_rate": "ROUND(1.0 * m.no_next_visit / NULLIF(m.np_rows, 0), 4)",
    "start_tx_rate": "ROUND(1.0 * m.started_tx / NULLIF(m.np_rows, 0), 4)",
    # average prime $ approved per approved NP / approved app / app ran / NP who applied
    "avg_prime_per_approved_np": "ROUND(m.prime_approved_amount / NULLIF(m.prime_approved, 0), 2)",
    "avg_prime_per_approved_app": "ROUND(m.prime_approved_amount / NULLIF(m.prime_apps_approved, 0), 2)",
    "avg_prime_per_app_ran": "ROUND(m.prime_approved_amount / NULLIF(m.prime_apps_ran, 0), 2)",
    "avg_prime_per_applicant": "ROUND(m.prime_approved_amount / NULLIF(m.prime_applicants, 0), 2)",
}
LOG_COLUMNS = [
    "office", "period_from", "period_thru",
    "total_collection", "insurance_collection", "new_patients", "np_first_visit",
    "existing_patient_referrals", "total_production",
    *[c for m in METRIC_NAMES for c in (m, f"{m}_as_of")],
    *CALCULATED,
    "source_file", "logged_at",
]


def log_view_sql():
    """monthly_log: one row per office per month, Executive Summary plus sheet metrics."""
    def value(m):
        v = f"MAX(CASE WHEN metric = '{m}' THEN value END)"
        return f"CAST({v} AS INTEGER)" if m in COUNT_METRICS else v
    pivots = ",\n".join(
        f"{value(m)} AS {m}, "
        f"MAX(CASE WHEN metric = '{m}' THEN as_of END) AS {m}_as_of"
        for m in METRIC_NAMES)
    metric_cols = ", ".join(f"m.{m}, m.{m}_as_of" for m in METRIC_NAMES)
    metric_cols += "".join(f", m.{m}_detail" for m in DETAIL_METRICS)
    pivots += "".join(f",\nMAX(CASE WHEN metric = '{m}' THEN detail END) AS {m}_detail" for m in DETAIL_METRICS)
    calculated = ",\n       ".join(f"{expr} AS {name}" for name, expr in CALCULATED.items())
    return f"""
DROP VIEW IF EXISTS monthly_log;
CREATE VIEW monthly_log AS
WITH m AS (
    SELECT office, period_from, MAX(as_of) AS latest_as_of, {pivots}
    FROM office_metrics GROUP BY office, period_from
)
SELECT COALESCE(s.office, m.office) AS office,
       COALESCE(s.period_from, m.period_from) AS period_from,
       COALESCE(s.period_thru, m.latest_as_of) AS period_thru,
       s.total_collection, s.insurance_collection, s.new_patients, s.np_first_visit,
       s.existing_patient_referrals, s.total_production,
       {metric_cols},
       {calculated},
       s.source_file, s.logged_at
FROM office_summary s
FULL OUTER JOIN m ON m.office = s.office AND m.period_from = s.period_from;
"""


def archive(conn, where, params, action):
    """Copy matching office_summary rows into office_history."""
    cols = ", ".join(DATA_COLUMNS)
    conn.execute(
        f"INSERT INTO office_history (action, action_at, summary_id, {cols}) "
        f"SELECT ?, ?, id, {cols} FROM office_summary WHERE {where}",
        (action, datetime.now().isoformat(timespec="seconds"), *params),
    )


BENCHMARK_SQL = """
CREATE TABLE benchmarks (
    metric      TEXT PRIMARY KEY,   -- a Log column, e.g. ortho_rate
    direction   TEXT NOT NULL,      -- 'higher' (bigger is better) or 'lower'
    green       REAL NOT NULL,      -- at or beyond this: green
    yellow      REAL NOT NULL,      -- at or beyond this (but not green): yellow; otherwise red
    updated_at  TEXT NOT NULL
);
-- Starting benchmark: % to ortho is good at 25%+, yellow from 20%, red below.
INSERT INTO benchmarks VALUES ('ortho_rate', 'higher', 0.25, 0.20, datetime('now'));
"""


def migrate(conn):
    cols = {r["name"] for r in conn.execute("PRAGMA table_info(office_summary)")}
    if cols and "total_collection" not in cols:  # logs created before this column existed
        conn.execute("ALTER TABLE office_summary ADD COLUMN total_collection REAL NOT NULL DEFAULT 0")
    table_sql = conn.execute(
        "SELECT sql FROM sqlite_master WHERE type = 'table' AND name = 'office_summary'"
    ).fetchone()
    if table_sql and "UNIQUE (office, period_from, period_thru)" in table_sql[0]:
        # Old logs allowed several rows per office per month. Keep the latest,
        # archive the rest, and rebuild the table with the one-per-month rule.
        with conn:
            conn.execute("ALTER TABLE office_summary RENAME TO office_summary_old")
            conn.executescript(SCHEMA)
            cols = ", ".join(COLUMNS)
            conn.execute(f"""
                INSERT INTO office_summary ({cols})
                SELECT {cols} FROM (
                    SELECT *, ROW_NUMBER() OVER (
                        PARTITION BY office, period_from ORDER BY period_thru DESC, logged_at DESC) AS rn
                    FROM office_summary_old) WHERE rn = 1""")
            data_cols = ", ".join(DATA_COLUMNS)
            conn.execute(f"""
                INSERT INTO office_history (action, action_at, summary_id, {data_cols})
                SELECT 'replaced', ?, id, {data_cols} FROM office_summary_old
                WHERE id NOT IN (SELECT id FROM office_summary)""",
                (datetime.now().isoformat(timespec="seconds"),))
            conn.execute("DROP TABLE office_summary_old")


_setup_lock = threading.Lock()
_setup_done = False


def db():
    """A connection to the log. The schema, upgrades and the monthly_log view are set up once
    per run: rebuilding the view on every request let two requests at once collide."""
    global _setup_done
    DATA_DIR.mkdir(exist_ok=True)
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    with _setup_lock:
        if not _setup_done:
            _setup(conn)
            _setup_done = True
    return conn


def _setup(conn):
    migrate(conn)
    conn.executescript(SCHEMA)
    if not conn.execute("SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'benchmarks'").fetchone():
        conn.executescript(BENCHMARK_SQL)
    for table in ("office_metrics", "metric_history"):   # logs created before details were kept
        if "detail" not in {r["name"] for r in conn.execute(f"PRAGMA table_info({table})")}:
            conn.execute(f"ALTER TABLE {table} ADD COLUMN detail TEXT")
    if "kind" not in {r["name"] for r in conn.execute("PRAGMA table_info(sheet_sources)")}:
        # 'np' = an office's New Patient Analysis; 'refunds' = the % TO GOAL sheet (refunds for every office)
        conn.execute("ALTER TABLE sheet_sources ADD COLUMN kind TEXT NOT NULL DEFAULT 'np'")
    if "replaced_by" not in {r["name"] for r in conn.execute("PRAGMA table_info(metric_history)")}:
        conn.execute("ALTER TABLE metric_history ADD COLUMN replaced_by TEXT")
    conn.executescript(log_view_sql())


def write_csv(conn):
    rows = conn.execute(
        f"SELECT {', '.join(LOG_COLUMNS)} FROM monthly_log ORDER BY period_from, office"
    ).fetchall()
    tmp = CSV_PATH.with_suffix(".csv.tmp")
    with open(tmp, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(LOG_COLUMNS)
        w.writerows([tuple(r) for r in rows])
    tmp.replace(CSV_PATH)


def to_iso(date_str):
    """Accept M/D/YYYY (from the report) or YYYY-MM-DD (from the form)."""
    s = (date_str or "").strip()
    for fmt in ("%Y-%m-%d", "%m/%d/%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    raise ValueError(f"Invalid date: {date_str!r}")


# Month-to-date numbers only grow, so a later report with a lower number (or a same-date report with different
# numbers) is shown to the person saving first; they confirm with "Save anyway". Not for the Google sync.
NOT_CUMULATIVE = {"no_next_visit", "collected_over_approved"}   # these can go down as the month goes on


class NeedsReview(Exception):
    def __init__(self, items):
        self.items = items


def check_not_future(day, label="Thru"):
    if day > date.today().isoformat():
        raise ValueError(f"The {label} date {day} hasn't happened yet. Use the report's last day (today or earlier).")


def review_item(review, office, metric, was, was_as_of, now, as_of):
    """Note a number to confirm: lower than what the log has, or changed in a same-date report."""
    if was is None or now is None or metric in NOT_CUMULATIVE or round(was, 2) == round(now, 2):
        return
    if now < was - 0.005 or was_as_of == as_of:
        review.append({"office": office, "metric": metric, "was": round(was, 2), "now": round(now, 2),
                       "was_as_of": was_as_of, "lower": now < was - 0.005})


def save_rows(payload):
    """Insert or replace one row per office for the month. Returns counts."""
    period_from = to_iso(payload.get("period_from"))
    period_thru = to_iso(payload.get("period_thru"))
    if not period_from.endswith("-01"):
        raise ValueError("Reports are month-to-date: 'From' must be the 1st of the month")
    if period_thru[:7] != period_from[:7]:
        raise ValueError("'Thru' must be in the same month as 'From'")
    if period_from > period_thru:
        raise ValueError("'From' date is after 'Thru' date")
    check_not_future(period_thru)
    offices = payload.get("offices") or []
    if not offices:
        raise ValueError("Nothing to save")

    now = datetime.now().isoformat(timespec="seconds")
    counts = {"added": 0, "replaced": 0, "skipped": []}
    review = []
    conn = db()
    try:
      with conn:
        for o in offices:
            office = str(o["office"]).strip()
            existing = conn.execute(
                "SELECT * FROM office_summary WHERE office = ? AND period_from = ?",
                (office, period_from),
            ).fetchone()
            if existing and existing["period_thru"] > period_thru:
                counts["skipped"].append(f"{office} (already have thru {existing['period_thru']})")
                continue
            if existing and not payload.get("confirm"):
                new = {"total_collection": o["total_collection"], "insurance_collection": o["insurance_collection"],
                       "np_first_visit": o["np_first_visit"], "existing_patient_referrals": o["existing_patient_referrals"],
                       "new_patients": int(o["np_first_visit"]) - int(o["existing_patient_referrals"]),
                       "total_production": o["total_production"]}
                for k, v in new.items():
                    review_item(review, office, k, existing[k], existing["period_thru"], float(v), period_thru)
            npfv = int(o["np_first_visit"])
            existing_refs = int(o["existing_patient_referrals"])
            values = (office, period_from, period_thru,
                      round(float(o["total_collection"]), 2), round(float(o["insurance_collection"]), 2),
                      npfv - existing_refs, npfv, existing_refs, round(float(o["total_production"]), 2),
                      payload.get("source_file"), now)
            if existing:
                archive(conn, "id = ?", (existing["id"],), "replaced")
                conn.execute(
                    f"UPDATE office_summary SET {', '.join(c + ' = ?' for c in DATA_COLUMNS)} WHERE id = ?",
                    (*values, existing["id"]),
                )
                counts["replaced"] += 1
            else:
                conn.execute(
                    f"INSERT INTO office_summary ({', '.join(DATA_COLUMNS)}) "
                    f"VALUES ({', '.join('?' for _ in DATA_COLUMNS)})",
                    values,
                )
                counts["added"] += 1
        if review:
            raise NeedsReview(review)   # undoes everything above; nothing is saved until confirmed
    except NeedsReview as e:
        conn.close()
        return {"needs_review": True, "review": e.items, "skipped": counts["skipped"]}
    write_csv(conn)
    conn.close()
    return counts


def known_offices(conn):
    return sorted({r[0] for r in conn.execute(
        "SELECT office FROM office_summary UNION SELECT office FROM office_metrics "
        "UNION SELECT office FROM office_alias")})


def read_paste(text):
    """Parse pasted sheet rows and suggest a log office for each."""
    result = parse_paste(text, office_in_first_column=True)   # the % TO GOAL header over the names is just "."
    conn = db()
    known = known_offices(conn)
    aliases = dict(conn.execute("SELECT alias, office FROM office_alias").fetchall())
    conn.close()
    for row in result["rows"]:
        remembered = aliases.get(normalize(row["sheet_office"]))
        cands = candidates(row["sheet_office"], known)
        if remembered:
            row["match"], row["how"] = remembered, "remembered"
        elif len(cands) == 1:
            row["match"], row["how"] = cands[0], "auto"
        else:
            row["match"], row["how"] = None, "ambiguous" if cands else "none"
        row["candidates"] = cands
        row["new_name"] = title_case(row["sheet_office"])
    result["known_offices"] = known
    return result


def save_metrics(payload):
    """Save pasted sheet values; latest as-of date wins for each office/month/metric."""
    as_of = to_iso(payload.get("as_of"))
    period_from = as_of[:8] + "01"
    if payload.get("period_from") and to_iso(payload["period_from"]) != period_from:
        raise ValueError("Reports are month-to-date: 'From' must be the 1st of the 'Thru' month")
    from_report = bool(payload.get("from_report"))  # office names already match the log
    check_not_future(as_of, "as-of" if not from_report else "Thru")
    reviewing = not payload.get("confirm") and payload.get("via") != "Google Sheets sync"
    rows = payload.get("rows") or []
    if not rows:
        raise ValueError("Nothing to save")
    now = datetime.now().isoformat(timespec="seconds")
    counts = {"added": 0, "replaced": 0, "unchanged": 0, "skipped": []}
    review = []
    conn = db()
    try:
      with conn:
        for row in rows:
            office = str(row.get("office") or "").strip()
            if not office:
                raise ValueError(f"Choose which office '{row.get('sheet_office')}' is")
            if not from_report:
                conn.execute("INSERT OR REPLACE INTO office_alias (alias, office) VALUES (?, ?)",
                             (normalize(row["sheet_office"]), office))
            for metric, value in (row.get("values") or {}).items():
                if metric not in METRIC_NAMES or value is None:
                    continue
                detail = (row.get("details") or {}).get(metric)
                detail = json.dumps(detail, separators=(",", ":")) if detail is not None else None
                existing = conn.execute(
                    "SELECT id, as_of, value, detail FROM office_metrics "
                    "WHERE office = ? AND period_from = ? AND metric = ?",
                    (office, period_from, metric)).fetchone()
                if existing and existing["as_of"] > as_of:
                    counts["skipped"].append(f"{office} {metric} (already have as of {existing['as_of']})")
                    continue
                if existing and existing["as_of"] == as_of and existing["value"] == round(float(value), 2) \
                        and existing["detail"] == detail:
                    counts["unchanged"] += 1
                    continue
                if existing and reviewing:
                    review_item(review, office, metric, existing["value"], existing["as_of"], float(value), as_of)
                values = (office, period_from, as_of, metric, round(float(value), 2),
                          payload.get("source"), now, detail)
                old_value = existing["value"] if existing else None
                if old_value is None or round(old_value, 2) != round(float(value), 2):
                    conn.execute(
                        "INSERT INTO change_log (changed_at, office, period_from, as_of, metric, old_value, new_value, "
                        "via, source) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                        (now, office, period_from, as_of, metric, old_value, round(float(value), 2),
                         payload.get("via") or ("upload" if from_report else "paste"), payload.get("source")))
                    counts.setdefault("changes", []).append(metric)
                if existing:
                    archive_metrics(conn, "id = ?", (existing["id"],), "replaced", payload.get("source"))
                    conn.execute(
                        "UPDATE office_metrics SET office = ?, period_from = ?, as_of = ?, metric = ?, "
                        "value = ?, source = ?, logged_at = ?, detail = ? WHERE id = ?", (*values, existing["id"]))
                    counts["replaced"] += 1
                else:
                    conn.execute(
                        "INSERT INTO office_metrics (office, period_from, as_of, metric, value, source, logged_at, detail) "
                        "VALUES (?, ?, ?, ?, ?, ?, ?, ?)", values)
                    counts["added"] += 1
        if review:
            raise NeedsReview(review)   # undoes everything above; nothing is saved until confirmed
    except NeedsReview as e:
        conn.close()
        return {"needs_review": True, "review": e.items, "skipped": counts["skipped"]}
    write_csv(conn)
    conn.close()
    return counts


def archive_metrics(conn, where, params, action, replaced_by=None):
    conn.execute(
        "INSERT INTO metric_history (action, action_at, metric_id, office, period_from, as_of, metric, "
        "value, source, logged_at, detail, replaced_by) SELECT ?, ?, id, office, period_from, as_of, metric, value, "
        f"source, logged_at, detail, ? FROM office_metrics WHERE {where}",
        (action, datetime.now().isoformat(timespec="seconds"), replaced_by, *params))


def delete_rows(keys):
    """Delete log rows (office + month), archiving everything removed."""
    pairs = [(str(k["office"]), str(k["period_from"])) for k in keys]
    if not pairs:
        raise ValueError("Nothing selected")
    conn = db()
    n = 0
    with conn:
        for office, period_from in pairs:
            archive(conn, "office = ? AND period_from = ?", (office, period_from), "deleted")
            archive_metrics(conn, "office = ? AND period_from = ?", (office, period_from), "deleted")
            a = conn.execute("DELETE FROM office_summary WHERE office = ? AND period_from = ?",
                             (office, period_from)).rowcount
            b = conn.execute("DELETE FROM office_metrics WHERE office = ? AND period_from = ?",
                             (office, period_from)).rowcount
            n += 1 if (a or b) else 0
    write_csv(conn)
    conn.close()
    return n


# ------------------------------------------------------------------ Google Sheets sync

account = google_sheets.ServiceAccount(GOOGLE_DIR)
sync_lock = threading.Lock()
sync_state = {"last_run": None, "last_result": None, "next_run": None, "running": False}


def refund_tabs(wb, how_many=2, since=None):
    """The newest month tabs of a % TO GOAL sheet that have refunds for this month: [(tab, as_of, rows)].
    A tab's numbers count for its own month: "Updated 10/1/26" on the September tab is September's final
    numbers (as of 9/30). A tab not updated since its month began (a copied template) is skipped."""
    found = []
    for month in np_analysis.dated_months(wb):
        first = date(month["year"], month["month"], 1)
        last = (first.replace(day=28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)
        if first > date.today():
            continue
        if since and first.isoformat()[:7] < since:
            break   # tabs are newest first
        try:
            parsed = parse_paste(rows_to_text(wb.rows(month["sheet"])), office_in_first_column=True)
        except PasteError:
            continue
        if parsed["as_of"] and parsed["as_of"] < first.isoformat():
            continue
        as_of = min(parsed["as_of"] or date.today().isoformat(), last.isoformat(), date.today().isoformat())
        found.append((month["sheet"], as_of, parsed["rows"]))
        if how_many and len(found) == how_many:
            break
    return found


def add_source(url):
    sid = google_sheets.spreadsheet_id(url)
    wb = google_sheets.SheetsWorkbook(account, sid)          # checks access right away
    try:
        # The office name comes from the newest month tab that has patients (a new month's tab is often still empty).
        latest = np_analysis.months_with_data(wb, 1)[0]
        kind, office, tab = "np", latest["offices"][0]["office"], latest["sheet"]
    except np_analysis.NPAnalysisError as np_error:
        tabs = refund_tabs(wb, 1)     # not a New Patient Analysis: maybe the % TO GOAL sheet (refunds)
        if not tabs:
            raise ValueError(f"{np_error}. It isn't a % TO GOAL sheet either (no month tab with a 'refunds' column).")
        kind, office, tab = "refunds", None, tabs[0][0]
    conn = db()
    with conn:
        conn.execute("INSERT INTO sheet_sources (spreadsheet_id, url, title, office, added_at, kind) VALUES (?, ?, ?, ?, ?, ?) "
                     "ON CONFLICT (spreadsheet_id) DO UPDATE SET url = excluded.url, title = excluded.title, "
                     "office = excluded.office, kind = excluded.kind",
                     (sid, url, wb.title, office, datetime.now().isoformat(timespec="seconds"), kind))
    conn.close()
    return {"title": wb.title, "office": office, "kind": kind, "latest_tab": tab}


def sync_refunds(wb, result, since=None):
    """Save refunds from a % TO GOAL sheet's newest two month tabs, matching office names like a paste."""
    conn = db()
    aliases = {r["alias"]: r["office"] for r in conn.execute("SELECT alias, office FROM office_alias")}
    known = {normalize(o): o for o in known_offices(conn)}
    conn.close()
    unmatched = set()
    for tab, as_of, rows in refund_tabs(wb, None if since else 2, since):
        matched = []
        for row in rows:
            office = aliases.get(normalize(row["sheet_office"])) or known.get(normalize(row["sheet_office"]))
            if office:
                matched.append({"sheet_office": row["sheet_office"], "office": office, "values": row["values"]})
            else:
                unmatched.add(row["sheet_office"])
        saved = save_metrics({"as_of": as_of, "source": f"Google Sheet: {wb.title} / {tab}",
                              "via": "Google Sheets sync", "rows": matched}) if matched else {"skipped": []}
        n = len(saved.get("changes", []))
        result["changes"] += n
        result["tabs"].append({"tab": f"{tab} (as of {as_of[5:].replace('-', '/')})", "changes": n,
                               "skipped": len(saved["skipped"])})
    if not result["tabs"]:
        raise ValueError("No month tab with a 'refunds' column was updated for its month yet")
    return sorted(unmatched)


def month_thru(parsed, month):
    """Thru date: the latest date on the tab, else today (this month) or the month's last day."""
    if parsed["period"].get("thru"):
        return min(parsed["period"]["thru"], date.today().isoformat())
    first = date(month["year"], month["month"], 1)
    last = (first.replace(day=28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)
    return min(date.today(), last).isoformat()


def sync_all(since=None):
    """Read the latest and previous month tab of every office sheet and save any changes.
    since="2026-01": load past months instead -- every month tab from that month on (the "Load past months" button)."""
    if not sync_lock.acquire(blocking=False):
        return {"error": "A sync is already running"}
    sync_state["running"] = True
    results = []
    try:
        conn = db()
        sources = [dict(r) for r in conn.execute("SELECT * FROM sheet_sources ORDER BY office")]
        conn.close()
        for src in sources:
            result = {"id": src["id"], "office": src["office"], "title": src["title"], "tabs": [], "changes": 0}
            try:
                wb = google_sheets.SheetsWorkbook(account, src["spreadsheet_id"])
                read, office_seen, unmatched = 0, None, []
                if src.get("kind") == "refunds":
                    unmatched = sync_refunds(wb, result, since)
                for month in ([] if src.get("kind") == "refunds" else np_analysis.dated_months(wb)):
                    # latest + previous month with patients (or every month since `since`)
                    if since and f"{month['year']}-{month['month']:02d}" < since:
                        break
                    if read == 2 and not since:
                        break
                    if date(month["year"], month["month"], 1) > date.today():
                        continue   # a tab set up for a month that hasn't started
                    try:
                        # A tab pasted without the "Office: ..." line uses the office this sheet was added for.
                        parsed = np_analysis.parse_month(wb, month, office_seen or src["office"])
                    except np_analysis.EmptyMonthError:
                        result["tabs"].append({"tab": month["sheet"], "empty": True})
                        continue
                    except np_analysis.NPAnalysisError as e:
                        result["tabs"].append({"tab": month["sheet"], "error": str(e)})
                        read += 1
                        continue
                    read += 1
                    office_seen = office_seen or parsed["offices"][0]["office"]
                    o = parsed["offices"][0]
                    saved = save_metrics({
                        "as_of": month_thru(parsed, month), "period_from": parsed["period"]["from"],
                        "source": f"Google Sheet: {wb.title} / {month['sheet']}", "from_report": True,
                        "via": "Google Sheets sync",
                        "rows": [{"sheet_office": o["office"], "office": o["office"],
                                  "values": {k: v for k, v in o.items() if k not in ("office", "details")},
                                  "details": o["details"]}]})
                    n = len(saved.get("changes", []))
                    result["changes"] += n
                    result["tabs"].append({"tab": month["sheet"], "changes": n, "skipped": len(saved["skipped"])})
                status = "ok: " + ", ".join(f"{t['tab']} " + ("empty, skipped" if t.get("empty") else
                                            f"error ({t['error']})" if "error" in t else f"{t['changes']} changed")
                                            for t in result["tabs"])
                if unmatched:
                    status += f"; not matched (paste once on the Upload tab to match): {', '.join(unmatched)}"
            except (google_sheets.GoogleError, np_analysis.NPAnalysisError, ValueError, PasteError) as e:
                result["error"] = str(e)
                status = f"error: {e}"
            conn = db()
            with conn:
                conn.execute("UPDATE sheet_sources SET last_sync_at = ?, last_status = ? WHERE id = ?",
                             (datetime.now().isoformat(timespec="seconds"), status, src["id"]))
            conn.close()
            results.append(result)
    finally:
        sync_state.update(running=False, last_run=datetime.now().isoformat(timespec="seconds"),
                          last_result=results)
        sync_lock.release()
    return {"results": results}


def sync_loop():
    """Sync every SYNC_EVERY seconds while the app runs (only once a key and sheets are set up)."""
    while True:
        sync_state["next_run"] = datetime.fromtimestamp(time.time() + SYNC_EVERY).isoformat(timespec="seconds")
        time.sleep(SYNC_EVERY)
        try:
            conn = db()
            has_sources = conn.execute("SELECT COUNT(*) FROM sheet_sources").fetchone()[0]
            conn.close()
            if account.exists() and has_sources:
                print(f"{datetime.now():%H:%M} Google Sheets sync:", sync_all())
        except Exception as e:  # keep the loop alive
            print("Google Sheets sync failed:", e)


def remove_source(source_id):
    """Stop syncing a sheet and undo the numbers it brought in: each number currently from this
    sheet goes back to its value from before the sheet (from history), or is removed if it didn't
    exist before. Numbers that have come from somewhere else since are left alone."""
    if not sync_lock.acquire(timeout=120):          # don't undo while a sync is writing
        raise ValueError("A sync is running; try again in a minute")
    try:
        conn = db()
        src = conn.execute("SELECT * FROM sheet_sources WHERE id = ?", (source_id,)).fetchone()
        if not src:
            raise ValueError("That sheet isn't in the list any more")
        prefix = "Google Sheet: " + (src["title"] or "").replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_") + " /%"
        mine = "source LIKE ? ESCAPE '\\'"
        now = datetime.now().isoformat(timespec="seconds")
        counts = {"restored": 0, "removed": 0, "changed": 0}
        with conn:
            conn.execute("DELETE FROM sheet_sources WHERE id = ?", (source_id,))
            for r in conn.execute(f"SELECT * FROM office_metrics WHERE {mine}", (prefix,)).fetchall():
                # The value this sheet replaced: the last history entry that this sheet took over from.
                # (Matched by office/month/number -- row ids can be reused after a delete.)
                prev = conn.execute("SELECT * FROM metric_history WHERE office = ? AND period_from = ? AND metric = ? "
                                    "AND replaced_by LIKE ? ESCAPE '\\' AND (source IS NULL OR NOT source LIKE ? ESCAPE '\\') "
                                    "ORDER BY history_id DESC LIMIT 1",
                                    (r["office"], r["period_from"], r["metric"], prefix, prefix)).fetchone()
                archive_metrics(conn, "id = ?", (r["id"],), "replaced" if prev else "deleted")
                if prev:
                    conn.execute("UPDATE office_metrics SET as_of = ?, value = ?, source = ?, logged_at = ?, detail = ? "
                                 "WHERE id = ?", (prev["as_of"], prev["value"], prev["source"], prev["logged_at"],
                                                  prev["detail"], r["id"]))
                    counts["restored"] += 1
                else:
                    conn.execute("DELETE FROM office_metrics WHERE id = ?", (r["id"],))
                    counts["removed"] += 1
                new = prev["value"] if prev else None
                if new != r["value"]:
                    counts["changed"] += 1
                    conn.execute("INSERT INTO change_log (changed_at, office, period_from, as_of, metric, old_value, "
                                 "new_value, via, source) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                                 (now, r["office"], r["period_from"], prev["as_of"] if prev else None, r["metric"],
                                  r["value"], new, "sheet removed", f"undid Google Sheet: {src['title']}"))
        write_csv(conn)
        conn.close()
        return {"title": src["title"], "office": src["office"], **counts}
    finally:
        sync_lock.release()


def save_benchmark(b):
    metric = str(b.get("metric") or "")
    if metric not in LOG_COLUMNS or metric in ("office", "period_from", "period_thru", "source_file", "logged_at") \
            or metric.endswith("_as_of"):
        raise ValueError("Pick a number column for the benchmark")
    direction = b.get("direction")
    if direction not in ("higher", "lower"):
        raise ValueError("Choose whether higher or lower is better")
    try:
        green, yellow = float(b["green"]), float(b["yellow"])
    except (KeyError, TypeError, ValueError):
        raise ValueError("Enter both the green and yellow cut-offs as numbers")
    if direction == "higher" and yellow > green or direction == "lower" and yellow < green:
        raise ValueError("The yellow cut-off must be " + ("below" if direction == "higher" else "above") +
                         " the green one")
    conn = db()
    with conn:
        conn.execute("INSERT INTO benchmarks VALUES (?, ?, ?, ?, ?) ON CONFLICT (metric) DO UPDATE SET "
                     "direction = excluded.direction, green = excluded.green, yellow = excluded.yellow, "
                     "updated_at = excluded.updated_at",
                     (metric, direction, green, yellow, datetime.now().isoformat(timespec="seconds")))
    conn.close()
    return {"ok": True}


EMAIL_RE = re.compile(r"[^@\s,;]+@[^@\s,;]+\.[A-Za-z]{2,}")


def _check_ref(c):
    office, period_from = str(c.get("office") or "").strip(), str(c.get("period_from") or "")
    key, sig = str(c.get("check_key") or ""), str(c.get("value_sig") or "")
    if not office or not re.fullmatch(r"\d{4}-\d{2}-01", period_from) or not key.startswith(("rate:", "row:")) or not sig:
        raise ValueError("That check is missing its office, month or value")
    return office, period_from, key, sig


def _log_check(conn, office, period_from, key, via, note=None):
    conn.execute("INSERT INTO change_log (changed_at, office, period_from, metric, via, source) VALUES (?, ?, ?, ?, ?, ?)",
                 (datetime.now().isoformat(timespec="seconds"), office, period_from, f"check: {key}", via, note))


def checks_info(viewer=False):
    """Marked-OK flags, emailed flags, office emails and where each office/month's sheet numbers came from."""
    conn = db()
    ok = [dict(r) for r in conn.execute("SELECT * FROM check_ok")]
    if viewer:
        conn.close()
        return {"ok": ok}
    notified = [dict(r) for r in conn.execute("SELECT * FROM check_notified")]
    contacts = {r["office"]: r["emails"] for r in conn.execute("SELECT * FROM office_contacts")}
    urls = {r["office"]: r["url"] for r in conn.execute("SELECT office, url FROM sheet_sources WHERE office IS NOT NULL")}
    sources = {}
    for r in conn.execute("SELECT office, period_from, source FROM office_metrics WHERE metric = 'np_rows'"):
        src = r["source"] or ""
        google = src.startswith("Google Sheet:")
        sources[f"{r['office']}|{r['period_from']}"] = {
            "tab": src.rsplit(" / ", 1)[1] if google and " / " in src else None,
            "file": None if google else src or None,
            "url": urls.get(r["office"]) if google else None}
    conn.close()
    return {"ok": ok, "notified": notified, "contacts": contacts, "sources": sources}


def mark_check_ok(c):
    office, period_from, key, sig = _check_ref(c)
    note = str(c.get("note") or "").strip()
    if not note:
        raise ValueError("Add a short note saying why the value is OK")
    conn = db()
    with conn:
        conn.execute("INSERT OR REPLACE INTO check_ok VALUES (?, ?, ?, ?, ?, ?)",
                     (office, period_from, key, sig, note[:500], datetime.now().isoformat(timespec="seconds")))
        _log_check(conn, office, period_from, key, "marked OK", note[:500])
    conn.close()
    return {"ok": True}


def unmark_check_ok(c):
    office, period_from, key = str(c.get("office") or ""), str(c.get("period_from") or ""), str(c.get("check_key") or "")
    conn = db()
    with conn:
        if conn.execute("DELETE FROM check_ok WHERE office = ? AND period_from = ? AND check_key = ?",
                        (office, period_from, key)).rowcount:
            _log_check(conn, office, period_from, key, "unmarked OK")
    conn.close()
    return {"ok": True}


def mark_notified(c):
    checks = [_check_ref(x) for x in c.get("checks") or []]
    if not checks:
        raise ValueError("No checks to record")
    now = datetime.now().isoformat(timespec="seconds")
    conn = db()
    with conn:
        for office, period_from, key, sig in checks:
            conn.execute("INSERT OR REPLACE INTO check_notified VALUES (?, ?, ?, ?, ?)", (office, period_from, key, sig, now))
            _log_check(conn, office, period_from, key, "email drafted", c.get("to"))
    conn.close()
    return {"ok": True}


def save_contacts(c):
    office = str(c.get("office") or "").strip()
    emails = EMAIL_RE.findall(str(c.get("emails") or ""))
    if not office:
        raise ValueError("Missing the office")
    if not emails and str(c.get("emails") or "").strip():
        raise ValueError("Enter email addresses like name@example.com, separated by commas")
    conn = db()
    with conn:
        if emails:
            conn.execute("INSERT OR REPLACE INTO office_contacts VALUES (?, ?, ?)",
                         (office, ", ".join(dict.fromkeys(emails)), datetime.now().isoformat(timespec="seconds")))
        else:
            conn.execute("DELETE FROM office_contacts WHERE office = ?", (office,))
        conn.execute("INSERT INTO change_log (changed_at, office, period_from, metric, via, source) VALUES (?, ?, '', ?, ?, ?)",
                     (datetime.now().isoformat(timespec="seconds"), office, "office emails", "contacts", ", ".join(emails) or "removed"))
    conn.close()
    return {"ok": True, "emails": ", ".join(dict.fromkeys(emails))}


def _pw_hash(password, salt, iterations):
    return hashlib.pbkdf2_hmac("sha256", password.encode(), bytes.fromhex(salt), iterations).hex()


def load_app_password():
    try:
        return json.loads(PASSWORD_FILE.read_text())
    except FileNotFoundError:
        return None


APP_PASSWORD = load_app_password()
_good_auth = set()   # sha256 of Authorization headers already checked, so each request isn't re-hashed


def app_password_ok(header, given):
    key = hashlib.sha256(header.encode()).hexdigest()
    if key in _good_auth:
        return True
    rec = APP_PASSWORD
    if rec and hmac.compare_digest(_pw_hash(given, rec["salt"], rec["iterations"]), rec["hash"]):
        _good_auth.add(key)
        return True
    return False


def set_app_password():
    pw = getpass.getpass("New password for the app (12+ characters): ")
    if len(pw) < 12:
        raise SystemExit("Too short: use at least 12 characters.")
    if getpass.getpass("Type it again: ") != pw:
        raise SystemExit("The passwords don't match.")
    salt, iterations = secrets.token_hex(16), 300_000
    DATA_DIR.mkdir(exist_ok=True)
    fd = os.open(PASSWORD_FILE, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(fd, "w") as f:
        json.dump({"salt": salt, "iterations": iterations, "hash": _pw_hash(pw, salt, iterations)}, f)
    print(f"Saved. Restart the app for the new password to take effect.")


def google_status():
    conn = db()
    sources = [dict(r) for r in conn.execute("SELECT id, url, title, office, kind, added_at, last_sync_at, last_status "
                                             "FROM sheet_sources ORDER BY office")]
    conn.close()
    return {"key_present": account.exists(), "service_account": account.email(), "sources": sources,
            "every_minutes": SYNC_EVERY // 60, **{k: v for k, v in sync_state.items() if k != "last_result"}}


def run_sql(query):
    """Run a read-only query against the log."""
    db().close()  # make sure the file and table exist
    conn = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True)
    try:
        conn.execute("PRAGMA query_only = ON")
        cur = conn.execute(query)
        cols = [d[0] for d in cur.description] if cur.description else []
        rows = cur.fetchmany(MAX_SQL_ROWS + 1)
    finally:
        conn.close()
    return {
        "columns": cols,
        "rows": [list(r) for r in rows[:MAX_SQL_ROWS]],
        "truncated": len(rows) > MAX_SQL_ROWS,
    }


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        print(f"{self.command} {self.path} -> {args[1] if len(args) > 1 else ''}")

    def _send(self, status, body, ctype="application/json", extra=None):
        data = body if isinstance(body, bytes) else json.dumps(body).encode()
        self.send_response(status)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(data)

    def _body(self):
        length = int(self.headers.get("Content-Length") or 0)
        if length > MAX_UPLOAD:
            raise ValueError("File is too large (25 MB max)")
        return self.rfile.read(length)

    def _blocked(self):
        """Ask for the password (view-only link, or the app password on the office network), and refuse
        what this request may not do. True if refused."""
        auth = self.headers.get("Authorization", "")
        try:
            given = base64.b64decode(auth[6:]).decode().partition(":")[2] if auth.startswith("Basic ") else ""
        except ValueError:
            given = ""
        if VIEWER_PASSWORD:
            if not hmac.compare_digest(given.encode(), VIEWER_PASSWORD.encode()):
                time.sleep(1)   # slow down guessing
                self._send(401, {"error": "Password required"},
                           extra={"WWW-Authenticate": 'Basic realm="Monthly report (view only)", charset="UTF-8"'})
                return True
            if self.command != "GET" or self.path not in VIEWER_PATHS:
                self._send(403, {"error": "This is a view-only link."})
                return True
        elif APP_PASSWORD and not app_password_ok(auth, given):
            time.sleep(1)
            self._send(401, {"error": "Password required"},
                       extra={"WWW-Authenticate": 'Basic realm="Monthly report", charset="UTF-8"'})
            return True
        if self.path in LOCAL_ONLY_PATHS and self.client_address[0] not in ("127.0.0.1", "::1", "::ffff:127.0.0.1"):
            self._send(403, {"error": "For safety, this only works on the computer that runs the app."})
            return True
        return False

    def end_headers(self):
        if VIEWER_PASSWORD or APP_PASSWORD:
            for k, v in (("Cache-Control", "no-store"), ("X-Robots-Tag", "noindex, nofollow"),
                         ("X-Frame-Options", "DENY"), ("Referrer-Policy", "no-referrer")):
                self.send_header(k, v)
        super().end_headers()

    def do_GET(self):
        if self._blocked():
            return
        if self.path in ("/", "/index.html"):
            page = (ROOT / "static" / "index.html").read_bytes()
            if VIEWER_PASSWORD:
                page = page.replace(b"</head>", VIEWER_HTML.encode() + b"</head>", 1).replace(
                    b'<body class="hide-select">', b'<body class="hide-select viewer">', 1)
            self._send(200, page, "text/html; charset=utf-8", {} if VIEWER_PASSWORD else {"Cache-Control": "no-store"})
        elif self.path == "/api/log":
            conn = db()
            rows = conn.execute(
                f"SELECT {', '.join(LOG_COLUMNS + [m + '_detail' for m in DETAIL_METRICS])} "
                "FROM monthly_log ORDER BY period_thru DESC, office"
            ).fetchall()
            conn.close()
            self._send(200, [dict(r) for r in rows])
        elif self.path == "/api/benchmarks":
            conn = db()
            rows = [dict(r) for r in conn.execute("SELECT * FROM benchmarks ORDER BY metric")]
            conn.close()
            self._send(200, rows)
        elif self.path == "/api/checks":
            self._send(200, checks_info(viewer=bool(VIEWER_PASSWORD)))
        elif self.path == "/api/google/status":
            self._send(200, google_status())
        elif self.path == "/api/changes":
            conn = db()
            rows = conn.execute("SELECT * FROM change_log ORDER BY id DESC LIMIT 2000").fetchall()
            conn.close()
            self._send(200, [dict(r) for r in rows])
        elif self.path == "/api/export.csv":
            conn = db()
            write_csv(conn)
            conn.close()
            self._send(200, CSV_PATH.read_bytes(), "text/csv",
                       {"Content-Disposition": 'attachment; filename="office_summary.csv"'})
        else:
            self._send(404, {"error": "Not found"})

    def do_POST(self):
        if self._blocked():
            return
        try:
            if self.path == "/api/parse":
                filename = unquote(self.headers.get("X-Filename", "upload"))
                suffix = Path(filename).suffix.lower()
                with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
                    tmp.write(self._body())
                try:
                    result = parse_report(tmp.name, filename)
                finally:
                    os.unlink(tmp.name)
                for key in ("from", "thru"):
                    if key in result["period"]:
                        result["period"][key] = to_iso(result["period"][key])
                result["source_file"] = filename
                self._send(200, result)
            elif self.path == "/api/save":
                self._send(200, save_rows(json.loads(self._body())))
            elif self.path == "/api/parse-paste":
                self._send(200, read_paste(json.loads(self._body()).get("text", "")))
            elif self.path == "/api/save-metrics":
                self._send(200, save_metrics(json.loads(self._body())))
            elif self.path == "/api/delete":
                self._send(200, {"deleted": delete_rows(json.loads(self._body())["keys"])})
            elif self.path == "/api/benchmarks":
                self._send(200, save_benchmark(json.loads(self._body())))
            elif self.path == "/api/benchmarks/delete":
                conn = db()
                with conn:
                    conn.execute("DELETE FROM benchmarks WHERE metric = ?", (json.loads(self._body()).get("metric"),))
                conn.close()
                self._send(200, {"ok": True})
            elif self.path == "/api/checks/ok":
                self._send(200, mark_check_ok(json.loads(self._body())))
            elif self.path == "/api/checks/ok/delete":
                self._send(200, unmark_check_ok(json.loads(self._body())))
            elif self.path == "/api/checks/notified":
                self._send(200, mark_notified(json.loads(self._body())))
            elif self.path == "/api/office-contacts":
                self._send(200, save_contacts(json.loads(self._body())))
            elif self.path == "/api/google/key":
                self._send(200, {"service_account": account.save_key(self._body().decode("utf-8", "replace"))})
            elif self.path == "/api/google/sources":
                self._send(200, add_source(json.loads(self._body()).get("url", "")))
            elif self.path == "/api/google/sources/delete":
                self._send(200, remove_source(int(json.loads(self._body())["id"])))
            elif self.path == "/api/google/sync":
                since = str(json.loads(self._body() or b"{}").get("since") or "")
                if since and not re.fullmatch(r"20\d\d-(0[1-9]|1[0-2])", since):
                    raise ValueError("Pick a start month like 2026-01")
                self._send(200, sync_all(since or None))
            elif self.path == "/api/sql":
                query = json.loads(self._body()).get("query", "").strip()
                if not query:
                    raise ValueError("Enter a query")
                self._send(200, run_sql(query))
            else:
                self._send(404, {"error": "Not found"})
        except (ParseError, PasteError, ValueError, KeyError, sqlite3.Error,
                google_sheets.GoogleError, np_analysis.NPAnalysisError) as e:
            self._send(400, {"error": str(e)})
        except Exception as e:  # keep the server up, report the problem
            self._send(500, {"error": f"{type(e).__name__}: {e}"})


if __name__ == "__main__":
    if sys.argv[1:] == ["--set-password"]:
        set_app_password()
        raise SystemExit
    if HOST not in ("127.0.0.1", "localhost", "::1") and not (APP_PASSWORD or VIEWER_PASSWORD):
        raise SystemExit("Opening the app to the network needs a password first: python3 app.py --set-password")
    db().close()
    if VIEWER_PASSWORD:
        print("View-only mode: password required, no changes allowed, Google sync off.")
    else:
        threading.Thread(target=sync_loop, daemon=True).start()
    print(f"Executive Summary logger running at http://{HOST}:{PORT}  (Ctrl+C to stop)")
    if HOST == "0.0.0.0":
        print(f"On the office network: http://{socket.gethostname()}.local:{PORT}  (password required)")
    print(f"Log: {DB_PATH}  |  CSV: {CSV_PATH}")
    try:
        ThreadingHTTPServer((HOST, PORT), Handler).serve_forever()
    except KeyboardInterrupt:
        pass
