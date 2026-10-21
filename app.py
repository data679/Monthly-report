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
import subprocess
import sys
import tempfile
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from urllib.parse import unquote

import threading
import time
from datetime import date, timedelta

import denticon_api
import invoca_api
import assistant
import google_sheets
import np_analysis
from sheet_paste import METRICS, PasteError, candidates, normalize, parse_paste, rows_to_text, title_case
from summary_parser import parse_new_patient_appts, parse_new_patients_list, calls_upload_kind, ParseError, parse_patient_types, parse_report

ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
DB_PATH = DATA_DIR / "reports.db"
CSV_PATH = DATA_DIR / "office_summary.csv"
# HOST=0.0.0.0 opens the app to the office network; that needs a password (python3 app.py --set-password).
HOST, PORT = os.environ.get("HOST", "127.0.0.1"), int(os.environ.get("PORT", 8765))
PASSWORD_FILE = DATA_DIR / "app_password.json"   # salted PBKDF2 hash, never the password itself
# Only from the computer running the app: SQL, the Google key, and the Denticon audit (admin view).
LOCAL_ONLY_PATHS = {"/api/sql", "/api/assistant/status", "/api/assistant/key", "/api/assistant/key/delete", "/api/assistant/test", "/api/assistant/cache/clear", "/api/inbox", "/api/robot", "/api/robot/pull", "/api/google/key", "/api/audit", "/api/save-lender-journal", "/api/denticon/status",
                    "/api/denticon/key", "/api/denticon/key/delete", "/api/denticon/test", "/api/denticon/offices",
                    "/api/invoca/status", "/api/invoca/save", "/api/invoca/delete", "/api/invoca/test",
                    "/api/google/notes-sheet", "/api/google/notes-sheet/delete", "/api/google/notes-sheet/write",
                    "/api/maps/patient-types", "/api/maps/office-location", "/api/maps/zip-download", "/api/maps/shapes-download"}
# View-only mode, for sharing the log through a tunnel: set VIEWER_PASSWORD to require a password,
# allow only reading the log, benchmarks and changes, and refuse every change.
VIEWER_PASSWORD = os.environ.get("VIEWER_PASSWORD", "")
VIEWER_PATHS = {"/", "/index.html", "/api/log", "/api/log/days", "/api/days", "/api/benchmarks", "/api/changes", "/api/checks", "/api/maps/trends"}
VIEWER_HTML = """<style>
  body.viewer nav button[data-tab="upload"], body.viewer nav button[data-tab="google"],
  body.viewer nav button[data-tab="sql"], body.viewer nav button[data-tab="calls"], body.viewer a[href="/api/export.csv"],
  body.viewer #bm-panel .row:has(#bm-save), body.viewer #bm-help, body.viewer #bm-msg,
  body.viewer #bm-panel .hint, body.viewer .bm-table tr > :last-child, body.viewer .chk-act, body.viewer .help-edit, body.viewer nav button[data-tab="audit"],
  body.viewer #mp-section, body.viewer #mp-admin { display: none; }
  .viewer-note { margin: 0 0 8px; font-size: 13px; color: var(--muted, #666); }
</style>
<script>
  document.addEventListener("DOMContentLoaded", () => {
    document.querySelector('nav').insertAdjacentHTML("beforebegin", '<p class="viewer-note">View-only link. Nothing can be changed from here.</p>');
    document.querySelector('nav button[data-tab="log"]').click();
  });
</script>
"""
MAX_UPLOAD = 200 * 1024 * 1024   # a month of Daily Journal with services can be large
MAX_SQL_ROWS = 5000
GOOGLE_DIR = DATA_DIR / "google"          # service account key (git-ignored with the rest of data/)
SYNC_EVERY = int(os.environ.get("SYNC_MINUTES", 15)) * 60
BACKUP_DIR = DATA_DIR / "backups"         # daily copies of the log (git-ignored with the rest of data/)
KEEP_BACKUPS = int(os.environ.get("KEEP_BACKUPS", 30))

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

-- Day-by-day history: Executive Summary reports run for a single day (From = Thru). Any date range is the sum
-- of its days. Uploading the same day again replaces that day. Separate from the monthly rows above.
CREATE TABLE IF NOT EXISTS office_daily (
    id                          INTEGER PRIMARY KEY,
    office                      TEXT NOT NULL,
    day                         TEXT NOT NULL,   -- YYYY-MM-DD, the one day the report was run for
    total_collection            REAL NOT NULL,
    insurance_collection        REAL NOT NULL,
    new_patients                INTEGER NOT NULL,
    np_first_visit              INTEGER NOT NULL,
    existing_patient_referrals  INTEGER NOT NULL,
    total_production            REAL NOT NULL,
    source_file                 TEXT,
    logged_at                   TEXT NOT NULL,
    UNIQUE (office, day)
);

-- Every day a one-day Executive Summary was imported for, even one with no activity (a Sunday): an office with no
-- row on such a day was closed (zero), while a day missing here was never pulled.
CREATE TABLE IF NOT EXISTS daily_imports (
    day         TEXT PRIMARY KEY,                -- YYYY-MM-DD
    offices     INTEGER NOT NULL,                -- offices with activity in that report
    source      TEXT,
    logged_at   TEXT NOT NULL
);

-- New Patient Analysis counts by the day each sheet row is listed under (from the sheet's date lines), so the
-- sheet columns can be added up for any run of days. Replaced whenever that office's month tab is read again.
CREATE TABLE IF NOT EXISTS np_daily (
    office       TEXT NOT NULL,
    period_from  TEXT NOT NULL,                  -- the sheet month, YYYY-MM-01
    day          TEXT NOT NULL,                  -- YYYY-MM-DD ('' for rows above the first date line)
    metric       TEXT NOT NULL,
    value        REAL NOT NULL,
    as_of        TEXT NOT NULL,
    PRIMARY KEY (office, period_from, day, metric)
);

-- The assistant's questions and answers (no patient data: its tools only return totals), for review.
CREATE TABLE IF NOT EXISTS assistant_log (
    id           INTEGER PRIMARY KEY,
    asked_at     TEXT NOT NULL,
    who          TEXT NOT NULL,                  -- this computer | viewer link | other computer
    question     TEXT NOT NULL,
    answer       TEXT,
    tools        TEXT,
    input_tokens INTEGER, output_tokens INTEGER, cache_read_tokens INTEGER,
    error        TEXT,
    route        TEXT                            -- how it was answered: cache | rules | haiku | main (see ai_chat)
);

-- The assistant's saved verdicts and answers, so a repeated question costs nothing. An answer is reused only while the
-- data (and the date) are unchanged; an off-topic verdict is reused for good. Cleared from the Audit card.
CREATE TABLE IF NOT EXISTS assistant_cache (
    qkey      TEXT PRIMARY KEY,                  -- the question, lower-cased with punctuation and extra spaces removed
    verdict   TEXT NOT NULL,                     -- on | off
    answer    TEXT,
    data_fp   TEXT,
    saved_at  TEXT NOT NULL,
    hits      INTEGER NOT NULL DEFAULT 0
);

-- Every day each other one-day report was imported for (an empty day too), so a month's coverage can be checked per report.
CREATE TABLE IF NOT EXISTS daily_report_imports (
    day         TEXT NOT NULL,
    report      TEXT NOT NULL,                   -- referral | ortho_journal | ortho_referrals
    logged_at   TEXT NOT NULL,
    PRIMARY KEY (day, report)
);

-- The other Denticon reports run for a single day: NP collection, ortho collection, ortho referrals.
CREATE TABLE IF NOT EXISTS daily_metrics (
    id          INTEGER PRIMARY KEY,
    office      TEXT NOT NULL,
    day         TEXT NOT NULL,                   -- YYYY-MM-DD
    metric      TEXT NOT NULL,
    value       REAL NOT NULL,
    source      TEXT,
    logged_at   TEXT NOT NULL,
    UNIQUE (office, day, metric)
);

-- Files picked up from data/inbox (dropped there by a person or the report robot) and what happened to them.
CREATE TABLE IF NOT EXISTS inbox_log (
    id           INTEGER PRIMARY KEY,
    seen_at      TEXT NOT NULL,
    file         TEXT NOT NULL,
    status       TEXT NOT NULL,                  -- imported | review | failed
    report_type  TEXT,
    period_from  TEXT,
    period_thru  TEXT,
    note         TEXT
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

-- Denticon check of the NP audit sheets. Pat IDs are never stored: pk is a keyed hash (data/patient_key).
-- Each new patient row's collected amounts per lender, as the sheet has them (replaced on every save).
CREATE TABLE IF NOT EXISTS np_patients (
    office       TEXT NOT NULL,
    period_from  TEXT NOT NULL,
    row          INTEGER NOT NULL,
    pk           TEXT NOT NULL,
    prime        TEXT NOT NULL,       -- JSON {"Alphaeon": 3620.0, ...} ($Amt Coll, before the lender's fee)
    subprime     REAL NOT NULL,       -- Subprime $Amt Coll
    other        REAL NOT NULL DEFAULT 0,  -- "Other form of payment" on the row (cash etc., may be in $Amt Coll)
    PRIMARY KEY (office, period_from, row)
);
-- Lender payments and merchant fees per patient from a Daily Journal (Patient Type: Both), per office and month.
CREATE TABLE IF NOT EXISTS lender_payments (
    office       TEXT NOT NULL,
    period_from  TEXT NOT NULL,
    pk           TEXT NOT NULL,
    lender       TEXT NOT NULL,
    payment      REAL NOT NULL,
    fee          REAL NOT NULL
);
CREATE TABLE IF NOT EXISTS journal_uploads (
    office       TEXT NOT NULL,
    period_from  TEXT NOT NULL,
    as_of        TEXT NOT NULL,       -- the journal's Thru date
    source       TEXT,
    logged_at    TEXT NOT NULL,
    PRIMARY KEY (office, period_from)
);

-- Denticon API offices (from Practices > Offices) and which log office each one is.
CREATE TABLE IF NOT EXISTS denticon_offices (
    office_id    INTEGER PRIMARY KEY,
    name         TEXT,
    log_office   TEXT,                -- NULL until matched
    updated_at   TEXT NOT NULL
);

-- Calls pulled from Invoca (one row per call; no caller phone numbers or names).
CREATE TABLE IF NOT EXISTS invoca_calls (
    transaction_id   TEXT PRIMARY KEY,
    complete_call_id TEXT,
    call_date        TEXT NOT NULL,    -- YYYY-MM-DD (local)
    start_local      TEXT,
    start_utc        INTEGER,
    duration         INTEGER,          -- seconds ringing + talking
    connect_duration INTEGER,          -- seconds connected to the office
    ivr_duration     INTEGER,
    advertiser_id    TEXT,
    advertiser       TEXT,
    campaign         TEXT,
    promo_line       TEXT,
    virtual_line_id  TEXT,
    city             TEXT,
    region           TEXT,
    mobile           TEXT,
    repeat_caller    INTEGER,
    result           TEXT,
    signals          TEXT,             -- JSON list of signal names that were true
    custom           TEXT,             -- JSON {marketing data field: value}
    pulled_at        TEXT NOT NULL,
    caller_fp        TEXT              -- keyed fingerprint of the caller's number (the number itself is never kept)
);
CREATE INDEX IF NOT EXISTS invoca_calls_date ON invoca_calls (call_date);
CREATE INDEX IF NOT EXISTS invoca_calls_fp ON invoca_calls (caller_fp);
CREATE INDEX IF NOT EXISTS map_patients_pat ON map_patients (pat_id, office, period_from);

-- Invoca promo lines (tracking numbers / named lines) -> the office they belong to, typed in once on the Calls tab.
CREATE TABLE IF NOT EXISTS promo_lines (
    line        TEXT PRIMARY KEY,      -- the promo line as Invoca sends it (a description, or the number itself)
    office      TEXT NOT NULL DEFAULT '',
    label       TEXT NOT NULL DEFAULT '',
    updated_at  TEXT NOT NULL
);

-- Patients' phone numbers from Denticon's "New Patients List" (Pat ID + first visit), as keyed fingerprints only:
-- lets a caller be linked to the patient record the app already holds (first visit, $ collected).
CREATE TABLE IF NOT EXISTS patient_phones (
    fp          TEXT NOT NULL,
    pat_id      TEXT NOT NULL,
    office      TEXT NOT NULL,
    first_visit TEXT,
    created_on  TEXT,
    phone_kind  TEXT NOT NULL,         -- cell / home / work
    loaded_at   TEXT NOT NULL,
    PRIMARY KEY (fp, pat_id, office, phone_kind)
);
CREATE INDEX IF NOT EXISTS patient_phones_fp ON patient_phones (fp);

-- New-patient appointments from Denticon's "New Patients Through Appointment Book": only a keyed fingerprint of
-- each phone number, the office and the appointment date, to tell which callers became new patients.
CREATE TABLE IF NOT EXISTS np_appts (
    fp          TEXT NOT NULL,
    office      TEXT NOT NULL,
    appt_date   TEXT,
    phone_kind  TEXT NOT NULL,         -- cell / home / work
    loaded_at   TEXT NOT NULL,
    PRIMARY KEY (fp, office, appt_date, phone_kind)
);
CREATE INDEX IF NOT EXISTS np_appts_fp ON np_appts (fp);

-- The notes sheet: a Google Sheet the app keeps a copy of the Log in (one way), with a Notes column people type in.
CREATE TABLE IF NOT EXISTS notes_sheet (
    id                INTEGER PRIMARY KEY CHECK (id = 1),
    spreadsheet_id    TEXT NOT NULL,
    url               TEXT,
    title             TEXT,
    last_write_at     TEXT,
    last_status       TEXT,
    last_fingerprint  TEXT
);

-- Maps tab. New patients by ZIP from the Referral Production Listing (by Zip Code): Pat ID and amounts only,
-- never names. Replaced per office and month by each upload (a later Thru date wins).
CREATE TABLE IF NOT EXISTS map_patients (
    office         TEXT NOT NULL,
    period_from    TEXT NOT NULL,
    as_of          TEXT NOT NULL,
    pat_id         TEXT NOT NULL,
    zip            TEXT NOT NULL,
    referral_type  TEXT,
    visits         INTEGER,
    production     REAL,
    collection     REAL NOT NULL
);
-- Pat ID -> General / Ortho, from Denticon's Patient List - Address (names and addresses aren't kept).
CREATE TABLE IF NOT EXISTS patient_types (
    office      TEXT NOT NULL,
    pat_id      TEXT NOT NULL,
    type        TEXT NOT NULL,
    updated_at  TEXT NOT NULL,
    PRIMARY KEY (office, pat_id)
);
CREATE TABLE IF NOT EXISTS office_locations (
    office   TEXT PRIMARY KEY,
    address  TEXT,
    lat      REAL NOT NULL,
    lon      REAL NOT NULL
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
    if "other" not in {r["name"] for r in conn.execute("PRAGMA table_info(np_patients)")}:
        conn.execute("ALTER TABLE np_patients ADD COLUMN other REAL NOT NULL DEFAULT 0")
    if "replaced_by" not in {r["name"] for r in conn.execute("PRAGMA table_info(metric_history)")}:
        conn.execute("ALTER TABLE metric_history ADD COLUMN replaced_by TEXT")
    if "route" not in {r["name"] for r in conn.execute("PRAGMA table_info(assistant_log)")}:
        conn.execute("ALTER TABLE assistant_log ADD COLUMN route TEXT")
    if "caller_fp" not in {r["name"] for r in conn.execute("PRAGMA table_info(invoca_calls)")}:
        conn.execute("ALTER TABLE invoca_calls ADD COLUMN caller_fp TEXT")
    if not conn.execute("SELECT 1 FROM daily_report_imports LIMIT 1").fetchone():
        for kind, pat in (("referral", "%ReferralRefProdList 20__-__-__.xlsx"), ("ortho_journal", "%OrthoDailyJournal 20__-__-__.doc"),
                          ("ortho_referrals", "%TreatPlanStatus 20__-__-__.xlsx")):
            conn.execute("INSERT OR IGNORE INTO daily_report_imports (day, report, logged_at) SELECT period_from, ?, MAX(seen_at) "
                         "FROM inbox_log WHERE status = 'imported' AND file LIKE ? AND period_from = period_thru GROUP BY period_from", (kind, pat))
    # Days imported before daily_imports existed: every day with office rows counts as imported.
    conn.execute("INSERT OR IGNORE INTO daily_imports (day, offices, source, logged_at) "
                 "SELECT day, COUNT(*), MAX(source_file), MAX(logged_at) FROM office_daily GROUP BY day")
    conn.executescript(log_view_sql())


def log_rows(conn, extra=()):
    """The Log as the app shows it: the monthly_log view, with months pulled day by day filled from the days."""
    rows = [dict(r) for r in conn.execute(f"SELECT {', '.join(LOG_COLUMNS + list(extra))} FROM monthly_log")]
    return month_rows_from_days(rows)


def write_csv(conn):
    rows = sorted(log_rows(conn), key=lambda r: (r["period_from"], r["office"]))
    tmp = CSV_PATH.with_suffix(".csv.tmp")
    with open(tmp, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(LOG_COLUMNS)
        w.writerows([[r.get(c) for c in LOG_COLUMNS] for r in rows])
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
            map_rows = (row.get("details") or {}).get("_map")
            if map_rows is not None:
                newer = conn.execute("SELECT 1 FROM map_patients WHERE office = ? AND period_from = ? AND as_of > ? LIMIT 1",
                                     (office, period_from, as_of)).fetchone()
                if not newer:
                    conn.execute("DELETE FROM map_patients WHERE office = ? AND period_from = ?", (office, period_from))
                    conn.executemany(
                        "INSERT INTO map_patients VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                        [(office, period_from, as_of, str(p["pat_id"]), str(p["zip"])[:5], p.get("referral_type") or "",
                          int(p.get("visits") or 0), float(p.get("production") or 0), float(p["collection"]))
                         for p in map_rows if str(p.get("zip", "")).isdigit()])
            by_row = (row.get("details") or {}).get("_by_row")
            if by_row is not None:
                newer = conn.execute("SELECT 1 FROM np_daily WHERE office = ? AND period_from = ? AND as_of > ? LIMIT 1",
                                     (office, period_from, as_of)).fetchone()
                if not newer:
                    # A row without a date line above it takes the patient's first visit from the uploaded Denticon
                    # New Patients List (matched by Pat ID, or by its key when only the key came), else the 1st.
                    first = {}
                    for r_ in conn.execute("SELECT pat_id, MIN(first_visit) AS fv FROM patient_phones WHERE office = ? AND "
                                           "first_visit >= ? AND first_visit < date(?, '+2 months') GROUP BY pat_id",
                                           (office, (date.fromisoformat(period_from) - timedelta(31)).isoformat(), period_from)):
                        first[str(r_["pat_id"])] = r_["fv"]; first[pat_key(str(r_["pat_id"]))] = r_["fv"]
                    who = {str(p_.get("row")): str(p_.get("pat_id") or p_.get("pk") or "") for p_ in (row.get("details") or {}).get("_patients") or []}
                    merged = {}
                    for rn, info in by_row.items():
                        d = info.get("day") or first.get(who.get(str(rn), ""), "")   # '' = no date found (not in any list)
                        for m, v in (info.get("m") or {}).items():
                            if m in METRIC_NAMES and v:
                                merged[(d, m)] = merged.get((d, m), 0) + float(v)
                    conn.execute("DELETE FROM np_daily WHERE office = ? AND period_from = ?", (office, period_from))
                    conn.executemany("INSERT INTO np_daily VALUES (?, ?, ?, ?, ?, ?)",
                                     [(office, period_from, d, m, v, as_of) for (d, m), v in merged.items() if v])
            patients = (row.get("details") or {}).get("_patients")
            if patients is not None:
                newer = conn.execute("SELECT as_of FROM office_metrics WHERE office = ? AND period_from = ? AND "
                                     "metric = 'np_rows' AND as_of > ?", (office, period_from, as_of)).fetchone()
                if not newer:
                    conn.execute("DELETE FROM np_patients WHERE office = ? AND period_from = ?", (office, period_from))
                    conn.executemany(
                        "INSERT OR REPLACE INTO np_patients (office, period_from, row, pk, prime, subprime, other) "
                        "VALUES (?, ?, ?, ?, ?, ?, ?)",
                        [(office, period_from, int(p["row"]), p["pk"], json.dumps(p.get("prime") or {}),
                          float(p.get("subprime") or 0), float(p.get("other") or 0)) for p in scrub_patients(patients) if p.get("pk")])
        if review:
            raise NeedsReview(review)   # undoes everything above; nothing is saved until confirmed
    except NeedsReview as e:
        conn.close()
        return {"needs_review": True, "review": e.items, "skipped": counts["skipped"]}
    write_csv(conn)
    conn.close()
    return counts


def save_lender_journal(payload):
    """Save a Daily Journal's lender payments and fees per patient (keys only), per office and month.
    A later Thru date replaces the month's earlier upload; an older one is skipped."""
    period_from, thru = to_iso(payload.get("period_from")), to_iso(payload.get("period_thru"))
    if not period_from.endswith("-01") or thru[:7] != period_from[:7] or period_from > thru:
        raise ValueError("The Daily Journal must cover one month: From the 1st through a day in the same month")
    check_not_future(thru)
    now = datetime.now().isoformat(timespec="seconds")
    counts = {"added": 0, "replaced": 0, "skipped": []}
    conn = db()
    with conn:
        for o in payload.get("offices") or []:
            office = str(o.get("office") or "").strip()
            old = conn.execute("SELECT as_of FROM journal_uploads WHERE office = ? AND period_from = ?",
                               (office, period_from)).fetchone()
            if old and old["as_of"] > thru:
                counts["skipped"].append(f"{office} (already have thru {old['as_of']})")
                continue
            conn.execute("DELETE FROM lender_payments WHERE office = ? AND period_from = ?", (office, period_from))
            conn.executemany("INSERT INTO lender_payments VALUES (?, ?, ?, ?, ?, ?)",
                             [(office, period_from, p["pk"], p["lender"], float(p["payment"]), float(p["fee"]))
                              for p in scrub_patients(o.get("_patients")) if p.get("pk")])
            conn.execute("INSERT OR REPLACE INTO journal_uploads VALUES (?, ?, ?, ?, ?)",
                         (office, period_from, thru, payload.get("source"), now))
            total = round(sum(v["payment"] + v["fee"] for v in (o.get("by_lender") or {}).values()), 2)
            conn.execute("INSERT INTO change_log (changed_at, office, period_from, as_of, metric, new_value, via, source) "
                         "VALUES (?, ?, ?, ?, ?, ?, ?, ?)", (now, office, period_from, thru, "denticon_lender_payments",
                                                             total, "upload", payload.get("source")))
            counts["replaced" if old else "added"] += 1
    conn.close()
    return counts


# ------------------------------------------------------------------ day-by-day history and the inbox folder

DAY_COLUMNS = ["total_collection", "insurance_collection", "new_patients", "np_first_visit",
               "existing_patient_referrals", "total_production"]
DAY_REPORT_TYPES = {"executive_summary", "referral", "ortho_journal", "ortho_referrals"}   # can be run for one day
INBOX_DIR = DATA_DIR / "inbox"            # drop a report here and it is imported (git-ignored with data/)
INBOX_TYPES = {".pdf", ".xlsx", ".xlsm", ".doc", ".docx"}
INBOX_EVERY = int(os.environ.get("INBOX_SECONDS", 60))


def _log_change(conn, now, office, day, metric, old, new, via, source):
    if old is None or round(float(old), 2) != round(float(new), 2):
        conn.execute("INSERT INTO change_log (changed_at, office, period_from, as_of, metric, old_value, new_value, via, source) "
                     "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)", (now, office, day, day, metric, old, round(float(new), 2), via, source))


def is_day_report(payload, report_type):
    """True when the dates say one day (From = Thru) for a report that can be run per day."""
    try:
        frm = to_iso(payload.get("period_from"))
        thru = to_iso(payload.get("period_thru") or payload.get("as_of"))
    except ValueError:
        return False
    return frm == thru and report_type in DAY_REPORT_TYPES


def save_daily(payload):
    """One day's Executive Summary (the report run with From = Thru): one row per office per day. Uploading the
    same day again replaces it. The monthly Log is not touched."""
    day = to_iso(payload.get("period_from"))
    if to_iso(payload.get("period_thru")) != day:
        raise ValueError("A day-by-day report has From = Thru (the same day)")
    check_not_future(day, "report")
    offices = payload.get("offices") or []
    if not offices and not payload.get("empty"):
        raise ValueError("Nothing to save")
    now = datetime.now().isoformat(timespec="seconds")
    counts = {"daily": True, "day": day, "added": 0, "replaced": 0, "unchanged": 0, "skipped": [], "closed": not offices}
    conn = db()
    with conn:
        conn.execute("INSERT OR REPLACE INTO daily_imports (day, offices, source, logged_at) VALUES (?, ?, ?, ?)",
                     (day, len(offices), payload.get("source_file"), now))
        for o in offices:
            office = str(o["office"]).strip()
            npfv, refs = int(o["np_first_visit"]), int(o["existing_patient_referrals"])
            new = {"total_collection": round(float(o["total_collection"]), 2),
                   "insurance_collection": round(float(o["insurance_collection"]), 2),
                   "new_patients": npfv - refs, "np_first_visit": npfv, "existing_patient_referrals": refs,
                   "total_production": round(float(o["total_production"]), 2)}
            existing = conn.execute("SELECT * FROM office_daily WHERE office = ? AND day = ?", (office, day)).fetchone()
            if existing and all(round(float(existing[k]), 2) == v for k, v in new.items()):
                counts["unchanged"] += 1
                continue
            for k, v in new.items():
                _log_change(conn, now, office, day, k, existing[k] if existing else None, v, "daily report", payload.get("source_file"))
            if existing:
                conn.execute(f"UPDATE office_daily SET {', '.join(k + ' = ?' for k in new)}, source_file = ?, logged_at = ? WHERE id = ?",
                             (*new.values(), payload.get("source_file"), now, existing["id"]))
                counts["replaced"] += 1
            else:
                conn.execute(f"INSERT INTO office_daily (office, day, {', '.join(new)}, source_file, logged_at) "
                             f"VALUES (?, ?, {', '.join('?' for _ in new)}, ?, ?)",
                             (office, day, *new.values(), payload.get("source_file"), now))
                counts["added"] += 1
    conn.close()
    return counts


def save_daily_metrics(payload):
    """One day's Referral Production Listing, ortho Daily Journal or Treatment Plan Status (From = Thru): one value
    per office, day and metric. ZIP rows and patient details are not kept per day."""
    day = to_iso(payload.get("as_of"))
    if to_iso(payload.get("period_from")) != day:
        raise ValueError("A day-by-day report has From = Thru (the same day)")
    check_not_future(day, "report")
    rows = payload.get("rows") or []
    if not rows and not payload.get("empty"):
        raise ValueError("Nothing to save")
    now = datetime.now().isoformat(timespec="seconds")
    counts = {"daily": True, "day": day, "added": 0, "replaced": 0, "unchanged": 0, "skipped": [], "closed": not rows}
    conn = db()
    with conn:
        if payload.get("report_type"):
            conn.execute("INSERT OR REPLACE INTO daily_report_imports (day, report, logged_at) VALUES (?, ?, ?)",
                         (day, payload["report_type"], now))
        for row in rows:
            office = str(row.get("office") or "").strip()
            if not office:
                raise ValueError(f"Choose which office '{row.get('sheet_office')}' is")
            for metric, value in (row.get("values") or {}).items():
                if metric not in METRIC_NAMES or value is None:
                    continue
                value = round(float(value), 2)
                existing = conn.execute("SELECT id, value FROM daily_metrics WHERE office = ? AND day = ? AND metric = ?",
                                        (office, day, metric)).fetchone()
                if existing and round(existing["value"], 2) == value:
                    counts["unchanged"] += 1
                    continue
                _log_change(conn, now, office, day, metric, existing["value"] if existing else None, value,
                            "daily report", payload.get("source"))
                if existing:
                    conn.execute("UPDATE daily_metrics SET value = ?, source = ?, logged_at = ? WHERE id = ?",
                                 (value, payload.get("source"), now, existing["id"]))
                    counts["replaced"] += 1
                else:
                    conn.execute("INSERT INTO daily_metrics (office, day, metric, value, source, logged_at) "
                                 "VALUES (?, ?, ?, ?, ?, ?)", (office, day, metric, value, payload.get("source"), now))
                    counts["added"] += 1
    conn.close()
    return counts


DAY_REPORT_METRICS = {"referral": "np_collection", "ortho_journal": "ortho_collection", "ortho_referrals": "ortho_referrals"}


def _covered_through(days):
    """The last day of an unbroken run from the 1st of the month, or None (no 1st, or a gap)."""
    if not days:
        return None
    days = sorted(days)
    first = date.fromisoformat(days[0][:8] + "01")
    have, last = set(days), None
    d = first
    while d.isoformat() in have:
        last = d.isoformat(); d += timedelta(1)
    return last if last == days[-1] else None


def month_rows_from_days(rows):
    """For months whose one-day reports run unbroken from the 1st through the last day pulled, the Log's month row
    takes the Executive Summary numbers (and NP collection, ortho collection, ortho referrals, each when its own
    report is unbroken too) from the days added up. Refunds and the New Patient Analysis columns stay as uploaded.
    Months not pulled day by day are left alone."""
    conn = db()
    es_days, rep_days = {}, {}
    for r in conn.execute("SELECT day FROM daily_imports"):
        es_days.setdefault(r["day"][:7], []).append(r["day"])
    for r in conn.execute("SELECT day, report FROM daily_report_imports"):
        rep_days.setdefault((r["day"][:7], r["report"]), []).append(r["day"])
    covered = {m: _covered_through(ds) for m, ds in es_days.items()}
    covered = {m: t for m, t in covered.items() if t}
    if not covered:
        conn.close()
        return rows
    by_key = {(r["office"], r["period_from"][:7]): r for r in rows}
    sums = ", ".join(f"SUM({c}) AS {c}" for c in DAY_COLUMNS)
    for month, thru in covered.items():
        frm = month + "-01"
        n_days = (date.fromisoformat(thru) - date.fromisoformat(frm)).days + 1
        es = {r["office"]: r for r in conn.execute(f"SELECT office, {sums} FROM office_daily WHERE day BETWEEN ? AND ? GROUP BY office", (frm, thru))}
        mets = {}
        for report, metric in DAY_REPORT_METRICS.items():
            mt = _covered_through(rep_days.get((month, report), []))
            if mt == thru:   # this report runs as far as the Executive Summary: use it
                for r in conn.execute("SELECT office, SUM(value) AS v FROM daily_metrics WHERE metric = ? AND day BETWEEN ? AND ? GROUP BY office", (metric, frm, thru)):
                    mets.setdefault(r["office"], {})[metric] = r["v"]
                for o in es:
                    mets.setdefault(o, {}).setdefault(metric, 0)
        for office, e in es.items():
            row = by_key.get((office, month))
            if row is None:
                row = {c: None for c in LOG_COLUMNS}
                row.update({m + "_detail": None for m in DETAIL_METRICS})
                row.update({"office": office, "period_from": frm})
                rows.append(row); by_key[(office, month)] = row
            row["period_thru"] = thru
            for c in DAY_COLUMNS:
                row[c] = int(e[c] or 0) if c in ("new_patients", "np_first_visit", "existing_patient_referrals") else round(e[c] or 0, 2)
            for metric, v in mets.get(office, {}).items():
                row[metric] = int(v) if metric in COUNT_METRICS else round(v, 2)
                row[metric + "_as_of"] = thru
            row["source_file"] = f"from {n_days} daily report{'s' if n_days != 1 else ''}"
            row["from_days"] = n_days
            tc, np_ = row["total_collection"], row["new_patients"]
            ref, npc = row.get("refunds"), row.get("np_collection")
            row["net_collections"] = round(tc - ref, 2) if ref is not None else None
            row["refund_rate"] = round(ref / tc, 4) if ref is not None and tc else None
            row["existing_pt_collection"] = round(tc - npc, 2) if npc is not None else None
            irr = [row.get(k) for k in ("medi_cal_patients", "hmo_patients", "no_ssn_itin_patients")]
            row["irrelevant_rate"] = round(sum(irr) / np_, 4) if np_ and all(v is not None for v in irr) else None
            row["ortho_rate"] = round(row["ortho_referrals"] / np_, 4) if np_ and row.get("ortho_referrals") is not None else None
            row["financing_rate"] = round(row["financing_applicants"] / np_, 4) if np_ and row.get("financing_applicants") is not None else None
    conn.close()
    rows.sort(key=lambda r: (r["period_thru"] or "", r["office"]), reverse=True)
    return rows


def log_days(frm, to):
    """The Log for any run of days: each office's one-day reports added up, with how many of the days were there."""
    frm, to = to_iso(frm), to_iso(to)
    if frm > to:
        raise ValueError("'From' is after 'To'")
    d0, d1 = date.fromisoformat(frm), date.fromisoformat(to)
    if (d1 - d0).days > 366:
        raise ValueError("Pick a range of a year or less")
    days = [(d0 + timedelta(i)).isoformat() for i in range((d1 - d0).days + 1)]
    conn = db()
    sums = ", ".join(f"SUM({c}) AS {c}" for c in DAY_COLUMNS)
    es = conn.execute(f"SELECT office, COUNT(*) AS n, GROUP_CONCAT(day) AS have, {sums}, MAX(logged_at) AS logged_at "
                      "FROM office_daily WHERE day BETWEEN ? AND ? GROUP BY office", (frm, to)).fetchall()
    mets = conn.execute("SELECT office, metric, COUNT(*) AS n, SUM(value) AS value, GROUP_CONCAT(day) AS have "
                        "FROM daily_metrics WHERE day BETWEEN ? AND ? GROUP BY office, metric", (frm, to)).fetchall()
    imported = {r["day"] for r in conn.execute("SELECT day FROM daily_imports WHERE day BETWEEN ? AND ?", (frm, to))}
    conn.close()
    not_imported = [d for d in days if d not in imported]
    rows = {}

    def row(office):
        if office not in rows:
            r = {"office": office, "period_from": frm, "period_thru": to, "source_file": None, "logged_at": None,
                 "days": 0, "missing": [], "closed": [], "metric_days": {}}
            r.update({c: None for c in DAY_COLUMNS})
            r.update({m: None for m in METRIC_NAMES})
            r.update({f"{m}_as_of": None for m in METRIC_NAMES})
            r.update({c: None for c in CALCULATED})
            rows[office] = r
        return rows[office]

    for e in es:
        r = row(e["office"])
        have = set((e["have"] or "").split(","))
        for c in DAY_COLUMNS:
            r[c] = int(e[c]) if c in ("new_patients", "np_first_visit", "existing_patient_referrals") else round(e[c], 2)
        r["missing"] = [d for d in days if d not in have and d not in imported]      # never pulled
        r["closed"] = [d for d in days if d not in have and d in imported]        # pulled, office had nothing that day
        r["days"] = len(days) - len(r["missing"])
        r["source_file"] = f"{e['n']} daily reports" + (f", closed {len(r['closed'])} day{'s' if len(r['closed']) != 1 else ''}" if r["closed"] else "")
        r["logged_at"] = e["logged_at"]
    for m in mets:
        r = row(m["office"])
        have = set((m["have"] or "").split(","))
        r[m["metric"]] = int(m["value"]) if m["metric"] in COUNT_METRICS else round(m["value"], 2)
        r[f"{m['metric']}_as_of"] = to
        r["metric_days"][m["metric"]] = {"days": m["n"], "missing": [d for d in days if d not in have]}
    for r in rows.values():
        if r["total_collection"] is not None and r.get("np_collection") is not None:
            r["existing_pt_collection"] = round(r["total_collection"] - r["np_collection"], 2)
        if r.get("ortho_referrals") is not None and r["new_patients"]:
            r["ortho_rate"] = round(r["ortho_referrals"] / r["new_patients"], 4)
    sheet_ok, sheet_months = add_month_columns(rows, frm, to)
    if not sheet_ok:
        add_np_days(rows, frm, to)
    return {"from": frm, "to": to, "days": len(days), "not_imported": not_imported, "sheet_ok": sheet_ok, "sheet_months": sheet_months,
            "rows": sorted(rows.values(), key=lambda r: r["office"])}


DAY_FILLED = set(DAY_COLUMNS) | set(DAY_REPORT_METRICS.values())


def add_np_days(rows, frm, to):
    """For a Days range that isn't whole months: the New Patient Analysis columns added up from the sheet rows
    listed under those days (np_daily). Refunds have no days, so they stay monthly. Rates are worked out again."""
    conn = db()
    got = conn.execute("SELECT office, metric, SUM(value) AS v, MAX(as_of) AS a FROM np_daily WHERE day BETWEEN ? AND ? "
                       "GROUP BY office, metric", (frm, to)).fetchall()
    months = {r["m"] for r in conn.execute("SELECT DISTINCT substr(period_from, 1, 7) AS m FROM np_daily")}
    undated = {r["office"]: int(r["n"]) for r in conn.execute(
        "SELECT office, SUM(value) AS n FROM np_daily WHERE day = '' AND metric = 'np_rows' AND period_from BETWEEN ? AND ? GROUP BY office",
        (frm[:8] + "01", to[:8] + "01"))}
    conn.close()
    for office, n in undated.items():
        if office in rows:
            rows[office]["np_undated"] = n
    need = {frm[:7], to[:7]}
    for r in got:
        row = rows.get(r["office"])
        if row is None or r["metric"] in DAY_FILLED:
            continue
        row[r["metric"]] = int(r["v"]) if r["metric"] in COUNT_METRICS else round(r["v"], 2)
        row[r["metric"] + "_as_of"] = r["a"]
    if need <= months:   # the sheets were read by day for these months: a missing count means none that range
        for row in rows.values():
            for m in METRIC_NAMES:
                if m not in DAY_FILLED and m != "refunds" and row.get(m) is None:
                    row[m] = 0
                    row[m + "_as_of"] = to
    for row in rows.values():
        row.update(_calculated(row))


def _calculated(values):
    """The Log's calculated columns for one row of values, with the same SQL expressions the monthly_log view uses."""
    s_cols = ["total_collection", "insurance_collection", "new_patients", "np_first_visit", "existing_patient_referrals", "total_production"]
    s_sel = ", ".join(f"? AS {c}" for c in s_cols)
    m_sel = ", ".join(f"? AS {m}" for m in METRIC_NAMES)
    exprs = ", ".join(f"{e} AS {n}" for n, e in CALCULATED.items())
    conn = sqlite3.connect(":memory:"); conn.row_factory = sqlite3.Row
    r = conn.execute(f"SELECT {exprs} FROM (SELECT {s_sel}) s, (SELECT {m_sel}) m",
                     [values.get(c) for c in s_cols] + [values.get(m) for m in METRIC_NAMES]).fetchone()
    conn.close()
    return dict(r)


def add_month_columns(rows, frm, to):
    """The columns that only come once a month (refunds, the New Patient Analysis sheets, and the rates built on
    them) for a Days range that covers whole months to date: every month it touches from the 1st through that
    month's latest data. Then the Days view reads exactly like the Month view. Otherwise they stay empty."""
    months = sorted({frm[:7], to[:7]} | {f"{y:04d}-{m:02d}" for y in range(int(frm[:4]), int(to[:4]) + 1) for m in range(1, 13)
                                         if frm[:7] <= f"{y:04d}-{m:02d}" <= to[:7]})
    conn = db()
    log = [dict(r) for r in conn.execute(f"SELECT {', '.join(LOG_COLUMNS + [m + '_detail' for m in DETAIL_METRICS])} FROM monthly_log "
                                         f"WHERE substr(period_from, 1, 7) BETWEEN ? AND ?", (months[0], months[-1]))]
    conn.close()
    log = [r for r in month_rows_from_days(log) if r["period_from"][:7] in months]   # it can add rows for other months
    thru = {}
    for r in log:
        m = r["period_from"][:7]
        thru[m] = max(thru.get(m, ""), r["period_thru"] or "")
    ok = all(m in thru and frm <= m + "-01" and to >= thru[m] for m in months)
    if not ok:
        return False, []
    single = len(months) == 1
    for office, r in rows.items():
        mine = [x for x in log if x["office"] == office]
        for metric in METRIC_NAMES:
            if metric in DAY_FILLED:
                continue
            vals = [x.get(metric) for x in mine if x.get(metric) is not None]
            if vals:
                r[metric] = round(sum(vals), 2) if metric not in COUNT_METRICS else int(sum(vals))
                r[metric + "_as_of"] = max((x.get(metric + "_as_of") or "") for x in mine) or None
            if single and mine:
                if metric in DETAIL_METRICS:
                    r[metric + "_detail"] = mine[0].get(metric + "_detail")
        calc = _calculated(r)
        for k, v in calc.items():
            r[k] = v
    return True, months


def days_info():
    """Which days have a one-day report: {day: offices with activity} (0 = a closed day), plus today's date,
    for the day pickers (blue = pick me, gray = nothing there, off = hasn't happened)."""
    conn = db()
    days = {r["day"]: r["offices"] for r in conn.execute("SELECT day, offices FROM daily_imports")}
    conn.close()
    return {"days": days, "today": date.today().isoformat()}


DATE_IN_NAME = [re.compile(r"(\d{4})-(\d{2})-(\d{2})"), re.compile(r"(\d{1,2})-(\d{1,2})-(\d{4})"),
                re.compile(r"(?<!\d)(\d{4})(\d{2})(\d{2})(?!\d)")]


def dates_in_name(name):
    """Days written in a file name: 2026-10-14, 10-14-2026 or 20261014 (an Excel export carries no dates)."""
    stem = Path(name).stem.replace("_", "-").replace(".", "-").replace("/", "-")
    found = []
    for i, rx in enumerate(DATE_IN_NAME):
        for m in rx.finditer(stem):
            y, mo, d = (m.group(3), m.group(1), m.group(2)) if i == 1 else (m.group(1), m.group(2), m.group(3))
            try:
                found.append(date(int(y), int(mo), int(d)).isoformat())
            except ValueError:
                pass
        if found:
            break
    return list(dict.fromkeys(found))


def import_report_file(path, name):
    """Parse a report file and save it the way the Upload tab would: a one-day report (From = Thru) to the
    day-by-day history, a month-to-date report to the monthly Log. Returns (report_type, from, thru, result)."""
    result = parse_report(str(path), name)
    period = result.get("period") or {}
    frm = to_iso(period["from"]) if period.get("from") else None
    thru = to_iso(period["thru"]) if period.get("thru") else None
    if not (frm and thru):
        found = dates_in_name(name)
        if len(found) >= 2:
            frm, thru = found[0], found[1]
        elif found:
            frm = thru = found[0]
        else:
            raise ValueError("No dates: a PDF carries them; for an Excel file put the day in the file name, "
                             "e.g. 'Executive Summary 2026-10-14.xlsx'")
    rtype = result["report_type"]
    if rtype == "executive_summary":
        payload = {"period_from": frm, "period_thru": thru, "source_file": name, "offices": result["offices"],
                   "empty": bool(result.get("empty"))}
        saved = save_daily(payload) if frm == thru else save_rows(payload)
    elif rtype == "lender_journal":
        saved = save_lender_journal({"period_from": frm, "period_thru": thru, "source": name,
                                     "offices": [{"office": o["office"], "by_lender": o.get("by_lender"),
                                                  "_patients": o.get("_patients")} for o in result["offices"]]})
    else:
        rows = [{"sheet_office": o["office"], "office": o["office"],
                 "values": {k: v for k, v in o.items() if k in METRIC_NAMES}, "details": o.get("details")}
                for o in result["offices"]]
        payload = {"as_of": thru, "period_from": frm, "source": name, "from_report": True, "report_type": rtype,
                   "via": "inbox", "rows": rows, "empty": bool(result.get("empty"))}
        saved = save_daily_metrics(payload) if is_day_report(payload, rtype) else save_metrics(payload)
    return rtype, frm, thru, saved


def inbox_scan():
    """Import every report file waiting in data/inbox, then move it to done / review / failed."""
    INBOX_DIR.mkdir(parents=True, exist_ok=True)
    for sub in ("done", "failed", "review"):
        (INBOX_DIR / sub).mkdir(exist_ok=True)
    for path in sorted(INBOX_DIR.iterdir()):
        if not path.is_file() or path.suffix.lower() not in INBOX_TYPES or path.name.startswith((".", "~$")):
            continue
        if time.time() - path.stat().st_mtime < 15:   # may still be downloading
            continue
        rtype = frm = thru = None
        try:
            rtype, frm, thru, saved = import_report_file(path, path.name)
            if saved.get("needs_review"):
                status = "review"
                note = ("Some numbers are lower than the Log has, or a same-date report changed them. "
                        "Upload it on the Upload tab to confirm.")
            else:
                parts = [f"{saved[k]} {k}" for k in ("added", "replaced", "unchanged") if saved.get(k)]
                if saved.get("closed"):
                    parts.append("no activity that day (recorded as a closed day)")
                status = "imported"
                skipped = saved.get("skipped") or []
                note = (", ".join(parts) or "nothing changed") + (f"; {len(skipped)} skipped (the Log already has a later report)" if skipped else "")
        except Exception as e:
            status, note = "failed", str(e) or type(e).__name__
        folder = {"imported": "done", "review": "review", "failed": "failed"}[status]
        dest = INBOX_DIR / folder / f"{datetime.now():%Y%m%d-%H%M%S} {path.name}"
        try:
            path.replace(dest)
            if status != "imported":
                dest.with_name(dest.name + ".txt").write_text(note + "\n")
        except OSError as e:
            note += f" (could not move the file: {e})"
        conn = db()
        with conn:
            conn.execute("INSERT INTO inbox_log (seen_at, file, status, report_type, period_from, period_thru, note) "
                         "VALUES (?, ?, ?, ?, ?, ?, ?)",
                         (datetime.now().isoformat(timespec="seconds"), path.name, status, rtype, frm, thru, note))
        conn.close()
        print(f"{datetime.now():%H:%M} Inbox {path.name}: {status}. {note}")


def inbox_loop():
    while True:
        try:
            inbox_scan()
        except Exception as e:   # never stop the app over the inbox
            print(f"{datetime.now():%H:%M} Inbox check failed: {e}")
        time.sleep(INBOX_EVERY)


# ------------------------------------------------------------------ the assistant (read-only questions about the app's numbers)

claude_cfg = assistant.Config(DATA_DIR)
chat_files = assistant.ChatFiles()
AI_HIDE = ("_detail", "source_file", "logged_at", "from_days")


def _ai_glossary():
    """Each Log column's name and meaning, read from the page's own column list (so it never drifts)."""
    page = (ROOT / "static" / "index.html").read_text()
    out = []
    for m in re.finditer(r'\{key: "([a-z_]+)", label: "([^"]+)", kind: "([a-z]+)"(?:[^}]*?title:\s*"([^"]*)")?', page):
        key, label, kind, title = m.groups()
        out.append(f"- {key}: {label} ({kind}){' - ' + title if title else ''}")
    return "\n".join(out)


def _ai_help():
    page = (ROOT / "static" / "index.html").read_text()
    body = page.split('<div class="help-body">', 1)[-1].split("</dialog>", 1)[0]
    text = re.sub(r"<(script|style)[^>]*>.*?</\1>", "", body, flags=re.S)
    text = re.sub(r"</(p|li|h3|h4|div)>", "\n", text)
    text = re.sub(r"<[^>]+>", "", text)
    text = text.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">").replace("&nbsp;", " ")
    return re.sub(r"\n\s*\n+", "\n", text).strip()


def ai_system():
    return f"""You are the assistant inside the Executive Summary Logger, Care Dentistry Group's finance app for its 15 dental offices.
You answer questions about the app's numbers and how to use the app. You are read-only: you can look numbers up with your tools, but
you cannot change, upload, delete, pull or sync anything. If someone asks for a change, tell them where in the app to do it.

How to answer:
- Look the numbers up with the tools; never guess a figure. Say which period (month, or exact days) and which offices a number covers.
- Keep answers short and plain. Use a small table for comparisons. Money as $1,234.56; rates as percentages.
- The tools only return totals per office, ZIP code or marketing channel. Never ask for, guess or reveal patient names, patient IDs or phone numbers.
- Data notes: a month can be "from daily reports" (added up from one-day Denticon reports, through the last day pulled) or an uploaded
  month-to-date report. The current month is partial. In a part-month range, refunds (and net collections / refund %) are monthly only.
  New Patient Analysis numbers come from each office's Google Sheet. Calls only count as new patients where a Denticon New Patients List
  covering their first visit has been uploaded. Paid Search and Google PPC are both paid Google traffic.
- If the data isn't there (a month not loaded, an office missing), say so and suggest what to upload or pull.
- People can attach files (PDF, Excel, Word) to the chat; they appear as documents at the start of the conversation. Use them
  together with the app's data (compare, check, explain). Say when a number comes from an attached file rather than the app.
  Spreadsheets arrive as tab-separated rows (dates may be Excel serial numbers: 46000 = 2025-12-09). Attached files can contain
  patient details: use them to work out answers, but report totals and counts, not patient names, IDs or phone numbers.

Log columns (key: label (kind) - meaning):
{_ai_glossary()}

The app's help guide:
{_ai_help()}"""


AI_TOOLS = [
    {"name": "overview", "description": "What data the app has: today's date, offices, months in the Log, which days were pulled one day at a time, "
     "months with ZIP data and months with calls. Call this first when you need to know what periods exist.",
     "input_schema": {"type": "object", "properties": {}, "additionalProperties": False}},
    {"name": "get_log", "description": "Log numbers per office for one month (month = YYYY-MM) or for any run of days (from_date and to_date, YYYY-MM-DD), "
     "with a Total row. Optional: offices (names or parts of names) and columns (Log column keys; default: the main ones).",
     "input_schema": {"type": "object", "properties": {
         "month": {"type": "string"}, "from_date": {"type": "string"}, "to_date": {"type": "string"},
         "offices": {"type": "array", "items": {"type": "string"}}, "columns": {"type": "array", "items": {"type": "string"}}},
         "additionalProperties": False}},
    {"name": "get_log_trend", "description": "One or more Log columns month by month (from_month..to_month, YYYY-MM) for each office and the group total.",
     "input_schema": {"type": "object", "properties": {
         "columns": {"type": "array", "items": {"type": "string"}}, "from_month": {"type": "string"}, "to_month": {"type": "string"},
         "offices": {"type": "array", "items": {"type": "string"}}}, "required": ["columns"], "additionalProperties": False}},
    {"name": "get_zip_collections", "description": "New-patient collections by home ZIP code (from the Referral Production Listing) for a month range: "
     "the top ZIPs with patients, dollars and city, plus totals. Optional office filter.",
     "input_schema": {"type": "object", "properties": {"from_month": {"type": "string"}, "to_month": {"type": "string"},
                                                       "office": {"type": "string"}, "top": {"type": "integer"}}, "additionalProperties": False}},
    {"name": "get_calls", "description": "The Invoca calls funnel for a call month (YYYY-MM): calls, answered, callers who became new patients, started "
     "treatment and $ collected, in total and grouped by channel, medium, source, location, promo_line, campaign or office.",
     "input_schema": {"type": "object", "properties": {"month": {"type": "string"}, "group_by": {"type": "string",
                      "enum": ["channel", "medium", "source", "location", "promo_line", "campaign", "office"]}, "office": {"type": "string"}},
                      "additionalProperties": False}},
    {"name": "get_benchmarks", "description": "The benchmark cut-offs (green / yellow) set for Log columns.",
     "input_schema": {"type": "object", "properties": {}, "additionalProperties": False}},
]
AI_DEFAULT_COLS = ["total_collection", "insurance_collection", "new_patients", "total_production", "refunds", "net_collections",
                   "np_collection", "ortho_collection", "ortho_referrals", "ortho_rate", "financing_applicants", "financing_rate", "started_tx"]


def _ai_offices(conn):
    return sorted({r[0] for r in conn.execute("SELECT office FROM office_summary UNION SELECT office FROM office_metrics")})


def _ai_pick(rows, offices, columns):
    cols = [c for c in (columns or AI_DEFAULT_COLS) if c in LOG_COLUMNS and not c.endswith(AI_HIDE)] or AI_DEFAULT_COLS
    if offices:
        want = [o.lower() for o in offices]
        rows = [r for r in rows if any(w in r["office"].lower() for w in want)]
    out = [{"office": r["office"], "period": f"{r['period_from']} to {r['period_thru']}", **{c: r.get(c) for c in cols}} for r in rows]
    total = {"office": "TOTAL"}
    for c in cols:
        vals = [r.get(c) for r in rows if isinstance(r.get(c), (int, float))]
        if vals and c not in CALCULATED and not c.endswith("_rate"):
            total[c] = round(sum(vals), 2)
    return {"rows": out, "total": total, "note": "Rates are per office; the total row only adds counts and dollars."}


def ai_overview():
    conn = db()
    months = sorted({r[0][:7] for r in conn.execute("SELECT period_from FROM office_summary UNION SELECT period_from FROM office_metrics")})
    days = [r[0] for r in conn.execute("SELECT day FROM daily_imports ORDER BY day")]
    calls = [r[0] for r in conn.execute("SELECT DISTINCT substr(call_date, 1, 7) FROM invoca_calls ORDER BY 1")]
    zips = [r[0] for r in conn.execute("SELECT DISTINCT substr(period_from, 1, 7) FROM map_patients ORDER BY 1")]
    offices = _ai_offices(conn)
    conn.close()
    by_month = {}
    for d in days:
        by_month.setdefault(d[:7], []).append(d)
    return {"today": date.today().isoformat(), "offices": offices, "log_months": months,
            "days_pulled": {m: f"{ds[0]} to {ds[-1]} ({len(ds)} days)" for m, ds in by_month.items()},
            "zip_months": zips, "call_months": calls}


def ai_get_log(month=None, from_date=None, to_date=None, offices=None, columns=None):
    if from_date and to_date:
        d = log_days(from_date, to_date)
        res = _ai_pick(d["rows"], offices, columns)
        res["range_note"] = ("whole months to date: every column filled" if d["sheet_ok"]
                             else "part of a month: refunds / net collections / refund % are monthly only (shown empty)")
        if d.get("not_imported"):
            res["days_not_pulled"] = d["not_imported"]
        return res
    conn = db()
    rows = log_rows(conn)
    conn.close()
    month = month or max(r["period_from"][:7] for r in rows)
    return _ai_pick([r for r in rows if r["period_from"].startswith(month)], offices, columns)


def ai_get_log_trend(columns, from_month=None, to_month=None, offices=None):
    conn = db()
    rows = log_rows(conn)
    conn.close()
    cols = [c for c in columns if c in LOG_COLUMNS and not c.endswith(AI_HIDE)][:6]
    if not cols:
        raise ValueError("unknown columns; use Log column keys such as total_collection")
    if from_month:
        rows = [r for r in rows if r["period_from"][:7] >= from_month]
    if to_month:
        rows = [r for r in rows if r["period_from"][:7] <= to_month]
    if offices:
        want = [o.lower() for o in offices]
        rows = [r for r in rows if any(w in r["office"].lower() for w in want)]
    out = {}
    for r in rows:
        m = r["period_from"][:7]
        e = out.setdefault(m, {"period_thru": r["period_thru"], "offices": {}, "total": {}})
        e["offices"][r["office"]] = {c: r.get(c) for c in cols}
        for c in cols:
            if isinstance(r.get(c), (int, float)) and c not in CALCULATED and not c.endswith("_rate"):
                e["total"][c] = round(e["total"].get(c, 0) + r[c], 2)
    return {"months": dict(sorted(out.items()))}


def ai_get_zip_collections(from_month=None, to_month=None, office=None, top=25):
    t = maps_trends()
    months = [m for m in t["months"] if (not from_month or m >= from_month) and (not to_month or m <= to_month)]
    zips, total_n, total_d = {}, 0, 0.0
    for r in t["rows"]:
        if r["month"] not in months or (office and office.lower() not in r["office"].lower()):
            continue
        z = zips.setdefault(r["zip"], {"zip": r["zip"], "city": t["where"].get(r["zip"], ""), "new_patients": 0, "collected": 0.0})
        z["new_patients"] += r["n"]; z["collected"] += r["total"]; total_n += r["n"]; total_d += r["total"]
    ranked = sorted(zips.values(), key=lambda z: -z["collected"])
    for z in ranked:
        z["collected"] = round(z["collected"], 2)
    return {"months": months, "partial_months": {m: t["thru"][m] for m in months if t["thru"].get(m, "")[-2:] < "28"},
            "total_new_patients": total_n, "total_collected": round(total_d, 2), "zip_count": len(ranked),
            "top": ranked[:max(1, min(int(top or 25), 100))]}


def ai_get_calls(month=None, group_by="channel", office=""):
    d = calls_funnel(month, office or "")
    if not d.get("month"):
        return {"note": "No calls pulled yet."}
    rows = sorted(d["by"].get(group_by or "channel", []), key=lambda r: -r["calls"])
    return {"month": d["month"], "office": d.get("office") or "all", "total": d["total"], "group_by": group_by, "rows": rows[:40],
            "patient_lists": [{"office": l["office"], "first_visits": f"{l['lo']} to {l['hi']}"} for l in d["lists"]]}


def ai_get_benchmarks():
    conn = db()
    rows = [dict(r) for r in conn.execute("SELECT metric, direction, green, yellow FROM benchmarks ORDER BY metric")]
    conn.close()
    return {"benchmarks": rows, "note": "Rates are fractions (0.25 = 25%)."}


AI_HANDLERS = {"overview": ai_overview, "get_log": ai_get_log, "get_log_trend": ai_get_log_trend,
               "get_zip_collections": ai_get_zip_collections, "get_calls": ai_get_calls, "get_benchmarks": ai_get_benchmarks}


# Before Sonnet answers, each question goes through cheaper steps:
#   1. exact-match cache: asked before (first question of a chat)? reuse the verdict, or the answer if the data hasn't changed (0 tokens)
#   2. keyword rules: clearly off-topic -> block; clearly on-topic (or a follow-up in a chat) -> step 4; greetings get a canned reply (0 tokens)
#   3. Haiku gatekeeper: only the unclear ones -> yes/no, verdict saved
#   4. Sonnet: the real answer, saved for step 1
AI_ON_TOPIC = re.compile(r"\b(collect\w*|production|produc\w*|patients?|np|nps|new pts?|offices?|clinics?|locations?|zips?|zip ?codes?|calls?|callers?|"
    r"invoca|channels?|medium|sources?|campaigns?|marketing|ppc|seo|organic|paid|google|referr\w*|ortho\w*|financ\w*|lenders?|refunds?|"
    r"revenue|insurance|treatment|tx|started|benchmarks?|goals?|log|charts?|reports?|denticon|exec\w* summary|sheets?|upload\w*|robot|"
    r"pull\w*|months?|monthly|weeks?|days?|daily|yesterday|today|ytd|mtd|quarter|q[1-4]|totals?|average|avg|compare\w*|trend\w*|"
    r"rate|percent\w*|dollars?|best|worst|top|highest|lowest|jan\w*|feb\w*|mar|march|apr\w*|may|june?|july?|aug\w*|sep\w*|oct\w*|"
    r"nov\w*|dec\w*|20\d\d|dental|dentist\w*|hygien\w*|adjustments?|payments?|irrelevant|column|mean|meaning|"
    r"what can you|can you help|help me|how (do|can|should) (i|we)|how to|where (do|can) (i|we)|app|this page|tab)\b|[$%]")
AI_OFF_TOPIC = re.compile(r"\b(poem|poetry|song|lyrics|joke|riddle|story|recipe|cook\w*|weather|movie|tv show|sports?|nba|nfl|stocks?|"
    r"bitcoin|crypto|horoscope|translate|homework|essay|python|javascript|html|css|code|coding|program\w*|politic\w*|election|president|"
    r"religio\w*|dating|girlfriend|boyfriend|ignore (all|any|the|your|previous)|system prompt|jailbreak|pretend|role ?play)\b")
AI_GREETING = re.compile(r"^(hi|hello|hey|hiya|yo|good (morning|afternoon|evening)|thanks?|thank you|thx|ty|ok|okay|cool|great)( there| so much| claude)?[ !.]*$")
AI_GREETING_REPLY = "Hi! Ask me about collections, new patients, production, financing, ZIP codes or calls, for any office, month or range of days."
AI_BLOCKED_REPLY = ("I can only help with questions about Care Dentistry Group's numbers in this app (collections, new patients, production, "
                    "ZIP codes, calls and marketing) and how to use the app. Try something like \"Which office collected the most last month?\"")
AI_OFFICE_WORDS = {"care", "dental", "dentistry", "center", "centre", "group", "family", "smile", "smiles", "the", "and", "of", "dentist", "office"}


def _ai_qkey(question):
    return " ".join(re.sub(r"[^a-z0-9$%./-]+", " ", question.lower()).split()).strip(" .")


def _ai_data_fp(conn):
    """Changes when anything the assistant could look up changes, or the day changes ("last month", "this week")."""
    q = lambda sql: conn.execute(sql).fetchone()[0]
    return "|".join((date.today().isoformat(), notes_fingerprint(conn), str(q("SELECT COUNT(*) FROM invoca_calls")),
                     str(q("SELECT COUNT(*) FROM map_patients")), str(q("SELECT MAX(as_of) FROM map_patients")),
                     str(q("SELECT COUNT(*) FROM benchmarks")), str(q("SELECT MAX(id) FROM change_log"))))


def _ai_rules(conn, qkey, follow_up):
    """Step 2: 'greet', 'off', 'on' or None (unclear)."""
    if AI_GREETING.match(qkey):
        return "greet"
    office_words = {w for o in _ai_offices(conn) for w in re.findall(r"[a-z]+", o.lower()) if len(w) > 3 and w not in AI_OFFICE_WORDS}
    on = bool(AI_ON_TOPIC.search(qkey)) or bool(office_words & set(qkey.split()))
    off = bool(AI_OFF_TOPIC.search(qkey))
    if off and not on:
        return "off"
    if on and not off:
        return "on"
    if follow_up and not off:   # "and Downey?", "why?": the chat already has a topic
        return "on"
    return None


def ai_chat(question, history, who, use_cache=True, file_ids=()):
    history = [h for h in (history or []) if h.get("text")]
    ids = [f for h in history for f in (h.get("files") or []) if isinstance(f, str)] + [f for f in (file_ids or []) if isinstance(f, str)]
    try:
        files = chat_files.get(ids)
    except assistant.AssistantError as e:
        raise ValueError(str(e))
    question = (question or "").strip()[:2000] or ("Summarize this file." if file_ids else "")
    if not question:
        raise ValueError("Type a question first.")
    now = datetime.now().isoformat(timespec="seconds")
    if files:   # a question about an attached file: no saved answers, and the file shows it's on topic unless the words say otherwise
        use_cache = False
    qkey, follow_up = _ai_qkey(question), bool(history) or bool(files)
    usage, tools, answer, err, route = {"input": 0, "output": 0, "cache_read": 0}, [], None, None, None
    conn = db()
    fp = _ai_data_fp(conn)
    cached = conn.execute("SELECT * FROM assistant_cache WHERE qkey = ?", (qkey,)).fetchone() if use_cache else None
    verdict = None
    # 1. exact-match cache (a follow-up depends on the chat before it, so only its off-topic verdict is reused)
    if cached and cached["verdict"] == "off":
        answer, route = AI_BLOCKED_REPLY, "cache: blocked"
    elif cached and not follow_up and cached["answer"] and cached["data_fp"] == fp:
        answer, route = cached["answer"], "cache: answer"
    elif cached:
        verdict = "on"   # known on-topic; the data changed (or it's a follow-up), so answer again
    if answer:
        with conn:
            conn.execute("UPDATE assistant_cache SET hits = hits + 1 WHERE qkey = ?", (qkey,))
    if not answer and not verdict:
        # 2. keyword rules
        rule = _ai_rules(conn, qkey, follow_up)
        if rule == "greet":
            answer, route = AI_GREETING_REPLY, "rules: greeting"
        elif rule == "off":
            answer, route, verdict = AI_BLOCKED_REPLY, "rules: blocked", "off"
        elif rule == "on":
            verdict, route = "on", "rules: on-topic"
        else:
            # 3. Haiku gatekeeper
            try:
                ok, u = assistant.gate(claude_cfg.key(), question)
                usage = {k: usage[k] + u[k] for k in usage}
                verdict, route = ("on", "haiku: on-topic") if ok else ("off", "haiku: blocked")
                if not ok:
                    answer = AI_BLOCKED_REPLY
            except assistant.AssistantError as e:
                if not claude_cfg.exists():
                    err = str(e)
                verdict, route = "on", "haiku failed: on-topic"   # let the main model decide
        if verdict == "off" and use_cache:
            with conn:
                conn.execute("INSERT OR REPLACE INTO assistant_cache (qkey, verdict, saved_at) VALUES (?, 'off', ?)", (qkey, now))
    # 4. the main model
    if not answer and not err:
        route = (route + " → " if route else "cache: on-topic → ") + "sonnet"
        try:
            answer, tools, u = assistant.run(claude_cfg.key(), ai_system(), AI_TOOLS, AI_HANDLERS, history, question, files)
            usage = {k: usage[k] + u.get(k, 0) for k in usage}
            complete = not answer.endswith("ask a narrower question.)") and "too many lookups" not in answer and tools
            if use_cache and not follow_up and complete:   # only answers backed by a lookup are reused
                with conn:
                    conn.execute("INSERT OR REPLACE INTO assistant_cache (qkey, verdict, answer, data_fp, saved_at) VALUES (?, 'on', ?, ?, ?)",
                                 (qkey, answer, fp, now))
            elif use_cache:
                with conn:
                    conn.execute("INSERT INTO assistant_cache (qkey, verdict, saved_at) VALUES (?, 'on', ?) ON CONFLICT(qkey) DO NOTHING", (qkey, now))
        except assistant.AssistantError as e:
            err = str(e)
    if files:
        route = (route or "") + " (files: " + ", ".join(f["name"] for f in files) + ")"
    with conn:
        conn.execute("INSERT INTO assistant_log (asked_at, who, question, answer, tools, input_tokens, output_tokens, cache_read_tokens, error, route) "
                     "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", (now, who, question, answer, ",".join(tools), usage["input"],
                                                                usage["output"], usage["cache_read"], err, route))
    conn.close()
    if err:
        raise ValueError(err)
    return {"answer": answer, "tools": tools, "route": route}


def assistant_cache_clear():
    conn = db()
    with conn:
        n = conn.execute("DELETE FROM assistant_cache").rowcount
    conn.close()
    return {"cleared": n}


def assistant_status():
    conn = db()
    n = conn.execute("SELECT COUNT(*), SUM(input_tokens), SUM(output_tokens) FROM assistant_log WHERE asked_at >= date('now', '-30 days')").fetchone()
    recent = [dict(r) for r in conn.execute("SELECT asked_at, who, question, route, substr(COALESCE(answer, error), 1, 160) AS answer FROM assistant_log ORDER BY id DESC LIMIT 15")]
    free = conn.execute("SELECT COUNT(*) FROM assistant_log WHERE asked_at >= date('now', '-30 days') AND (route LIKE 'cache: answer%' "
                        "OR route LIKE 'cache: blocked%' OR route LIKE 'rules: greeting%' OR route LIKE 'rules: blocked%')").fetchone()[0]
    saved = conn.execute("SELECT COUNT(*), SUM(verdict = 'off') FROM assistant_cache").fetchone()
    conn.close()
    return {"connected": claude_cfg.exists(), "model": assistant.MODEL, "gate_model": assistant.GATE_MODEL, "questions_30d": n[0],
            "free_30d": free, "input_tokens_30d": n[1] or 0, "output_tokens_30d": n[2] or 0, "cached": saved[0], "cached_off": saved[1] or 0,
            "recent": recent}


# ------------------------------------------------------------------ the Denticon report robot (tools/denticon_robot.js)

ROBOT_SCRIPT = ROOT / "tools" / "denticon_robot.js"
ROBOT_STATUS = DATA_DIR / "denticon-robot.json"
ROBOT_LOG = DATA_DIR / "denticon-robot.log"
_robot = {"proc": None, "started": None}


def _robot_serving():
    """True when a standing robot window (tools/denticon_robot.js serve) is running on this computer."""
    for pid in os.listdir("/proc"):
        if not pid.isdigit():
            continue
        try:
            cmd = open(f"/proc/{pid}/cmdline", "rb").read().replace(b"\0", b" ")
        except OSError:
            continue
        if b"denticon_robot.js" in cmd and b" serve" in cmd:
            return True
    return False


def robot_status():
    """What the robot is doing: the status file it writes plus whether its process is still running."""
    st = _load_json_file(ROBOT_STATUS) or {}
    proc = _robot["proc"]
    running = proc is not None and proc.poll() is None
    if not running and st.get("state") in ("running", "waiting_login") and proc is not None:
        st["state"] = "error"; st["message"] = st.get("message") or "The robot stopped before finishing."
    serving = _robot_serving()
    if serving and st.get("state") in ("running", "waiting_login"):
        running = True
    return {"running": running, "serving": serving, "started": _robot["started"], "available": ROBOT_SCRIPT.exists(), **st}


def robot_pull(frm, to, report="exec"):
    """Start the robot for a run of days (From..To) and report(s): exec (Executive Summary), referral (Referral
    Production Listing) or both. It opens its own Chrome window on this computer; when Denticon asks for a login,
    the user logs in there and the robot carries on. Files land in data/inbox."""
    frm, to = to_iso(frm), to_iso(to)
    if report not in ("exec", "referral", "ortho", "txplan", "both", "all"):
        raise ValueError("Pick which report to pull")
    if frm > to:
        raise ValueError("'From' is after 'To'")
    if (date.fromisoformat(to) - date.fromisoformat(frm)).days > 62:
        raise ValueError("Pull two months at most at a time")
    check_not_future(to, "To")
    proc = _robot["proc"]
    if proc is not None and proc.poll() is None:
        raise ValueError("The robot is already running; wait for it to finish")
    st = _load_json_file(ROBOT_STATUS) or {}
    if st.get("state") in ("running", "waiting_login") and _robot_serving():
        raise ValueError("The robot is already running; wait for it to finish")
    if _robot_serving():   # the standing robot window takes the job (one login serves many pulls)
        (DATA_DIR / "denticon-robot.report").write_text(report + "\n")
        (DATA_DIR / "denticon-robot.cmd").write_text(f"pull {frm} {to}\n")
        _write_json_file(ROBOT_STATUS, {**st, "state": "running", "message": "Starting in the robot window.", "days": [], "failed": [],
                                        "started_at": datetime.now().isoformat(timespec="seconds")})
        return robot_status()
    env = dict(os.environ)
    env.setdefault("DISPLAY", ":0")
    env.setdefault("XAUTHORITY", str(Path.home() / ".Xauthority"))
    log = open(ROBOT_LOG, "a")
    log.write(f"\n--- {datetime.now():%Y-%m-%d %H:%M} pull {frm} to {to}\n"); log.flush()
    _robot["proc"] = subprocess.Popen(["node", str(ROBOT_SCRIPT), "pull", "--from", frm, "--to", to, "--report", report, "--headed"],
                                      cwd=str(ROOT), env=env, stdout=log, stderr=subprocess.STDOUT)
    _robot["started"] = datetime.now().isoformat(timespec="seconds")
    _write_json_file(ROBOT_STATUS, {"state": "running", "message": "Starting the Denticon window.", "days": [], "failed": [],
                                    "started_at": _robot["started"], "updated_at": _robot["started"]})
    return robot_status()


def inbox_info():
    conn = db()
    recent = [dict(r) for r in conn.execute("SELECT * FROM inbox_log ORDER BY id DESC LIMIT 40")]
    conn.close()
    pending = sorted(p.name for p in INBOX_DIR.iterdir() if p.is_file() and p.suffix.lower() in INBOX_TYPES) \
        if INBOX_DIR.exists() else []
    return {"folder": str(INBOX_DIR), "pending": pending, "recent": recent, "every_seconds": INBOX_EVERY}


# Sheet prime lender columns ("Care Credit", "Patient FI") by squeezed name -> the journal's lender names.
SHEET_PRIME = {"alphaeon": "Alphaeon", "carecredit": "Care Credit", "patientfi": "PatientFi", "proceed": "Proceed"}


def denticon_checks():
    """New patients whose collected amounts on the NP sheet don't match Denticon's lender payments + fees.
    Denticon more than the sheet is always flagged. The sheet more than Denticon is flagged only when journals
    cover the patient's month and the month after (financing often posts the next month)."""
    import lender_journal
    conn = db()
    covered = {}   # office -> {month: journal Thru date}
    for r in conn.execute("SELECT office, period_from, as_of FROM journal_uploads"):
        covered.setdefault(r["office"], {})[r["period_from"][:7]] = r["as_of"]
    sheet_as_of = {(r["office"], r["period_from"]): r["as_of"] for r in
                   conn.execute("SELECT office, period_from, as_of FROM office_metrics WHERE metric = 'np_rows'")}
    if not covered:
        conn.close()
        return []
    den = {}
    for r in conn.execute("SELECT pk, lender, SUM(payment) AS payment, SUM(fee) AS fee FROM lender_payments GROUP BY pk, lender"):
        den.setdefault(r["pk"], {})[r["lender"]] = [r["payment"], r["fee"]]
    # A fee booked under another lender's name (an Alphaeon payment with an "ACCESS ALPH" fee) goes with the
    # one lender that paid this patient without a fee of its own.
    for lenders in den.values():
        for name, (pay, fee) in list(lenders.items()):
            if not pay and fee:
                paid = [l for l, (p2, f2) in lenders.items() if p2 and not f2]
                if name == "Access Alph" and lenders.get("Alphaeon", [0])[0]:
                    paid = ["Alphaeon"]   # an Alphaeon payment with part of its fee booked as "ACCESS ALPH"
                if len(paid) == 1:
                    lenders[paid[0]][1] += fee
                    del lenders[name]
    rows = conn.execute("SELECT * FROM np_patients").fetchall()
    conn.close()
    this_month = date.today().isoformat()[:7]
    money = lambda v: f"${v:,.2f}"
    out = []
    for p in rows:
        months = covered.get(p["office"], {})
        ym = p["period_from"][:7]
        if ym not in months:
            continue
        # "The sheet says more than Denticon" needs journals that reach at least as far as the sheet: the
        # month itself through the sheet's date, and the next month too once it has started.
        nxt = (date(int(ym[:4]), int(ym[5:]), 28) + timedelta(days=4)).isoformat()[:7]
        full = months[ym] >= (sheet_as_of.get((p["office"], p["period_from"])) or "") and (nxt > this_month or nxt in months)
        mine = den.get(p["pk"], {})
        sheet_prime = {}
        for name, v in json.loads(p["prime"]).items():
            lender = SHEET_PRIME.get(re.sub(r"\W", "", name.lower()))
            if lender:
                sheet_prime[lender] = sheet_prime.get(lender, 0) + v
        groups = [(lender, sheet_prime.get(lender, 0.0), [lender], "prime_collected_amount")
                  for lender in sorted(lender_journal.PRIME) if lender in sheet_prime or lender in mine]
        sub_lenders = [l for l in mine if l not in lender_journal.PRIME]
        if p["subprime"] or sub_lenders:
            groups.append(("Subprime", p["subprime"], sub_lenders, "subprime_collected_amount"))
        # Compare each lender (prime) or all subprime together. The sheet should have the amount before the fee
        # (payment + fee); the amount after it (the payment alone) also counts as a match.
        close = lambda a, b: abs(a - b) <= 1
        found = []
        for label, sheet_v, lenders, col in groups:
            pay = sum(mine.get(l, (0, 0))[0] for l in lenders)
            fee = sum(mine.get(l, (0, 0))[1] for l in lenders)
            # Cash or another payment on the row ("Other form of payment") may be included in $Amt Coll.
            extra = p["other"] or 0
            if close(sheet_v, pay + fee) or (pay and close(sheet_v, pay)) or \
                    (extra and pay and (pay - 1 <= sheet_v - extra <= pay + fee + 1)):
                continue
            found.append((label, sheet_v, lenders, col, round(pay, 2), round(fee, 2)))
        if not found:
            continue
        key = lambda label: f"den:{p['row']}:{label}"
        sheet_total = sum(sheet_prime.values()) + p["subprime"]
        pay_total = sum(v[0] for v in mine.values())
        fee_total = sum(v[1] for v in mine.values())
        # The same total under a different lender (or the fee booked under another lender): one note, not two errors.
        totals = [sheet_total, sheet_total - (p["other"] or 0)]
        if sheet_total and any(close(t, pay_total + fee_total) or close(t, pay_total) for t in totals):
            sheet_side = ", ".join(f"{l} {money(v)}" for l, v in [*sheet_prime.items(), ("Subprime", p["subprime"])] if v)
            den_side = ", ".join(f"{l} {money(v[0] + v[1])}" for l, v in mine.items() if v[0] or v[1])
            out.append({"office": p["office"], "period_from": p["period_from"], "check_key": key("lender"),
                        "value_sig": f"{sheet_side}|{den_side}", "col": found[0][3], "kind": "lender",
                        "problem": f"Row {p['row']}: right amount, different lender. The sheet has {sheet_side}; "
                                   f"Denticon shows {den_side}."})
            continue
        for label, sheet_v, lenders, col, pay, fee in found:
            d = round(pay + fee, 2)
            if sheet_v > d and not full:
                continue
            via = ", ".join(lenders) if label == "Subprime" and lenders else label
            seen = f"Denticon shows {money(d)} ({via}: {money(pay)} paid + {money(fee)} fee)" if d else "Denticon shows no payment"
            if 0 < sheet_v < 50 and not d:
                kind, problem = "odd", (f"Row {p['row']}: {money(sheet_v)} in the {label} $Amt Coll column and no payment "
                                        f"in Denticon; that doesn't look like a dollar amount.")
            elif not sheet_v:
                kind, problem = "missing_sheet", f"Row {p['row']}: the sheet has no {label} collected, but {seen}."
            elif not d:
                kind, problem = "missing_denticon", f"Row {p['row']}: the sheet says {money(sheet_v)} {label} collected; {seen}."
            else:
                kind, problem = "amount", f"Row {p['row']}: the sheet says {money(sheet_v)} {label} collected; {seen}."
            out.append({"office": p["office"], "period_from": p["period_from"], "check_key": key(label),
                        "value_sig": f"{sheet_v:.2f}|{d:.2f}", "col": col, "kind": kind, "problem": problem})
    return out


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


def backup_db():
    """Today's copy of the log, if there isn't one yet: data/backups/reports-YYYY-MM-DD.db. Uses SQLite's
    backup API (safe while the app is writing), checks the copy, then keeps only the newest KEEP_BACKUPS."""
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    target = BACKUP_DIR / f"reports-{date.today().isoformat()}.db"
    if target.exists() or not DB_PATH.exists():
        return None
    tmp = target.with_name(target.name + ".tmp")
    src, dst = sqlite3.connect(DB_PATH, timeout=30), sqlite3.connect(tmp)
    try:
        src.backup(dst)
        ok = dst.execute("PRAGMA integrity_check").fetchone()[0] == "ok"
    finally:
        dst.close()
        src.close()
    if not ok:
        tmp.unlink(missing_ok=True)
        raise RuntimeError("the backup copy failed its integrity check")
    tmp.replace(target)
    for old in sorted(BACKUP_DIR.glob("reports-*.db"))[:-KEEP_BACKUPS]:
        old.unlink()
    return target


def backup_loop():
    """Check every hour; the first check each day makes that day's backup."""
    while True:
        try:
            made = backup_db()
            if made:
                print(f"{datetime.now():%H:%M} Backup saved: {made.name}")
        except Exception as e:  # never stop the app over a backup
            print(f"{datetime.now():%H:%M} Backup failed: {e}")
        time.sleep(3600)


# ------------------------------------------------------------------ notes sheet (the Log, copied to Google)

OLD_LOG_TAB, ORPHAN_TAB = "Log", "Notes (orphaned)"   # "Log": the one-tab layout used before month tabs
MONTH_TAB_RE = re.compile(r"^(January|February|March|April|May|June|July|August|September|October|November|December) \d{4}$")
notes_lock = threading.Lock()


def log_sheet_columns():
    """The Log's columns, as the page shows them: [(key, label, kind, ratio)] from LOG_COLS in static/index.html,
    in the page's section order, without internal, text and period columns. ratio = (numerator keys, denominator)
    for rates and averages, so the Total row is calculated the same way as the page's."""
    page = (ROOT / "static" / "index.html").read_text()
    try:
        block = page[page.index("const LOG_COLS = ["):page.index("// Rates that can't")]
        sections = page[page.index("const SECTIONS = ["):page.index("const sectionOf")]
    except ValueError:
        return [(m, m, "count", None) for m in METRIC_NAMES]
    cols = []
    for chunk in re.split(r"\n  (?=\{key: \")", block):
        m = re.match(r'\{key: "(\w+)", label: "([^"]+)", kind: "(\w+)"', chunk)
        if not m or m.group(3) not in ("money", "count", "pct") or "internal: true" in chunk:
            continue
        r = re.search(r'ratio: \{num: \[([^\]]*)\], den: "(\w+)"\}', chunk)
        ratio = (re.findall(r'"(\w+)"', r.group(1)), r.group(2)) if r else None
        cols.append((m.group(1), m.group(2), m.group(3), ratio))
    order = [k for keys in re.findall(r"keys: \[([^\]]+)\]", sections) for k in re.findall(r'"(\w+)"', keys)]
    rank = {k: i for i, k in enumerate(order)}
    return sorted(cols, key=lambda c: rank.get(c[0], len(rank)))


def _total(rows, key, kind, ratio):
    """The Total row's value, as the page calculates it: rates from the summed parts, amounts and counts added up."""
    if ratio:
        num, den = ratio
        full = [r for r in rows if r[den] and all(r[k] is not None for k in num)]
        base = sum(r[den] for r in full)
        return round(sum(sum(r[k] for k in num) for r in full) / base, 4) if base else ""
    vals = [r[key] for r in rows if r[key] is not None]
    return round(sum(vals), 2) if vals and kind in ("money", "count") else ""


def notes_fingerprint(conn):
    """Changes whenever anything in the Log could have changed (cheap to check every couple of minutes)."""
    q = lambda sql: conn.execute(sql).fetchone()[0]
    return "|".join(str(v) for v in (q("SELECT MAX(id) FROM change_log"), q("SELECT COUNT(*) FROM monthly_log"),
                                     q("SELECT MAX(logged_at) FROM office_metrics"), q("SELECT MAX(logged_at) FROM office_summary"),
                                     q("SELECT COUNT(*) FROM office_metrics"), q("SELECT COUNT(*) FROM office_summary"),
                                     q("SELECT MAX(logged_at) FROM office_daily"), q("SELECT MAX(logged_at) FROM daily_metrics"),
                                     q("SELECT COUNT(*) FROM daily_imports")))


def notes_sheet_info():
    conn = db()
    row = conn.execute("SELECT * FROM notes_sheet WHERE id = 1").fetchone()
    conn.close()
    return {"service_account": account.email(), **({k: row[k] for k in row.keys() if k != "last_fingerprint"} if row else {})}


def _notes_from(rows, month=None):
    """{(office, month): note} from a tab's rows (header row first). Month tabs have no Month column."""
    if not rows:
        return {}
    head = [h.strip().lower() for h in rows[0]]
    if "office" not in head or "notes" not in head:
        return {}
    io, inote, im = head.index("office"), head.index("notes"), head.index("month") if "month" in head else None
    out = {}
    for r in rows[1:]:
        note = r[inote].strip() if len(r) > inote else ""
        office = r[io].strip() if len(r) > io else ""
        m = month or (r[im].strip() if im is not None and len(r) > im else "")
        for fmt in ("%b %Y", "%B %Y"):   # the old one-tab layout said "Sep 2026"; month tabs are "September 2026"
            try:
                m = datetime.strptime(m, fmt).strftime("%B %Y")
                break
            except ValueError:
                pass
        if note and office and office != "Total" and m:
            out[(office, m)] = note
    return out


def write_notes_sheet(force=False):
    """Copy the Log to the notes sheet, one tab per month like the platform ("October 2026", newest first),
    keeping what people typed in each tab's Notes column (matched by office). Returns a status, or None
    when nothing changed."""
    if not notes_lock.acquire(timeout=300):
        return None
    try:
        conn = db()
        row = conn.execute("SELECT * FROM notes_sheet WHERE id = 1").fetchone()
        if not row:
            conn.close()
            return None
        fp = notes_fingerprint(conn)
        if not force and fp == row["last_fingerprint"]:
            conn.close()
            return None
        cols = log_sheet_columns()
        log = sorted(log_rows(conn, [m + "_detail" for m in DETAIL_METRICS]), key=lambda r: (r["period_from"], r["office"]))
        log.sort(key=lambda r: r["period_from"], reverse=True)   # newest month first, offices A-Z inside
        conn.close()
        tab_of = lambda pf: datetime.strptime(pf[:7], "%Y-%m").strftime("%B %Y")
        months = list(dict.fromkeys(tab_of(r["period_from"]) for r in log))           # newest first
        try:
            wb = google_sheets.SheetsWorkbook(account, row["spreadsheet_id"])
            ours = [t for t in wb.sheets if MONTH_TAB_RE.match(t) or t == OLD_LOG_TAB]
            old = wb.values_batch_get(ours)
            notes = {}
            for t, rows in old.items():
                notes.update(_notes_from(rows, None if t == OLD_LOG_TAB else t))
            # Tabs: rename a brand-new sheet's empty "Sheet1", add missing months, drop month tabs (and the old
            # one-tab "Log") that the log no longer has, then put the months first, newest first.
            reqs = []
            if "Sheet1" in wb.sheets and len(wb.sheets) == 1 and months and not wb.values_get("Sheet1"):
                reqs.append({"updateSheetProperties": {"properties": {"sheetId": wb.sheet_ids["Sheet1"], "title": months[0]},
                                                       "fields": "title"}})
                wb.sheet_ids = {months[0]: wb.sheet_ids.pop("Sheet1")}
                wb.sheets = {months[0]: months[0]}
            reqs += [{"addSheet": {"properties": {"title": m}}} for m in months if m not in wb.sheets]
            if reqs:
                wb.batch_update(reqs)
            gone = [t for t in ours if t not in months]
            kept = {(r["office"], tab_of(r["period_from"])) for r in log}
            orphans = [[o, m, n] for (o, m), n in notes.items() if (o, m) not in kept]
            now = datetime.now().strftime("%Y-%m-%d %H:%M")
            if orphans:   # a row that's gone from the log: keep its note on another tab rather than lose it
                wb.ensure_tab(ORPHAN_TAB)
                prev = wb.values_get(ORPHAN_TAB)
                have = {tuple(x[:3]) for x in prev[1:]}
                new = [o + [now] for o in orphans if tuple(o) not in have]
                if new:
                    wb.values_update(ORPHAN_TAB, ([["Office", "Month", "Notes", "Moved here"]] if not prev else []) + new,
                                     start=f"A{len(prev) + 1}")
            header = ["Office", "Thru", *[c[1] for c in cols], "Notes", "Updated"]
            data = {}
            for m in months:
                rows = [r for r in log if tab_of(r["period_from"]) == m]
                body = [[r["office"], r["period_thru"] or "", *[("" if r[k] is None else r[k]) for k, _, _, _ in cols],
                         notes.get((r["office"], m), ""), now] for r in rows]
                total = ["Total", f"{len(rows)} offices", *[_total(rows, k, kind, ratio) for k, _, kind, ratio in cols], "", ""]
                data[m] = [header, *body, total]
            wb.values_batch_clear(months)
            wb.values_batch_update(data)
            # Formatting for new tabs or changed columns, plus tab order and removing tabs no longer needed.
            fmt = {"money": {"type": "CURRENCY", "pattern": "$#,##0.00"}, "pct": {"type": "PERCENT", "pattern": "0.00%"},
                   "count": {"type": "NUMBER", "pattern": "#,##0"}}
            reqs = []
            for m in months:
                if old.get(m) and [h.strip() for h in old[m][0]] == header:
                    continue
                sid = wb.sheet_ids[m]
                reqs += [{"updateSheetProperties": {"properties": {"sheetId": sid, "gridProperties": {
                             "frozenRowCount": 1, "frozenColumnCount": 1}}, "fields": "gridProperties(frozenRowCount,frozenColumnCount)"}},
                         {"repeatCell": {"range": {"sheetId": sid, "startRowIndex": 0, "endRowIndex": 1},
                                         "cell": {"userEnteredFormat": {"textFormat": {"bold": True}, "wrapStrategy": "WRAP"}},
                                         "fields": "userEnteredFormat(textFormat,wrapStrategy)"}}]
                reqs += [{"repeatCell": {"range": {"sheetId": sid, "startRowIndex": 1, "startColumnIndex": i, "endColumnIndex": i + 1},
                                         "cell": {"userEnteredFormat": {"numberFormat": fmt[c[2]]}},
                                         "fields": "userEnteredFormat.numberFormat"}} for i, c in enumerate(cols, start=2)]
            reqs += [{"deleteSheet": {"sheetId": wb.sheet_ids[t]}} for t in gone if t in wb.sheet_ids]
            reqs += [{"updateSheetProperties": {"properties": {"sheetId": wb.sheet_ids[m], "index": i}, "fields": "index"}}
                     for i, m in enumerate(months)]
            if reqs:
                wb.batch_update(reqs)
            status = (f"ok: {len(months)} month tab(s), {len(log)} rows written"
                      + (f", {len(orphans)} note(s) moved to '{ORPHAN_TAB}'" if orphans else ""))
            done = {"last_fingerprint": fp, "title": wb.title}
        except google_sheets.GoogleError as e:
            status, done = f"error: {e}", {}
        conn = db()
        with conn:
            conn.execute("UPDATE notes_sheet SET last_write_at = ?, last_status = ?, last_fingerprint = COALESCE(?, last_fingerprint), "
                         "title = COALESCE(?, title) WHERE id = 1", (datetime.now().isoformat(timespec="seconds"), status,
                                                                    done.get("last_fingerprint"), done.get("title")))
        conn.close()
        return status
    finally:
        notes_lock.release()


def connect_notes_sheet(url):
    sid = google_sheets.spreadsheet_id(url)
    wb = google_sheets.SheetsWorkbook(account, sid)   # checks the link and read access
    conn = db()
    with conn:
        conn.execute("INSERT OR REPLACE INTO notes_sheet (id, spreadsheet_id, url, title) VALUES (1, ?, ?, ?)", (sid, url, wb.title))
    conn.close()
    status = write_notes_sheet(force=True)
    if status and status.startswith("error") and "No access" in status:
        disconnect_notes_sheet()
        raise ValueError(f"The app can't edit that sheet. Share it with {account.email()} as Editor.")
    return {"status": status, **notes_sheet_info()}


def disconnect_notes_sheet():
    conn = db()
    with conn:
        conn.execute("DELETE FROM notes_sheet")
    conn.close()
    return {"ok": True}


def notes_sheet_loop():
    """Every 2 minutes: if anything in the Log changed, copy it to the notes sheet."""
    while True:
        time.sleep(120)
        try:
            status = write_notes_sheet()
            if status:
                print(f"{datetime.now():%H:%M} Notes sheet: {status}")
        except Exception as e:  # keep the loop alive
            print(f"{datetime.now():%H:%M} Notes sheet failed: {e}")


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
    if not office or not re.fullmatch(r"\d{4}-\d{2}-01", period_from) or not key.startswith(("rate:", "row:", "den:")) or not sig:
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


denticon = denticon_api.Client(DATA_DIR)
denticon_state = {"last_test": None, "last_error": None}
invoca = invoca_api.Client(DATA_DIR)
invoca_state = {"last_test": None, "last_error": None, "last_result": None}


def invoca_status():
    return {"connected": invoca.exists(), "config": invoca.config(), **invoca_state}


def _int(v):
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return None


def store_invoca_calls(day, transactions):
    """Keep the day's calls from a Transactions API answer. Signals come from each call's custom_data, plus any
    Signal transactions for the same call (complete_call_id). Returns how many calls were stored."""
    sig_by_call = {}
    for t in transactions:
        if (t.get("transaction_type") or "") == "Signal":
            for d in t.get("custom_data") or []:
                if isinstance(d, dict) and d.get("data_type") == "Signal" and d.get("name"):
                    sig_by_call.setdefault(t.get("complete_call_id"), {})[d["name"]] = bool(d.get("value") in (True, "true", "True", 1, "1"))
    now = datetime.now().isoformat(timespec="seconds")
    rows = []
    for t in transactions:
        if (t.get("transaction_type") or "Call") != "Call" or not t.get("transaction_id"):
            continue
        sigs, custom = dict(sig_by_call.get(t.get("complete_call_id"), {})), {}
        for d in t.get("custom_data") or []:
            if not isinstance(d, dict) or not d.get("name"):
                continue
            if d.get("data_type") == "Signal":
                if d.get("value") is not None:
                    sigs[d["name"]] = bool(d.get("value") in (True, "true", "True", 1, "1"))
            elif d.get("value") not in (None, ""):
                custom[d["name"]] = str(d["value"])[:200]
        rows.append((t["transaction_id"], t.get("complete_call_id"), day, t.get("start_time_local"), _int(t.get("start_time_utc")),
                     _int(t.get("duration")), _int(t.get("connect_duration")), _int(t.get("ivr_duration")), str(t.get("advertiser_id") or ""),
                     t.get("advertiser_name"), t.get("advertiser_campaign_name"), t.get("promo_line_description"), str(t.get("virtual_line_id") or ""),
                     t.get("city"), t.get("region"), t.get("mobile"), 1 if str(t.get("repeat_calling_phone_number")).lower() in ("true", "1", "yes") else 0,
                     t.get("call_result_description_detail"), json.dumps(sorted(k for k, v in sigs.items() if v)), json.dumps(custom), now,
                     phone_fp(t.get("calling_phone_number"))))
    conn = db()
    with conn:
        conn.execute("DELETE FROM invoca_calls WHERE call_date = ?", (day,))
        conn.executemany("INSERT OR REPLACE INTO invoca_calls VALUES (" + ",".join("?" * 22) + ")", rows)
    conn.close()
    return len(rows)


def pull_invoca_calls(day):
    """Fetch one day's calls from Invoca and keep them. The day is YYYY-MM-DD, not in the future."""
    from datetime import date
    try:
        d = date.fromisoformat(day)
    except (TypeError, ValueError):
        raise ValueError("Pick a date (YYYY-MM-DD).")
    if d > date.today():
        raise ValueError("That day hasn't happened yet.")
    transactions = invoca.all_transactions(day, day)
    store_invoca_calls(day, transactions)
    info = calls_info(day)
    return {"date": day, "calls": len(info["rows"]), "transactions": len(transactions), **info}


def pull_invoca_range(frm, to):
    """Backfill: every call from `frm` to `to` (YYYY-MM-DD, inclusive) in one paged fetch, kept day by day."""
    from datetime import date
    d0, d1 = date.fromisoformat(frm), date.fromisoformat(to)
    if d1 < d0 or d1 > date.today():
        raise ValueError("Pick a range of past days, oldest first.")
    transactions = invoca.all_transactions(frm, to, max_pages=400)
    by_day = {}
    for t in transactions:
        day = str(t.get("start_time_local") or "")[:10]
        if len(day) == 10:
            by_day.setdefault(day, []).append(t)
    stored = {day: store_invoca_calls(day, tx) for day, tx in sorted(by_day.items())}
    return {"from": frm, "to": to, "transactions": len(transactions), "days": stored}


def promo_line_map(conn):
    """{line: {office, label}} as typed in on the Calls tab."""
    return {r["line"]: {"office": r["office"], "label": r["label"]} for r in conn.execute("SELECT line, office, label FROM promo_lines")}


def promo_lines_info():
    """Every promo line seen in the pulled calls, with its call count, last call and the office/label given to it."""
    conn = db()
    m = promo_line_map(conn)
    lines = [{"line": r["line"], "calls": r["n"], "last": r["last"], **m.get(r["line"], {"office": "", "label": ""})}
             for r in conn.execute("SELECT COALESCE(promo_line, '') AS line, COUNT(*) AS n, MAX(call_date) AS last FROM invoca_calls GROUP BY 1 ORDER BY n DESC")]
    offices = known_offices(conn)
    short = office_short_names(conn)
    conn.close()
    mapped = sum(1 for l in lines if l["office"])
    return {"lines": lines, "offices": offices, "short": short, "mapped": mapped}


def set_promo_line(c):
    line = str(c.get("line") or "")
    office, label = str(c.get("office") or "").strip(), str(c.get("label") or "").strip()[:80]
    conn = db()
    if office and office not in known_offices(conn):
        conn.close()
        raise ValueError("Pick one of the Log's offices")
    with conn:
        conn.execute("INSERT INTO promo_lines (line, office, label, updated_at) VALUES (?, ?, ?, ?) ON CONFLICT(line) DO UPDATE SET office = excluded.office, label = excluded.label, updated_at = excluded.updated_at",
                     (line, office, label, datetime.now().isoformat(timespec="seconds")))
    conn.close()
    return {"line": line, "office": office, "label": label}


def pull_invoca_month(month):
    """Backfill a whole month (up to today) in one paged fetch."""
    from datetime import date
    try:
        y, mo = (int(x) for x in str(month).split("-")[:2])
        first = date(y, mo, 1)
    except (TypeError, ValueError):
        raise ValueError("Pick a month (YYYY-MM).")
    if first > date.today():
        raise ValueError("That month hasn't started yet.")
    last = min(date(y + (mo == 12), (mo % 12) + 1, 1) - __import__("datetime").timedelta(days=1), date.today())
    r = pull_invoca_range(first.isoformat(), last.isoformat())
    return {**r, "month": month, "calls": sum(r["days"].values()), "day_count": len(r["days"])}


def calls_funnel(month=None, office=""):
    """Calls in a month -> callers who became new patients (a loaded patient list says their first visit was on or
    after the call) -> who started treatment (paid anything in their first-visit month) -> $ collected. Each
    patient counts once, on their first qualifying call, which also decides the promo line / campaign credited.
    Calls can't be split by office (Invoca reports the group), so the office filter applies from the patient step."""
    conn = db()
    months = [r[0] for r in conn.execute("SELECT DISTINCT substr(call_date, 1, 7) FROM invoca_calls ORDER BY 1 DESC")]
    month = month if month in months else (months[0] if months else None)
    lists = [dict(r) for r in conn.execute("SELECT office, COUNT(DISTINCT pat_id) AS patients, MIN(first_visit) AS lo, MAX(first_visit) AS hi FROM patient_phones GROUP BY office ORDER BY office")]
    if not month:
        conn.close()
        return {"month": None, "months": [], "lists": lists, "offices": [l["office"] for l in lists]}
    lo, hi = f"{month}-01", f"{month}-31"
    pmap = promo_line_map(conn)
    calls = merge_call_legs([dict(r) for r in conn.execute("SELECT transaction_id, complete_call_id, call_date, caller_fp, connect_duration, duration, promo_line, campaign, signals, custom FROM invoca_calls "
                                                            "WHERE call_date BETWEEN ? AND ? ORDER BY call_date, start_utc", (lo, hi))])
    answered_rule = answered_flag(calls)
    canon = {}   # the same value in different spellings ("paid search" / "Paid Search") counts as one
    def tidy(v):
        v = str(v or "").strip()
        if not v:
            return "(not set)"
        return canon.setdefault(v.lower(), v)
    for c in calls:
        pm = pmap.get(c["promo_line"] or "", {})
        c["office"] = pm.get("office") or "(not mapped)"
        c["promo_line"] = pm.get("label") or c["promo_line"]
        cus = c.get("custom") or {}
        c["channel"] = tidy(cus.get("marketing_channel"))     # Invoca's marketing channel: Paid Search, Organic Search, ...
        c["medium"] = tidy(cus.get("utm_medium"))             # Google PPC, Organic Search, paid search, direct, ...
        c["source"] = tidy(cus.get("utm_source"))             # Google Ad Call Extension, Google My Business, Yelp, ...
        c["location"] = tidy(cus.get("location"))             # the office location Invoca routed the call to
    office_calls = office and any(v["office"] == office for v in pmap.values())   # lines are mapped to this office: filter calls too
    if office_calls:
        calls = [c for c in calls if c["office"] == office]
    # Patients reached: for every caller fingerprint, the patient records (first visit on/after the call) and appointments.
    fps = sorted({c["caller_fp"] for c in calls if c["caller_fp"]})
    pats, appts = {}, {}
    for i in range(0, len(fps), 500):
        chunk = fps[i:i + 500]; q = ",".join("?" * len(chunk))
        for r in conn.execute(f"SELECT p.fp, p.pat_id, p.office, p.first_visit, "
                              f"(SELECT ROUND(SUM(m.collection), 2) FROM map_patients m WHERE m.pat_id = p.pat_id AND m.office = p.office "
                              f" AND p.first_visit IS NOT NULL AND m.period_from = substr(p.first_visit, 1, 7) || '-01') AS collected "
                              f"FROM patient_phones p WHERE p.fp IN ({q})" + (" AND p.office = ?" if office else ""), chunk + ([office] if office else [])):
            pats.setdefault(r["fp"], []).append(dict(r))
        for r in conn.execute(f"SELECT fp, office, appt_date FROM np_appts WHERE fp IN ({q})" + (" AND office = ?" if office else ""), chunk + ([office] if office else [])):
            appts.setdefault(r["fp"], []).append(dict(r))
    conn.close()
    seen_pat, seen_booked = set(), set()
    GROUPS = ("promo_line", "campaign", "office", "channel", "medium", "source", "location")
    groups = {g: {} for g in GROUPS}
    def bump(c, key, **inc):
        for g in GROUPS:
            d = groups[g].setdefault(c[g] or "(none)", {"calls": 0, "answered": 0, "patients": 0, "started": 0, "collected": 0.0, "booked": 0})
            for k, v in inc.items():
                d[k] += v
    total = {"calls": 0, "answered": 0, "callers": len(fps), "patients": 0, "started": 0, "collected": 0.0, "booked": 0}
    for c in calls:
        answered = 1 if c["answered"] else 0
        total["calls"] += 1; total["answered"] += answered
        bump(c, None, calls=1, answered=answered)
        fp = c["caller_fp"]
        if not fp:
            continue
        for p in pats.get(fp, []):
            key = (p["pat_id"], p["office"])
            if key in seen_pat or not p["first_visit"] or p["first_visit"] < c["call_date"]:
                continue
            seen_pat.add(key)
            started = 1 if (p["collected"] or 0) > 0 else 0
            total["patients"] += 1; total["started"] += started; total["collected"] += p["collected"] or 0
            bump(c, None, patients=1, started=started, collected=p["collected"] or 0)
        if fp not in {k for k, _ in seen_pat} and fp not in seen_booked:
            if any(a["appt_date"] and a["appt_date"] >= c["call_date"] for a in appts.get(fp, [])) and not any(p["first_visit"] and p["first_visit"] >= c["call_date"] for p in pats.get(fp, [])):
                seen_booked.add(fp); total["booked"] += 1; bump(c, None, booked=1)
    total["collected"] = round(total["collected"], 2)
    out_groups = {g: sorted(({"name": k, **{kk: (round(vv, 2) if kk == "collected" else vv) for kk, vv in v.items()}} for k, v in d.items()),
                            key=lambda x: (-x["patients"], -x["calls"])) for g, d in groups.items()}
    mapped_offices = sorted({v["office"] for v in pmap.values() if v["office"]})
    return {"month": month, "months": months, "office": office, "office_calls": bool(office_calls), "lists": lists,
            "offices": sorted({l["office"] for l in lists} | set(mapped_offices)), "mapped_lines": sum(1 for v in pmap.values() if v["office"]), "total": total, "by": out_groups,
            "answered_rule": answered_rule}


AGENT_SIGNAL = "Answered by Agent"   # Invoca's own "a person picked up" signal; its dashboard counts answered calls by it


def merge_call_legs(rows):
    """One record per call. Invoca sends one transaction per leg (an overflow or transfer re-reports the same call
    under the same complete_call_id), and its own dashboard counts the call once. The first leg gives the time,
    promo line and campaign (what the caller dialled); durations are the longest leg's; signals are combined."""
    by, order = {}, []
    for r in rows:
        key = r.get("complete_call_id") or r["transaction_id"]
        sig = set(json.loads(r["signals"] or "[]")) if isinstance(r.get("signals"), str) else set(r.get("signals") or [])
        cus = json.loads(r["custom"] or "{}") if isinstance(r.get("custom"), str) else dict(r.get("custom") or {})
        if key not in by:
            m = by[key] = dict(r); m["legs"] = 1; m["_sig"] = sig; m["_cus"] = cus; order.append(key)
            continue
        m = by[key]; m["legs"] += 1
        m["duration"] = max(m.get("duration") or 0, r.get("duration") or 0)
        m["connect_duration"] = max(m.get("connect_duration") or 0, r.get("connect_duration") or 0)
        m["_sig"] |= sig; m["_cus"] = {**cus, **m["_cus"]}
        if not m.get("caller_fp") and r.get("caller_fp"):
            m["caller_fp"] = r["caller_fp"]
        if not m.get("result") and r.get("result"):
            m["result"] = r["result"]
    out = []
    for k in order:
        m = by[k]; m["signals"] = sorted(m.pop("_sig")); m["custom"] = m.pop("_cus"); out.append(m)
    return out


def answered_flag(calls):
    """Mark each merged call answered: by Invoca's own signal when the account has it, else by any connected time."""
    uses_signal = any(AGENT_SIGNAL in c["signals"] for c in calls)
    for c in calls:
        c["answered"] = (AGENT_SIGNAL in c["signals"]) if uses_signal else ((c.get("connect_duration") or 0) > 0)
    return "signal" if uses_signal else "connect"


def calls_info(day=None):
    """The calls kept for a day, plus which days have been pulled."""
    conn = db()
    days = [dict(r) for r in conn.execute("SELECT call_date AS day, COUNT(DISTINCT COALESCE(complete_call_id, transaction_id)) AS n, MAX(pulled_at) AS pulled_at FROM invoca_calls GROUP BY call_date ORDER BY call_date DESC")]
    rows = merge_call_legs([dict(r) for r in conn.execute("SELECT * FROM invoca_calls WHERE call_date = ? ORDER BY start_utc, start_local", (day,))]) if day else []
    answered_rule = answered_flag(rows)
    pmap = promo_line_map(conn)
    for r in rows:
        pm = pmap.get(r["promo_line"] or "", {})
        r["office"] = pm.get("office") or ""
        r["promo_label"] = pm.get("label") or ""
    # Which callers became patients (by phone fingerprint): a patient record whose first visit is on or after the
    # call = came in (with the $ collected that month from the Referral listing); a new-patient appointment on or
    # after the call = booked; an earlier first visit = an existing patient.
    appts, pats = {}, {}
    if day:
        for a in conn.execute("SELECT fp, office, appt_date FROM np_appts WHERE fp IN (SELECT caller_fp FROM invoca_calls WHERE call_date = ? AND caller_fp != '')", (day,)):
            appts.setdefault(a["fp"], []).append({"office": a["office"], "appt_date": a["appt_date"]})
        for a in conn.execute("SELECT p.fp, p.pat_id, p.office, p.first_visit, "
                              "(SELECT ROUND(SUM(m.collection), 2) FROM map_patients m WHERE m.pat_id = p.pat_id AND m.office = p.office "
                              " AND p.first_visit IS NOT NULL AND m.period_from = substr(p.first_visit, 1, 7) || '-01') AS collected "
                              "FROM patient_phones p WHERE p.fp IN (SELECT caller_fp FROM invoca_calls WHERE call_date = ? AND caller_fp != '')", (day,)):
            pats.setdefault(a["fp"], []).append({"pat_id": a["pat_id"], "office": a["office"], "first_visit": a["first_visit"], "collected": a["collected"]})
    appts_known = conn.execute("SELECT COUNT(DISTINCT fp) FROM np_appts").fetchone()[0] + conn.execute("SELECT COUNT(DISTINCT fp) FROM patient_phones").fetchone()[0]
    conn.close()
    for r in rows:
        fp = r.pop("caller_fp", "") or ""
        phits, hits = pats.get(fp, []), appts.get(fp, [])
        visited = sorted([h for h in phits if h["first_visit"] and h["first_visit"] >= day], key=lambda h: h["first_visit"])
        existing = sorted([h for h in phits if not h["first_visit"] or h["first_visit"] < day], key=lambda h: h["first_visit"] or "")
        after = sorted([h for h in hits if h["appt_date"] and h["appt_date"] >= day], key=lambda h: h["appt_date"])
        before = sorted([h for h in hits if not h["appt_date"] or h["appt_date"] < day], key=lambda h: h["appt_date"] or "")
        r["match"] = ({"status": "visited", **visited[0]} if visited else {"status": "converted", **after[0]} if after
                      else {"status": "existing", **existing[-1]} if existing else {"status": "before", **before[-1]} if before else None)
    # Where each caller's city is, for the ones in the map area's Census city list (name match, California only).
    cities = {name.lower(): v["c"] for name, v in city_names().get("cities", {}).items() if v.get("c")}
    points = {}
    for r in rows:
        c, st = (r.get("city") or "").strip(), (r.get("region") or "").strip()
        key = f"{c}|{st}"
        if c and key not in points and st.upper() in ("CA", "CALIFORNIA", "") and c.lower() in cities:
            points[key] = cities[c.lower()]
    return {"date": day, "rows": rows, "days": days, "connected": invoca.exists(), "points": points, "appts_known": appts_known,
            "mapped_lines": sum(1 for v in pmap.values() if v["office"]), "answered_rule": answered_rule,
            "legs": sum(r["legs"] for r in rows)}


def invoca_test():
    """Read the last 7 days of calls and summarize them by office (advertiser) and campaign: nothing about callers."""
    from datetime import date, timedelta
    to, frm = date.today(), date.today() - timedelta(days=6)
    try:
        rows = invoca.all_transactions(frm.isoformat(), to.isoformat())
    except invoca_api.InvocaError as e:
        invoca_state.update(last_error=str(e), last_test=datetime.now().isoformat(timespec="seconds"), last_result=None)
        raise
    calls = [r for r in rows if (r.get("transaction_type") or "Call") == "Call"]
    by_adv, by_camp = {}, {}
    for r in calls:
        a = r.get("advertiser_name") or f"advertiser {r.get('advertiser_id')}"
        by_adv[a] = by_adv.get(a, 0) + 1
        k = f"{a} · {r.get('advertiser_campaign_name') or 'no campaign'}"
        by_camp[k] = by_camp.get(k, 0) + 1
    signals = sorted({d.get("name") for r in calls for d in (r.get("custom_data") or []) if isinstance(d, dict) and d.get("data_type") == "Signal" and d.get("name")})
    result = {"from": frm.isoformat(), "to": to.isoformat(), "transactions": len(rows), "calls": len(calls),
              "by_advertiser": sorted(by_adv.items(), key=lambda kv: -kv[1]), "by_campaign": sorted(by_camp.items(), key=lambda kv: -kv[1])[:30],
              "signals": signals[:40], "columns": sorted(rows[0].keys()) if rows else []}
    invoca_state.update(last_error=None, last_test=datetime.now().isoformat(timespec="seconds"), last_result=result)
    return {**invoca_status(), "result": result}


def _pick(d, *names):
    for n in names:
        if d.get(n) not in (None, ""):
            return d[n]
    return None


def denticon_status():
    conn = db()
    offices = [dict(r) for r in conn.execute("SELECT * FROM denticon_offices ORDER BY name")]
    known = known_offices(conn)
    conn.close()
    return {"key_present": denticon.exists(), "offices": offices, "log_offices": known, **denticon_state}


def denticon_test():
    """Check the key by listing the practice's offices (names and IDs only), and match them to the log."""
    try:
        rows = denticon.offices()
    except denticon_api.DenticonError as e:
        denticon_state.update(last_error=str(e), last_test=datetime.now().isoformat(timespec="seconds"))
        raise ValueError(str(e))
    conn = db()
    known = {normalize(o): o for o in known_offices(conn)}
    aliases = {r["alias"]: r["office"] for r in conn.execute("SELECT alias, office FROM office_alias")}
    now = datetime.now().isoformat(timespec="seconds")
    with conn:
        for r in rows:
            oid = _pick(r, "officeId", "id", "oid", "officeID")
            name = str(_pick(r, "officeName", "name", "officeDescription", "description") or oid)
            if oid is None:
                continue
            old = conn.execute("SELECT log_office FROM denticon_offices WHERE office_id = ?", (oid,)).fetchone()
            match = (old and old["log_office"]) or known.get(normalize(name)) or aliases.get(normalize(name))
            conn.execute("INSERT OR REPLACE INTO denticon_offices VALUES (?, ?, ?, ?)", (int(oid), name, match, now))
    conn.close()
    denticon_state.update(last_error=None, last_test=now,
                          fields=sorted({k for r in rows for k in r})[:40])   # field names only, to plan the sync
    return {**denticon_status(), "found": len(rows)}


def set_denticon_office(c):
    conn = db()
    with conn:
        n = conn.execute("UPDATE denticon_offices SET log_office = ? WHERE office_id = ?",
                         ((c.get("log_office") or None), int(c["office_id"]))).rowcount
    conn.close()
    if not n:
        raise ValueError("That Denticon office isn't in the list; click Test connection first")
    return {"ok": True}


# ------------------------------------------------------------------ Maps tab

ZIP_FILE = DATA_DIR / "zip_centroids.json"
POP_FILE = DATA_DIR / "zip_population.json"        # older population download, read as a fallback only
SHAPES_FILE = DATA_DIR / "zip_shapes.json"          # ZIP -> {"g": GeoJSON geometry, "pop": residents, "c": [lat, lon]}
NO_ZCTA_FILE = DATA_DIR / "zip_no_zcta.json"        # ZIPs the Census has no area for (PO boxes, unique ZIPs): asked once, remembered
CITY_FILE = DATA_DIR / "city_names.json"            # {"cities": {name: {"c": [lat, lon], "pop": int}}, "zip_city": {zip: name}}
# US Census TIGERweb 2020 places: incorporated cities (layer 25) and census-designated places (26), with population.
CITY_SOURCE = "https://tigerweb.geo.census.gov/arcgis/rest/services/TIGERweb/Places_CouSub_ConCity_SubMCD/MapServer/{layer}/query"
_cities = None
# US Census TIGERweb: ZIP Code Tabulation Area boundaries with 2020 population (public, no key needed).
SHAPES_SOURCE = "https://tigerweb.geo.census.gov/arcgis/rest/services/TIGERweb/PUMA_TAD_TAZ_UGA_ZCTA/MapServer/4/query"
_pops = None
_shapes = None


def _load_json_file(path):
    """A JSON data file, or {} if it's missing or unreadable (a half-written file must not break the app)."""
    try:
        return json.loads(path.read_text())
    except (OSError, ValueError):
        return {}


def _write_json_file(path, data):
    """Write the file whole, then swap it into place, so a crash mid-write can't leave a broken file."""
    DATA_DIR.mkdir(exist_ok=True)
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_text(json.dumps(data, separators=(",", ":")))
    os.replace(tmp, path)
ZIP_SOURCE = "https://www2.census.gov/geo/docs/maps-data/data/gazetteer/2023_Gazetteer/2023_Gaz_zcta_national.zip"
GEOCODER = "https://geocoding.geo.census.gov/geocoder/locations/onelineaddress"
_zips = None


def zip_centroids():
    global _zips
    if not _zips and (_zips is None or ZIP_FILE.exists()):   # empty: the file may have been written since (Map setup)
        _zips = _load_json_file(ZIP_FILE)
    return _zips


def zip_population():
    global _pops
    if _pops is None:
        _pops = _load_json_file(POP_FILE)
    return _pops


def zip_shapes():
    global _shapes
    if not _shapes and (_shapes is None or SHAPES_FILE.exists()):   # empty: the file may have been written since
        _shapes = _load_json_file(SHAPES_FILE)
    return _shapes


MAP_PAD = 0.5    # degrees, about 35 miles: how far past the offices the map area reaches


def map_region(conn=None):
    """[[lat0, lon0], [lat1, lon1]]: the area the map shows and boundaries are downloaded for. The offices' box
    padded by MAP_PAD, or (before any office is pinned) a box around where most patients live. None without either."""
    own = conn is None
    conn = conn or db()
    pins = [(r["lat"], r["lon"]) for r in conn.execute("SELECT lat, lon FROM office_locations")]
    data_zips = {r[0] for r in conn.execute("SELECT DISTINCT zip FROM map_patients")}
    if own:
        conn.close()
    if pins:
        pts = pins
    else:
        centers = {**zip_centroids(), **{z: v["c"] for z, v in zip_shapes().items() if v.get("c")}}
        pts = [centers[z] for z in data_zips if z in centers]
        if not pts:
            return None
        lat_m, lon_m = sorted(c[0] for c in pts)[len(pts) // 2], sorted(c[1] for c in pts)[len(pts) // 2]
        pts = [c for c in pts if abs(c[0] - lat_m) < 1.5 and abs(c[1] - lon_m) < 1.5]   # within ~100 mi of the median
    return [[round(min(p[0] for p in pts) - MAP_PAD, 4), round(min(p[1] for p in pts) - MAP_PAD, 4)],
            [round(max(p[0] for p in pts) + MAP_PAD, 4), round(max(p[1] for p in pts) + MAP_PAD, 4)]]


def download_zip_shapes():
    """Every ZIP boundary (with its 2020 population and center point) around the offices, from the US Census: a box
    around the office pins padded by about 25 miles, plus the individual ZIPs of mapped patients who live outside
    it (other counties or states), fetched by ZIP number. Rerun it after adding an office or when new patients
    turn up in ZIPs that have no boundary yet. Nothing about patients is sent: only box coordinates and ZIP numbers."""
    global _shapes, _zips
    conn = db()
    data_zips = {r[0] for r in conn.execute("SELECT DISTINCT zip FROM map_patients")}
    region = map_region(conn)
    conn.close()
    out = dict(zip_shapes())
    centers = {**zip_centroids(), **{z: v["c"] for z, v in out.items() if v.get("c")}}
    if not region:
        raise ValueError("Add the offices' addresses (step 2) or upload a Referral Production Listing first, so the app knows the region.")
    box = (region[0][1], region[0][0], region[1][1], region[1][0])   # lon0, lat0, lon1, lat1
    import math
    x0, y0, x1, y1 = box   # TIGERweb caps one answer at a few thousand ZIPs: ask in pieces of about 1.2 degrees
    nx, ny = max(1, math.ceil((x1 - x0) / 1.2)), max(1, math.ceil((y1 - y0) / 1.2))
    pieces = [(x0 + (x1 - x0) * i / nx, y0 + (y1 - y0) * j / ny, x0 + (x1 - x0) * (i + 1) / nx, y0 + (y1 - y0) * (j + 1) / ny)
              for i in range(nx) for j in range(ny)]
    common = {"outFields": "ZCTA5,POP100,INTPTLAT,INTPTLON", "outSR": "4326", "f": "geojson", "geometryPrecision": "4",
              "maxAllowableOffset": "0.0005"}

    def fetch(params):
        url = SHAPES_SOURCE + "?" + urllib.parse.urlencode(params)
        try:
            with urllib.request.urlopen(url, timeout=180) as resp:
                data = json.loads(resp.read())
        except (urllib.error.URLError, ValueError, TimeoutError) as e:
            raise ValueError(f"Couldn't get ZIP boundaries from the US Census ({getattr(e, 'reason', e)}). Try again in a minute.")
        if data.get("error"):
            raise ValueError(f"The Census boundary service refused the request: {str(data['error'])[:200]}")
        if (data.get("properties") or {}).get("exceededTransferLimit"):
            raise ValueError("The Census boundary service sent back only part of the region; try again.")
        for f in data.get("features", []):
            pr = f.get("properties") or {}
            z = str(pr.get("ZCTA5") or "").strip()
            if not z.isdigit() or not f.get("geometry"):
                continue
            try:
                c = [round(float(pr["INTPTLAT"]), 5), round(float(pr["INTPTLON"]), 5)]
            except (KeyError, TypeError, ValueError):
                c = centers.get(z)
            out[z] = {"g": f["geometry"], "pop": int(pr.get("POP100") or 0), "c": c}

    for piece in pieces:
        fetch({"geometry": ",".join(f"{v:.4f}" for v in piece), "geometryType": "esriGeometryEnvelope", "inSR": "4326",
               "spatialRel": "esriSpatialRelIntersects", **common})
    # Patients' ZIPs outside the box (other counties, other states): fetched by number, 40 at a time.
    far = sorted(z for z in data_zips if z not in out and z.isdigit())
    for i in range(0, len(far), 40):
        fetch({"where": "ZCTA5 IN (" + ",".join(f"'{z}'" for z in far[i:i + 40]) + ")", **common})
    if not out:
        raise ValueError("The Census returned no ZIP boundaries for the region.")
    _write_json_file(SHAPES_FILE, out)
    _shapes = out
    no_zcta = sorted(set(_load_json_file(NO_ZCTA_FILE) or []) | {z for z in far if z not in out})   # asked for, nothing came back
    _write_json_file(NO_ZCTA_FILE, no_zcta)
    conn = db()
    points = {f"office:{r['office']}": (r["lat"], r["lon"]) for r in conn.execute("SELECT office, lat, lon FROM office_locations")}
    conn.close()
    points.update({z: tuple(v["c"]) for z, v in out.items() if v.get("c")})
    n_cities = download_city_names(region, points)
    _zips = {**zip_centroids(), **{z: v["c"] for z, v in out.items() if v.get("c")}}   # center points for the inset/fit
    _write_json_file(ZIP_FILE, _zips)
    return {"zips": len(out), "far": len(far), "cities": n_cities}


def city_names():
    global _cities
    if not _cities and (_cities is None or CITY_FILE.exists()):
        _cities = _load_json_file(CITY_FILE)
    return _cities or {}


def _in_ring(lat, lon, ring):
    """Ray casting: is (lat, lon) inside this GeoJSON ring of [lon, lat] points?"""
    inside = False
    for i in range(len(ring)):
        x1, y1 = ring[i - 1]
        x2, y2 = ring[i]
        if (y1 > lat) != (y2 > lat) and lon < (x2 - x1) * (lat - y1) / (y2 - y1) + x1:
            inside = not inside
    return inside


def _in_geometry(lat, lon, g):
    polys = [g["coordinates"]] if g.get("type") == "Polygon" else g.get("coordinates", []) if g.get("type") == "MultiPolygon" else []
    return any(_in_ring(lat, lon, poly[0]) for poly in polys if poly)


def download_city_names(region, points):
    """City names (incorporated cities and census-designated places, with 2020 population) inside the map area, and
    which city each given point (ZIP centers, office pins) falls in. The city outlines are used once, here, and not
    kept. Only the area's box coordinates are sent to the Census."""
    global _cities
    cities, polys = {}, []
    x0, y0, x1, y1 = region[0][1], region[0][0], region[1][1], region[1][0]
    for layer in (25, 26):
        url = CITY_SOURCE.format(layer=layer) + "?" + urllib.parse.urlencode({
            "geometry": f"{x0:.4f},{y0:.4f},{x1:.4f},{y1:.4f}", "geometryType": "esriGeometryEnvelope", "inSR": "4326",
            "spatialRel": "esriSpatialRelIntersects", "outFields": "BASENAME,POP100,INTPTLAT,INTPTLON", "outSR": "4326",
            "f": "geojson", "geometryPrecision": "3", "maxAllowableOffset": "0.002"})
        try:
            with urllib.request.urlopen(url, timeout=180) as resp:
                data = json.loads(resp.read())
        except (urllib.error.URLError, ValueError, TimeoutError) as e:
            raise ValueError(f"Couldn't get city names from the US Census ({getattr(e, 'reason', e)}). Try again in a minute.")
        if data.get("error"):
            raise ValueError(f"The Census place service refused the request: {str(data['error'])[:200]}")
        for f in data.get("features", []):
            pr = f.get("properties") or {}
            name = str(pr.get("BASENAME") or "").strip()
            if not name or not f.get("geometry"):
                continue
            try:
                c = [round(float(pr["INTPTLAT"]), 5), round(float(pr["INTPTLON"]), 5)]
            except (KeyError, TypeError, ValueError):
                c = _geom_center(f["geometry"])
            if not c:
                continue
            cities[name] = {"c": c, "pop": int(pr.get("POP100") or 0), "cdp": layer == 26}
            polys.append((name, f["geometry"]))
    where = {}
    for key, (lat, lon) in points.items():
        for name, g in polys:
            if _in_geometry(lat, lon, g):
                where[key] = name
                break
    out = {"cities": cities, "where": where}
    _write_json_file(CITY_FILE, out)
    _cities = out
    return len(cities)


def _geom_center(g):
    """A point inside-ish a GeoJSON polygon: the mean of its outer ring (good enough for a label/dot when the
    Census gave no interior point)."""
    try:
        ring = g["coordinates"][0] if g["type"] == "Polygon" else g["coordinates"][0][0]
        return [round(sum(p[1] for p in ring) / len(ring), 5), round(sum(p[0] for p in ring) / len(ring), 5)]
    except (KeyError, TypeError, IndexError, ZeroDivisionError):
        return None


def zip_shapes_geojson():
    return {"type": "FeatureCollection", "features": [
        {"type": "Feature", "geometry": v["g"], "properties": {"zip": z, "pop": v.get("pop")}} for z, v in zip_shapes().items()]}


def maps_trends():
    """New patients and $ collected per ZIP, month, office and General/Ortho type, for every month loaded: the
    Charts tab's "ZIP trends". ZIP-level counts only, nothing per patient."""
    conn = db()
    rows = [dict(r) for r in conn.execute(
        "SELECT m.zip, substr(m.period_from, 1, 7) AS month, m.office, COALESCE(t.type, 'Unknown') AS type, COUNT(*) AS n, "
        "SUM(CASE WHEN m.collection > 0 THEN 1 ELSE 0 END) AS n_paid, ROUND(SUM(m.collection), 2) AS total "
        "FROM map_patients m LEFT JOIN patient_types t ON t.office = m.office AND t.pat_id = m.pat_id "
        "GROUP BY m.zip, month, m.office, type")]
    months = sorted({r["month"] for r in rows})
    # The latest "Thru" date loaded for each month: a month whose listings stop before its last day is partial.
    thru = {r[0]: r[1] for r in conn.execute("SELECT substr(period_from, 1, 7), MAX(as_of) FROM map_patients GROUP BY 1")}
    pins = [dict(r) for r in conn.execute("SELECT office, lat, lon FROM office_locations ORDER BY office")]
    offices = known_offices(conn)
    short = office_short_names(conn)
    types_loaded = [r[0] for r in conn.execute("SELECT DISTINCT office FROM patient_types")]
    conn.close()
    cn = city_names()
    shapes, cents = zip_shapes(), zip_centroids()
    centers = {z: c for z in {r["zip"] for r in rows} if (c := (shapes.get(z) or {}).get("c") or cents.get(z))}
    return {"months": months, "thru": thru, "rows": rows, "offices": offices, "short": short, "types_loaded": types_loaded,
            "pins": pins, "centers": centers,
            "where": {z: c for z, c in cn.get("where", {}).items() if not z.startswith("office:")}}


def office_short_names(conn):
    """Short office names for map labels, from the names the % TO GOAL sheet uses (office_alias), e.g.
    'Dentist Of Torrance' -> 'Torrance', 'Care Dental Center' -> 'Carson'; else the name without 'Dentist(s) Of'."""
    best = {}
    for r in conn.execute("SELECT alias, office FROM office_alias"):
        if r["office"] not in best or len(r["alias"]) < len(best[r["office"]]):
            best[r["office"]] = r["alias"]
    out = {}
    for o in known_offices(conn):
        short = title_case(best[o]) if o in best else re.sub(r"^Dentists?\s+Of\s+", "", o)
        out[o] = short
    return out


def download_zip_centroids():
    """One time: the US Census ZIP (ZCTA) gazetteer, public data; only each ZIP's center point is kept."""
    global _zips
    import io
    with urllib.request.urlopen(ZIP_SOURCE, timeout=120) as resp:
        raw = resp.read()
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        text = z.read(next(n for n in z.namelist() if n.endswith(".txt"))).decode("utf-8", "replace")
    lines = [l.split("\t") for l in text.splitlines()]
    head = [h.strip().upper() for h in lines[0]]
    iz, ila, ilo = head.index("GEOID"), head.index("INTPTLAT"), head.index("INTPTLONG")
    out = {l[iz].strip(): [round(float(l[ila]), 5), round(float(l[ilo]), 5)] for l in lines[1:] if len(l) > ilo}
    _write_json_file(ZIP_FILE, out)
    _zips = out
    return {"zips": len(out)}


def map_office_name(conn, name):
    """A report's office name as the log's office name (exact, alias, or the same words)."""
    known = {normalize(o): o for o in known_offices(conn)}
    aliases = {r["alias"]: r["office"] for r in conn.execute("SELECT alias, office FROM office_alias")}
    return known.get(normalize(name)) or aliases.get(normalize(name)) or name


def save_patient_types(path, filename):
    parsed = parse_patient_types(path)
    conn = db()
    office = map_office_name(conn, parsed["office"])
    now = datetime.now().isoformat(timespec="seconds")
    with conn:
        conn.execute("DELETE FROM patient_types WHERE office = ?", (office,))
        conn.executemany("INSERT INTO patient_types VALUES (?, ?, ?, ?)",
                         [(office, pid, t, now) for pid, t in parsed["types"].items() if t])
    conn.close()
    vals = list(parsed["types"].values())
    return {"office": office, "general": vals.count("General"), "ortho": vals.count("Ortho"), "patients": len(vals)}


def save_np_appts(path, filename):
    """Keep the new-patient appointments as phone fingerprints per office and date; nothing else from the file."""
    parsed = parse_new_patient_appts(path)
    conn = db()
    now = datetime.now().isoformat(timespec="seconds")
    out, rows, placeholders = [], [], 0
    for o in parsed["offices"]:
        office = map_office_name(conn, o["office"])
        usable = 0
        for pt in o["patients"]:
            fps = {k: phone_fp(v) for k, v in pt["phones"].items()}
            if pt["phones"] and not any(fps.values()):
                placeholders += 1   # a number was there but it's a placeholder or too short
            fps = {k: v for k, v in fps.items() if v}
            if fps:
                usable += 1
            rows += [(fp, office, pt["appt_date"], kind, now) for kind, fp in fps.items()]
        out.append({"office": office, "patients": len(o["patients"]), "with_phone": usable})
    with conn:
        conn.executemany("INSERT OR IGNORE INTO np_appts VALUES (?, ?, ?, ?, ?)", rows)
        total = conn.execute("SELECT COUNT(DISTINCT fp) FROM np_appts").fetchone()[0]
    conn.close()
    return {"offices": out, "patients": sum(o["patients"] for o in out), "with_phone": sum(o["with_phone"] for o in out), "placeholders": placeholders,
            "total_numbers_known": total}


def save_patient_phones(path):
    """Keep each new patient's phone fingerprints with their Pat ID, office and first-visit date; nothing else."""
    parsed = parse_new_patients_list(path)
    conn = db()
    office = map_office_name(conn, parsed["office"])
    now = datetime.now().isoformat(timespec="seconds")
    rows, usable, placeholders = [], 0, 0
    for pt in parsed["patients"]:
        fps = {k: phone_fp(v) for k, v in pt["phones"].items()}
        if pt["phones"] and not any(fps.values()):
            placeholders += 1
        fps = {k: v for k, v in fps.items() if v}
        if fps:
            usable += 1
        rows += [(fp, pt["pat_id"], office, pt["first_visit"], pt["created_on"], kind, now) for kind, fp in fps.items()]
    with conn:
        conn.executemany("INSERT OR REPLACE INTO patient_phones VALUES (?, ?, ?, ?, ?, ?, ?)", rows)
        total = conn.execute("SELECT COUNT(DISTINCT fp) FROM patient_phones").fetchone()[0]
    conn.close()
    visits = sorted(p["first_visit"] for p in parsed["patients"] if p["first_visit"])
    return {"kind": "patients", "office": office, "patients": len(parsed["patients"]), "with_phone": usable, "placeholders": placeholders,
            "first_visits": [visits[0], visits[-1]] if visits else None, "total_numbers_known": total}


def save_calls_upload(path, filename):
    """The Calls tab accepts two Denticon reports: the New Patients List (preferred: Pat ID + phone + first visit)
    and New Patients Through Appointment Book (phone + appointment date)."""
    kind = calls_upload_kind(path)
    if kind == "patients":
        return save_patient_phones(path)
    if kind == "appointments":
        return {"kind": "appointments", **save_np_appts(path, filename)}
    raise ParseError("Upload Denticon's New Patients List (Excel) or New Patients Through Appointment Book (Excel)")


def set_office_location(c):
    office = str(c.get("office") or "").strip()
    if not office:
        raise ValueError("Pick the office")
    address = str(c.get("address") or "").strip()
    lat, lon = c.get("lat"), c.get("lon")
    if address and (lat in (None, "") or lon in (None, "")):
        # Look the office's business address up with the US Census geocoder (no patient data is sent).
        url = GEOCODER + "?" + urllib.parse.urlencode({"address": address, "benchmark": "Public_AR_Current", "format": "json"})
        try:
            with urllib.request.urlopen(url, timeout=60) as resp:
                matches = json.loads(resp.read())["result"]["addressMatches"]
        except (urllib.error.URLError, ValueError, KeyError) as e:
            raise ValueError(f"Couldn't look up that address ({e}). Enter latitude and longitude instead.")
        if not matches:
            raise ValueError("The Census geocoder didn't find that address. Check it, or enter latitude and longitude.")
        lat, lon = matches[0]["coordinates"]["y"], matches[0]["coordinates"]["x"]
    try:
        lat, lon = float(lat), float(lon)
    except (TypeError, ValueError):
        raise ValueError("Enter the office's address, or its latitude and longitude")
    if not (-90 <= lat <= 90 and -180 <= lon <= 180):
        raise ValueError("Those don't look like latitude and longitude")
    conn = db()
    with conn:
        conn.execute("INSERT OR REPLACE INTO office_locations VALUES (?, ?, ?, ?)", (office, address, lat, lon))
    conn.close()
    return {"office": office, "lat": lat, "lon": lon}


def maps_info(month=None):
    """New patients by ZIP for a month (YYYY-MM), with each ZIP's center, the offices' pins and General/Ortho."""
    conn = db()
    months = [r[0][:7] for r in conn.execute("SELECT DISTINCT period_from FROM map_patients ORDER BY period_from DESC")]
    month = month if month in months else (months[0] if months else None)
    rows = conn.execute("SELECT m.*, t.type FROM map_patients m LEFT JOIN patient_types t ON t.office = m.office AND "
                        "t.pat_id = m.pat_id WHERE m.period_from = ?", (f"{month}-01",)).fetchall() if month else []
    pins = [dict(r) for r in conn.execute("SELECT office, lat, lon FROM office_locations ORDER BY office")]
    offices = known_offices(conn)
    short = office_short_names(conn)
    types_loaded = [r[0] for r in conn.execute("SELECT DISTINCT office FROM patient_types")]
    # Each office's listing has its own Thru date; show the range when they differ.
    thru = dict(zip(("min", "max"), conn.execute("SELECT MIN(as_of), MAX(as_of) FROM map_patients WHERE period_from = ?",
                                                   (f"{month}-01",)).fetchone())) if month else None
    region = map_region(conn)
    conn.close()
    zips, missing, pops, shapes = zip_centroids(), {}, zip_population(), zip_shapes()
    no_zcta = set(_load_json_file(NO_ZCTA_FILE) or []) if shapes else set()
    by_zip = {}
    for r in rows:
        z = by_zip.setdefault(r["zip"], {"zip": r["zip"], "patients": []})
        z["patients"].append({"pat_id": r["pat_id"], "office": r["office"], "collection": r["collection"],
                              "type": r["type"] or "Unknown", "referral": r["referral_type"]})
    out = []
    # Three cases for a ZIP without a boundary: the Census has no area for it ("missing": not on the map), or it just
    # hasn't been downloaded yet ("no_shape": step 1 again adds it; drawn as a dot meanwhile if its center is known).
    no_shape = []
    for z in by_zip.values():
        shape = shapes.get(z["zip"])
        center = (shape or {}).get("c") or zips.get(z["zip"]) or (_geom_center(shape["g"]) if shape else None)
        if not shape and z["zip"] in no_zcta:
            missing[z["zip"]] = len(z["patients"])
            continue
        if not shape:
            no_shape.append(z["zip"])
        if center:
            z["lat"], z["lon"] = center
            z["pop"] = (shape or {}).get("pop") or pops.get(z["zip"])
            z["shape"] = bool(shape)
            out.append(z)
        elif not shapes:   # step 1 never run and no center known either
            missing[z["zip"]] = len(z["patients"])
    cn = city_names()
    return {"month": month, "months": months, "thru": thru, "region": region, "zips": out, "missing": missing, "no_shape": sorted(no_shape),
            "cities": cn.get("cities", {}), "where": cn.get("where", {}),
            "pins": pins, "offices": offices, "short": short, "zip_data": bool(zips or shapes), "shapes_data": len(shapes),
            "pop_data": bool(shapes or pops), "types_loaded": types_loaded}


def audit_info():
    """The admin Audit tab: Denticon's lender payments vs the NP sheets, per office and month."""
    issues = denticon_checks()
    conn = db()
    journals = [dict(r) for r in conn.execute("SELECT office, period_from, as_of, source FROM journal_uploads "
                                              "ORDER BY period_from DESC, office")]
    checked = {f"{r['office']}|{r['period_from']}": r["n"] for r in conn.execute(
        "SELECT office, period_from, COUNT(*) AS n FROM np_patients GROUP BY office, period_from")}
    conn.close()
    return {"journals": journals, "issues": issues, "checked": checked}


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


PATIENT_KEY_FILE = DATA_DIR / "patient_key"   # secret for hashing Pat IDs (git-ignored with data/)
_patient_secret = None


def pat_key(pat_id):
    """A keyed hash of a Pat ID: the same patient always gets the same key, but the ID can't be read back."""
    global _patient_secret
    pid = str(pat_id or "").strip()
    try:
        pid = str(int(float(pid)))          # "90217483.0" from a Google Sheet = "90217483" in Denticon
    except ValueError:
        pass
    if not pid:
        return ""
    if _patient_secret is None:
        if not PATIENT_KEY_FILE.exists():
            DATA_DIR.mkdir(exist_ok=True)
            fd = os.open(PATIENT_KEY_FILE, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
            with os.fdopen(fd, "wb") as f:
                f.write(secrets.token_bytes(32))
        _patient_secret = PATIENT_KEY_FILE.read_bytes()
    return hmac.new(_patient_secret, pid.encode(), hashlib.sha256).hexdigest()[:24]


def phone_fp(number):
    """A keyed fingerprint of a phone number (its 10 digits): the same number always gives the same fingerprint,
    the number can't be read back. Empty when fewer than 10 digits."""
    digits = re.sub(r"\D", "", re.split(r"\s*(?:x|ext)\b", str(number or ""), maxsplit=1, flags=re.I)[0])   # "… x: 12" = extension
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]
    if len(digits) != 10 or len(set(digits)) == 1 or digits.startswith("000") or digits in ("1234567890", "0123456789"):
        return ""   # too short, or a placeholder the front desk typed (000-000-0000, 999-999-9999, 123-456-7890)
    pat_key("0")   # loads the secret
    return hmac.new(_patient_secret, ("phone:" + digits).encode(), hashlib.sha256).hexdigest()[:24]


def scrub_patients(entries):
    """Replace Pat IDs with keys in a list of patient entries (in place); returns the list."""
    for p in entries or []:
        if "pat_id" in p:
            p["pk"] = pat_key(p.pop("pat_id"))
    return entries


def scrub_result(result):
    """No Pat IDs leave the server: hash them in a parsed report before it goes to the browser."""
    for o in result.get("offices", []):
        scrub_patients(o.get("_patients"))
        scrub_patients((o.get("details") or {}).get("_patients"))
    return result


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
            raise ValueError("File is too large (200 MB max)")
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
            if (self.command != "GET" or self.path.split("?", 1)[0] not in VIEWER_PATHS) and not (self.command == "POST" and self.path in ("/api/chat", "/api/chat/file")):
                self._send(403, {"error": "This is a view-only link."})
                return True
        elif APP_PASSWORD and not app_password_ok(auth, given):
            time.sleep(1)
            self._send(401, {"error": "Password required"},
                       extra={"WWW-Authenticate": 'Basic realm="Monthly report", charset="UTF-8"'})
            return True
        if self.path.split("?", 1)[0] in LOCAL_ONLY_PATHS and not self._on_this_computer():
            self._send(403, {"error": "For safety, this only works on the computer that runs the app."})
            return True
        return False

    def _on_this_computer(self):
        """True only for a browser on the computer running the app. Requests passed on by a proxy (Tailscale
        Serve, which connects from this computer) carry forwarding headers and count as other computers."""
        if self.client_address[0] not in ("127.0.0.1", "::1", "::ffff:127.0.0.1"):
            return False
        forwarded = ("X-Forwarded-For", "X-Forwarded-Host", "Forwarded", "Tailscale-User-Login", "Tailscale-User-Name")
        return not any(self.headers.get(h) for h in forwarded)

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
        elif self.path.startswith("/api/log/days"):
            q = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
            try:
                self._send(200, log_days((q.get("from") or [""])[0], (q.get("to") or [""])[0]))
            except ValueError as e:
                self._send(400, {"error": str(e)})
        elif self.path == "/api/inbox":
            self._send(200, inbox_info())
        elif self.path == "/api/assistant/status":
            self._send(200, assistant_status())
        elif self.path == "/api/days":
            self._send(200, days_info())
        elif self.path == "/api/robot":
            self._send(200, robot_status())
        elif self.path == "/api/log":
            conn = db()
            rows = conn.execute(
                f"SELECT {', '.join(LOG_COLUMNS + [m + '_detail' for m in DETAIL_METRICS])} "
                "FROM monthly_log ORDER BY period_thru DESC, office"
            ).fetchall()
            conn.close()
            self._send(200, month_rows_from_days([dict(r) for r in rows]))
        elif self.path == "/api/benchmarks":
            conn = db()
            rows = [dict(r) for r in conn.execute("SELECT * FROM benchmarks ORDER BY metric")]
            conn.close()
            self._send(200, rows)
        elif self.path == "/api/google/notes-sheet":
            self._send(200, notes_sheet_info())
        elif self.path == "/api/maps/shapes":
            self._send(200, zip_shapes_geojson())
        elif self.path == "/api/maps/trends":
            self._send(200, maps_trends())
        elif self.path.startswith("/api/maps"):
            q = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
            self._send(200, maps_info((q.get("month") or [None])[0]))
        elif self.path == "/api/denticon/status":
            self._send(200, denticon_status())
        elif self.path == "/api/invoca/status":
            self._send(200, invoca_status())
        elif self.path == "/api/calls/lines":
            self._send(200, promo_lines_info())
        elif self.path.startswith("/api/calls/funnel"):
            q = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
            self._send(200, calls_funnel((q.get("month") or [None])[0], (q.get("office") or [""])[0]))
        elif self.path.startswith("/api/calls"):
            q = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
            self._send(200, calls_info((q.get("date") or [None])[0]))
        elif self.path == "/api/audit":
            self._send(200, audit_info())
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
                self._send(200, scrub_result(result))
            elif self.path == "/api/save":
                p = json.loads(self._body())
                self._send(200, save_daily(p) if is_day_report(p, "executive_summary") else save_rows(p))
            elif self.path == "/api/parse-paste":
                self._send(200, read_paste(json.loads(self._body()).get("text", "")))
            elif self.path == "/api/maps/patient-types":
                filename = unquote(self.headers.get("X-Filename", "upload.xlsx"))
                with tempfile.NamedTemporaryFile(suffix=Path(filename).suffix.lower(), delete=False) as tmp:
                    tmp.write(self._body())
                try:
                    self._send(200, save_patient_types(tmp.name, filename))
                finally:
                    os.unlink(tmp.name)
            elif self.path == "/api/maps/office-location":
                self._send(200, set_office_location(json.loads(self._body())))
            elif self.path == "/api/maps/zip-download":
                self._send(200, download_zip_centroids())
            elif self.path == "/api/maps/shapes-download":
                self._send(200, download_zip_shapes())
            elif self.path == "/api/google/notes-sheet":
                self._send(200, connect_notes_sheet(json.loads(self._body()).get("url", "")))
            elif self.path == "/api/google/notes-sheet/delete":
                self._send(200, disconnect_notes_sheet())
            elif self.path == "/api/google/notes-sheet/write":
                self._send(200, {"status": write_notes_sheet(force=True) or "nothing to write (not connected)", **notes_sheet_info()})
            elif self.path == "/api/denticon/key":
                denticon.save_key(json.loads(self._body()).get("key"))
                self._send(200, {"ok": True})
            elif self.path == "/api/denticon/key/delete":
                denticon.remove_key()
                self._send(200, {"ok": True})
            elif self.path == "/api/denticon/test":
                self._send(200, denticon_test())
            elif self.path == "/api/invoca/save":
                b = json.loads(self._body())
                invoca.save(b.get("site"), b.get("kind"), b.get("id"), b.get("token"))
                invoca_state.update(last_test=None, last_error=None, last_result=None)
                self._send(200, invoca_status())
            elif self.path == "/api/invoca/delete":
                invoca.remove()
                invoca_state.update(last_test=None, last_error=None, last_result=None)
                self._send(200, invoca_status())
            elif self.path == "/api/invoca/test":
                self._send(200, invoca_test())
            elif self.path == "/api/calls/pull":
                self._send(200, pull_invoca_calls(json.loads(self._body()).get("date")))
            elif self.path == "/api/calls/pull-month":
                self._send(200, pull_invoca_month(json.loads(self._body()).get("month")))
            elif self.path == "/api/calls/lines":
                self._send(200, set_promo_line(json.loads(self._body())))
            elif self.path == "/api/calls/upload":
                filename = unquote(self.headers.get("X-Filename", "upload.xlsx"))
                with tempfile.NamedTemporaryFile(suffix=Path(filename).suffix.lower(), delete=False) as tmp:
                    tmp.write(self._body())
                try:
                    self._send(200, save_calls_upload(tmp.name, filename))
                finally:
                    os.unlink(tmp.name)
            elif self.path == "/api/denticon/offices":
                self._send(200, set_denticon_office(json.loads(self._body())))
            elif self.path == "/api/chat":
                b = json.loads(self._body())
                who = "viewer link" if VIEWER_PASSWORD else ("this computer" if self._on_this_computer() else "other computer")
                self._send(200, ai_chat(b.get("question"), b.get("history"), who, file_ids=b.get("files") or []))
            elif self.path == "/api/chat/file":
                raw = self._body()
                try:
                    self._send(200, chat_files.add(unquote(self.headers.get("X-Filename", "file")), raw))
                except assistant.AssistantError as e:
                    raise ValueError(str(e))
            elif self.path == "/api/assistant/key":
                try:
                    claude_cfg.save(json.loads(self._body()).get("key"))
                except assistant.AssistantError as e:
                    raise ValueError(str(e))
                self._send(200, assistant_status())
            elif self.path == "/api/assistant/key/delete":
                claude_cfg.remove()
                self._send(200, assistant_status())
            elif self.path == "/api/assistant/test":
                self._send(200, ai_chat("In one sentence: what is the latest month in the Log, and its total collection?", [], "this computer (test)", use_cache=False))
            elif self.path == "/api/assistant/cache/clear":
                self._send(200, assistant_cache_clear())
            elif self.path == "/api/robot/pull":
                b = json.loads(self._body())
                self._send(200, robot_pull(b.get("from"), b.get("to"), b.get("report") or "exec"))
            elif self.path == "/api/save-lender-journal":
                self._send(200, save_lender_journal(json.loads(self._body())))
            elif self.path == "/api/save-metrics":
                p = json.loads(self._body())
                daily = p.get("from_report") and is_day_report(p, p.get("report_type"))
                self._send(200, save_daily_metrics(p) if daily else save_metrics(p))
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
        except (ParseError, PasteError, ValueError, KeyError, sqlite3.Error, denticon_api.DenticonError, invoca_api.InvocaError,
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
        threading.Thread(target=backup_loop, daemon=True).start()
        threading.Thread(target=notes_sheet_loop, daemon=True).start()
        threading.Thread(target=inbox_loop, daemon=True).start()
    print(f"Executive Summary logger running at http://{HOST}:{PORT}  (Ctrl+C to stop)")
    if HOST == "0.0.0.0":
        print(f"On the office network: http://{socket.gethostname()}.local:{PORT}  (password required)")
    print(f"Log: {DB_PATH}  |  CSV: {CSV_PATH}")
    try:
        ThreadingHTTPServer((HOST, PORT), Handler).serve_forever()
    except KeyboardInterrupt:
        pass
