# Monthly-report

Executive Summary logger for Care Dentistry Group. Upload a Denticon / PlanetDDS
**Executive Summary** (`.pdf` or `.xlsx`), and it pulls these numbers for every office in the report:

| Field | Where it comes from |
|---|---|
| Total collection | Collection (logged as a positive amount) |
| Insurance collection | Collection → Insurance (logged as a positive amount) |
| New patients | New Patients by First Visit Date − Referrals → EXISTING PATIENT |
| Total production | Production → Total |

It also keeps the two new-patient inputs (`np_first_visit`, `existing_patient_referrals`) in the log so you can check the math.

## Run it

```bash
python3 app.py
```

Then open http://127.0.0.1:8765. You only need Python 3 and `pdftotext` (`sudo apt-get install poppler-utils`).

- **Upload**: drop in a report. Each office is marked **New**, **Changed** (old values shown) or **Unchanged** compared to the log. Set the period (PDFs fill it in automatically; XLSX exports don't include the dates), then **Save to log**.
- **Log**: browse rows, open the original report, view each row's history, void rows, download the CSV, and lock or unlock months.
- **SQL**: run read-only SQL. There are example queries, including an audit trail, and you can download the results as CSV.

## Audit trail

- **Original reports are kept.** Every report that parses is saved in `data/files/`, named by its SHA-256 fingerprint so it's clear it hasn't been changed. Each row links to the report it came from.
- **Nothing is deleted.** Rows are *voided* with a required reason. Voided rows drop out of totals and the CSV but stay in the log (tick **Show voided**).
- **Changes keep their history.** Re-uploading a period with different numbers updates the row and records the old and new values in `summary_history`.
- **Closed months can be locked.** A locked month rejects saves and voids until someone unlocks it with a reason. Locks and unlocks are both recorded.
- Not yet: *who* did each action. That comes with Auth0 in the main platform.

## Where the data lives

| Path | What |
|---|---|
| `data/reports.db` | SQLite database, the source of truth |
| `data/files/` | The original uploaded reports |
| `data/office_summary.csv` | Live (non-voided) rows, rewritten after every change, for Excel |

Tables: `office_summary` (all rows, including voided), `active_summary` (a view of the live rows; use this for reporting), `summary_history`, `source_files`, `period_locks`. The SQL is kept portable so the tables can move to PostgreSQL.

`data/` and report files are git-ignored, so **financial data is never committed**. This repo is public. Nothing is backed up online; back up the `data/` folder yourself.

## Command line

```bash
python3 summary_parser.py "RV Anaheim.pdf"
```

This prints the parsed numbers as JSON without saving anything.
