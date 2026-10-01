# Monthly-report — handoff notes

Context for picking this project up in a new session. Written 2026-09-30.
**This repo is public**: this file has no patient data, credentials or office financial figures, and none should be added.

---

## 1. What it is

A local web app for **Care Dentistry Group's finance monthly audit**. It pulls specific numbers per office per month out of
Denticon/PlanetDDS reports, pasted spreadsheet rows, and each office's "New Patient Analysis" workbook / Google Sheet, and
logs them in one table (the **Log**) with history, a change list, SQL access and CSV/copy export.

- User: `data679` (GitHub), works in finance/data at Care Dentistry Group. Non-developer; wants plain explanations,
  step-by-step instructions, and to verify numbers (hence hover lists showing which sheet rows were counted).
- Long-term: may be integrated into another platform (Node.js, PostgreSQL, BullMQ + staging tables, React/TS +
  Tremor/Recharts, Auth0, Render/Railway/AWS, HIPAA). Keep SQL portable to PostgreSQL.

## 2. Running it

- Folder: `~/Monthly-report` (git repo, remote `https://github.com/data679/Monthly-report.git`, branch `main`).
- Start: `./run.sh` in the "Report logger" terminal tab → http://127.0.0.1:8765 (localhost only).
- `run.sh` loops `python3 app.py`, so **to load code changes, kill the `python3 app.py` process** (its cwd is
  `~/Monthly-report`) and `run.sh` restarts it in the same tab. Page-only changes (static/index.html) need no restart —
  just a hard refresh (Ctrl+Shift+R); index.html is served with `Cache-Control: no-store`.
- The terminal panel refuses a 7th tab opened by Claude; avoid opening new tabs — restart via kill + run.sh.
- Python 3.13, **standard library only** (no pip on this machine). Needs `pdftotext` (poppler) for PDFs and `openssl`
  for Google sign-in.
- After a reboot the app isn't running; start `./run.sh` again.

## 3. Git status

- **Nothing has been committed since the first "Add README" commit.** All code is untracked/modified. User has been
  offered commit+push several times; not done yet — ask before committing.
- `.gitignore` excludes `data/`, `*.pdf`, `*.xlsx`, `*.xlsm`, `*.csv`, `*.db`, `__pycache__/`. The log, backups and the
  Google key live in `data/` and must never be committed.
- `README.md` is **stale**: at some point it was rewritten (not by Claude) to describe features that were never built
  (voiding, locked months, stored source files, `summary_history`, `period_locks`). It needs rewriting to match reality.
- `.claude/launch.json` exists (preview config, port 8765).

## 4. Files

| File | Purpose |
|---|---|
| `app.py` | HTTP server (stdlib `ThreadingHTTPServer`), SQLite schema/migrations, save logic, change log, Google sync loop, endpoints |
| `summary_parser.py` | Detects report type and parses Denticon reports (PDF via pdftotext, XLSX via zipfile/XML). Routes NP workbooks to `np_analysis` |
| `np_analysis.py` | Reads an office's **New Patient Analysis** workbook (downloaded `.xlsm/.xlsx` or Google Sheet) and computes all per-office counts + row-level detail lists |
| `google_sheets.py` | Service-account sign-in (JWT signed with `openssl`), Sheets API v4 read-only; `SheetsWorkbook` mimics `np_analysis.Workbook` |
| `sheet_paste.py` | Parses rows pasted from a spreadsheet (tab-separated); currently only the `refunds` column; office-name matching/aliases |
| `static/index.html` | Whole UI (vanilla JS/CSS, light+dark) |
| `run.sh` | Auto-restart loop |
| `data/reports.db` | SQLite log (git-ignored). Backups: `data/reports.backup-*.db` |
| `data/office_summary.csv` | CSV mirror of the `monthly_log` view, rewritten on every change |
| `data/google/service_account.json` | Google service-account key (chmod 600, git-ignored) |

## 5. Data model (SQLite, `data/reports.db`)

- `office_summary` — Executive Summary numbers. **One row per office per month**: `UNIQUE(office, period_from)`.
- `office_history` — old versions of replaced/deleted `office_summary` rows.
- `office_metrics` — every other number: `(office, period_from, as_of, metric, value, source, logged_at, detail)`,
  `UNIQUE(office, period_from, metric)`. `detail` = JSON list for hover (sheet rows etc., never names/IDs).
- `metric_history` — old versions of `office_metrics`, with `action` (replaced/deleted) and **`replaced_by`** (source
  that replaced it — used to undo a sheet). Match history by office+month+metric, **not** `metric_id` (ids get reused).
- `change_log` — every change to a metric value: when, office, month, metric, old → new, `via` (upload / paste /
  "Google Sheets sync" / "sheet removed" / "undo test sheet"), source. Shown in the Changes tab.
- `office_alias` — pasted-sheet office names → log office names (e.g. `carson` → Care Dental Center).
- `sheet_sources` — connected Google Sheets (spreadsheet_id, url, title, office, last sync/status).
- View **`monthly_log`** — one row per office+month: office_summary FULL OUTER JOIN pivoted office_metrics, plus
  calculated columns (below). The Log, CSV and SQL tab read this.

### Save rules
- Reports are **month-to-date**: `period_from` must be the 1st; `period_thru` in the same month.
- Per office+month+metric the **newest `as_of` wins**; an older upload/sync is **skipped**; same `as_of` → replaced if
  the value differs (old value → history + change_log), "unchanged" otherwise.
- Replaced/deleted values always go to history first.

## 6. Report types & how each number is extracted

Upload box auto-detects the type. Uploaded files are parsed from a temp file and deleted; only numbers are stored.

**Executive Summary** (PDF or XLSX; group export has one sheet per office):
- `total_collection` = Collection line total (abs). `insurance_collection` = Collection → Insurance (abs).
- `new_patients` = New Patients by First Visit Date − Referrals → EXISTING PATIENT count
  (`np_first_visit`, `existing_patient_referrals` also stored).
- `total_production` = Production → **Total** (includes "Non visit code transactions"; user was told, may want the
  plain Production line instead — not decided).
- Offices with blank data are skipped. XLSX has no dates → user picks Thru; From auto = 1st.

**Refunds** — pasted from the manager's sheet (Upload tab → "Paste from a spreadsheet"); header row with `Office` and
`refunds`; "Updated M/D/YY" → as-of date. Office matching via aliases (all 15 known now).

**Referral Production Listing** (PDF/XLSX) → `np_collection` = "Total for Office, X [n]" Collection (abs). Must equal
Grand Total or the file is rejected. Handles wrapped lines and numbers one space apart in PDFs.

**Daily Journal – Detail** (PDF, must be Patient Type: Ortho) → `ortho_collection` = "Summary for X" → Total Payments
(last column). Must equal Grand Total Summary.

**Treatment Plan Status – Detail** (PDF/XLSX, proc code ZREFORTH) → `ortho_referrals` = the `[n]` on
"Total for office, X [n]". Cross-checked against provider sub-totals. PDF must show ZREFORTH filter (XLSX can't; UI warns).

**New Patient Analysis workbook** (per office, `.xlsm/.xlsx` or Google Sheet) — see §7.

## 7. New Patient Analysis rules (np_analysis.py)

- One tab per month; reads the **latest month tab that has patients** (empty newer tabs, e.g. next month set up early,
  are skipped). Sync reads latest + previous month with patients. Tab names like "September 2026📚", "SEPT 2026",
  "Sept 2026", "APR26", "August 1 2026", "November 25" (=Nov 2025) are recognised; "Master…", "Original…",
  "do not use…" are ignored.
- Header detection: a row with **"Insurance Type"** is the team header. **"Pat ID"** may be on that row or only on
  Denticon's repeated header rows; each pasted Denticon block's own header row also sets **Pat ID and Collection
  columns** (handles pastes shifted sideways, e.g. Bixby Knolls).
- **Patient row** = numeric Pat ID. **Count every row** on the tab, including rows below the sheet's Grand Total and
  duplicate patients (user's choice).
- Office from "Office: X" / "Total for Office, X" line, **trailing dates stripped** (Montclair writes
  "Office: Dentist Of Montclair 09/02/26-09/04/26"); if a tab has no Office line, use the office from the workbook's
  other month tab.
- Thru date = latest Excel date marker in column A within the month (else user picks; sync falls back to today / month end).
- Cells showing `#DIV/0!` etc. are ignored.

Metrics (all have hover detail lists: sheet row + reason; never names/Pat IDs):
- `medi_cal_patients`, `hmo_patients` — Insurance Type contains Medi-Cal / HMO; **multi-select cells** ("Cash, Medi-Cal")
  count if any part matches. Medi-Cal only (not Medicare). "DENTICAL" is **not** counted yet (open question).
- `no_ssn_itin_patients` — any text cell on the row (except column A) matches no SSN/social/SS#/ITIN phrasing, incl.
  "no ssn only itin" and "itin only"; "no id" doesn't count; or a dedicated "No SSN/ITIN" column = Yes/Y/X.
- `existing_finance_patients` — row text says existing/pre-existing/already-has a lender account (Care Credit, Alphaeon,
  Cherry, Sunbit, …). Family accounts ("parents had existing care credit") count.
- `financing_applicants` — any prime lender "Ran" = Yes, or a subprime lender named under "Subprime Co. Ran".
- Prime lenders = columns with "… Yes / No" (with or without "Ran"), each followed by "$Amt Approved", "$Amt Coll".
  - `prime_applicants` — any prime Ran = Yes. `prime_apps_ran` — count of Yes cells (4 lenders = 4).
  - `prime_approved` — applied AND an approved or collected amount. `prime_apps_approved` — per lender, only lenders ran.
  - Existing-account collection without a new app (e.g. parents' Care Credit) is **not** an applicant/approval.
  - `prime_approved_amount`, `prime_collected_amount` — sums of $ approved / $ collected (collected includes existing-account money).
  - `prime_funded` — any prime collected amount (incl. existing accounts). `prime_approved_funded` — approved AND funded.
- Subprime (one set: Subprime Co. Ran / Co. Approved / $Amt Approved / $Amt Coll):
  - `subprime_applicants` — lender named under Ran. `subprime_apps_ran` — number of lenders named.
  - `subprime_approved` — applied AND (lender named under Approved OR an amount). `subprime_apps_approved` — lenders named under Approved.
  - `subprime_approved_amount`, `subprime_collected_amount`, `subprime_funded` (collected > 0).
- Combined, **each patient once**: `fin_approved`, `fin_funded`, `fin_approved_funded` (detail shows Prime / Subprime / both).
- Treatment: `np_rows` (patient rows = "per NP" denominator), `tx_diagnosed_amount` ("$ Treatment Diagnosed"),
  `np_collection_wb` (Denticon "Collection" for these patients, as positive), `no_next_visit` ("Has an Appointment" =
  **No** only), `started_tx` (paid anything: Denticon collection, financing collected, or other form of payment).

## 8. Calculated columns (in the `monthly_log` view; NULL when an input is missing)

| Column | Formula |
|---|---|
| net_collections | total_collection − refunds |
| refund_rate | refunds ÷ total_collection |
| existing_pt_collection | total_collection − np_collection |
| irrelevant_rate | (medi_cal + hmo + no_ssn_itin) ÷ new_patients (Exec Summary) |
| ortho_rate | ortho_referrals ÷ new_patients |
| financing_rate | financing_applicants ÷ new_patients |
| prime_approval_rate | prime_approved ÷ prime_applicants |
| prime_collected_rate | prime_collected_amount ÷ prime_approved_amount |
| prime_funded_rate | prime_approved_funded ÷ prime_approved |
| avg_prime_per_approved_np / _approved_app / _app_ran / _applicant | prime_approved_amount ÷ prime_approved / prime_apps_approved / prime_apps_ran / prime_applicants |
| subprime_approval_rate | subprime_approved ÷ subprime_applicants |
| subprime_collected_rate | subprime_collected_amount ÷ subprime_approved_amount |
| fin_approval_rate | fin_approved ÷ financing_applicants |
| fin_apps_ran / fin_apps_approved | prime + subprime |
| fin_app_approval_rate | fin_apps_approved ÷ fin_apps_ran |
| fin_approved_amount / fin_collected_amount / fin_collected_rate | prime + subprime $, and collected ÷ approved |
| fin_funded_rate | fin_approved_funded ÷ fin_approved |
| avg_tx_dx_per_np / collection_per_np | tx_diagnosed_amount / np_collection_wb ÷ np_rows |
| collection_per_dx | np_collection_wb ÷ tx_diagnosed_amount |
| np_loss_rate / start_tx_rate | no_next_visit / started_tx ÷ np_rows |

Totals row: sums; rates/averages = **sum of numerators ÷ sum of denominators** (not an average of rates); columns with no
data show "—" (and copy blank).

## 9. UI (static/index.html)

Tabs: **Upload · Log · Google Sheets · Changes · SQL**.
- **Upload**: drop a file → review table (numbers with hover lists) → Period (From = 1st, Thru editable) → Save to log.
  Paste box for the refunds sheet with office matching dropdowns (remembered as aliases).
- **Log**: one month at a time — **◀ Month ▶** navigator + month grid (months with data enabled) + "All months";
  opens on latest month. Office search; **Sections** switches (Executive Summary, Collections, Ortho, New patients,
  Prime, Subprime, Prime + Subprime, Treatment) with colored bands; **Columns** menu (show/hide, drag/↑↓ reorder,
  Reset); **Office order** menu (drag/↑↓, "Use this order", A–Z); sortable headers; sticky Office column; compact
  full-width table; **Total row**; **Copy** (visible rows/columns as TSV+HTML for Excel/Sheets); Download CSV.
  Row checkboxes / Delete selected are **hidden** (`body.hide-select`), deletions archive to history.
  "as of MM/DD" shown when a value's date differs from the row's period end.
- **Hover lists**: numbers with dotted underline → custom popup (hover / click to pin / Esc), table of sheet rows,
  scrolls inside without bouncing; Copy link when pinned.
- **Google Sheets tab**: key status + service-account email (Copy), setup steps, add sheet link, list with last sync /
  status, **Remove** (stops syncing AND undoes that sheet's numbers from history), Sync now; auto-sync every 15 min.
- **Changes tab**: all value changes, filter, Copy.
- **SQL tab**: read-only queries (`PRAGMA query_only`, read-only connection), examples, CSV of results.
- Layout/sort/sections/office order/hidden columns are saved per browser (localStorage).
- **Benchmarks** (Log toolbar): per-column green/yellow/red cut-offs, higher- or lower-is-better, same for every
  office; table `benchmarks` (seeded with % to ortho 25%/20%). Colors + ●▲▼ marks on office rows only (not the
  Total row, not copies); ◎ in the header with the rule. Percent cut-offs are stored as fractions.
- **Data checks** (Log toolbar "Checks" with a count badge): values that look wrong are **flagged, never changed**.
  - Rules: a rate that can't exceed 100% is over 100% (`RATE_MAX` in index.html; the message names the 3 largest
    amounts behind it); a patient with more collected than approved, per prime lender with an approved amount, or
    subprime (`collected_over_approved` count + detail from np_analysis, "Data checks" section). Collected with no
    approved amount filled in is *not* flagged (common, usually fine).
  - ⚠ badge on the cell (office rows only); the upload review warns before saving.
  - **Mark OK** (note required) → `check_ok`, hidden while the value is the same (`value_sig`); Undo in "Marked OK".
  - **Email office** opens a mailto draft in the user's email app (they press Send; nothing is sent by the app)
    listing every open issue + sheet link; emails saved in `office_contacts`; `check_notified` shows
    "email drafted <date>". Long messages are copied to the clipboard instead. **Copy message** for Teams/Chat.
  - All of these write `change_log` rows (metric `check: <key>` / `office emails`).
  - Uploaded offices need a re-upload to get the per-patient check (West Covina, See Me Smile Oxnard had cases).
- **View-only sharing**: `./share.sh [minutes]` asks for a viewer password, runs a second copy with
  `VIEWER_PASSWORD` on port 8766 (Log + Changes only, every change refused, no sync, no emails/sheet links) and a
  Cloudflare quick tunnel (`~/.local/bin/cloudflared`); stops after 60 min by default or Ctrl+C.

## 10. Google Sheets sync

- User created a Google Cloud project + service account + JSON key and added the key in the app (stored in
  `data/google/`). Offices share their sheet with the service-account email as **Viewer**.
- Connected now: El Segundo and Downey sheets. Others not yet.
- Sync reads latest + previous month tabs with patients, via the same `np_analysis` code as uploads; changes go to
  `change_log` ("Google Sheets sync"). A sheet that's an Excel file in Drive gives a clear error (Save as Google Sheets).
- **Don't upload a workbook for an office whose sheet is connected** (explained to user; an upload with a later Thru
  date would make syncs skip). An upload-warning safeguard was offered, not built.
- Testing tip given to user: test copies should use "Office: TEST Office" or they write into the real office's numbers.
  (A "test" sheet did this once; it was undone from history.)

## 11. Known office/data quirks

- Office names differ between sources; aliases map sheet names (e.g. GARDENA → Gardena Dental Care,
  GARDENA 2 → Dentist Of Gardena, SANTA BARBARA → See Me Smile, CARSON → Care Dental Center, Oxnard → See Me Smile Dental Of Oxnard).
- Northridge appears in the Exec Summary export with no data (skipped); not in other reports.
- **Montclair**: date ranges after the office name on each paste (handled). **Cerritos**: empty next-month tab (handled).
- **Carson**: August tab named "August 1 2026" (handled); August tab has no Pat ID/chart numbers → can't be read
  (only matters for sync's previous month). Carson and Gardena(Dental Care) have their own No SSN/ITIN column.
- **Gardena Dental Care** workbook: team header row labels the Pat ID column "Referral" (handled via Denticon headers);
  lender "PatientFi Yes / No" without "Ran" (handled).
- **Bixby Knolls**: second half of September pasted two columns to the right with team columns not filled in — reads
  50 patients now (was undercounted as 21). **The log still has the old undercounted Bixby numbers → user should
  re-upload Bixby.** Bixby's second section has no "Total for Office" line.
- **Anaheim** workbook in the log was uploaded before most workbook metrics existed → re-upload to fill them.
- Exec Summary (thru 9/20), refunds sheet (as of 9/25) and workbooks (thru ~9/26–9/30) have different dates, so mixed
  rates (net collections, refund %, irrelevant %, % applying) mix date ranges until month-end reports.

## 12. Open questions / offered but not built

- Should "DENTICAL" (Denti-Cal) count as Medi-Cal? (Recommended yes; awaiting answer.)
- Subprime approved: a lender named with no amount (e.g. one El Segundo row) currently counts — user hasn't said to change it.
- Total production: include "Non visit code transactions" (current) or not?
- Two "NP collection" numbers exist (Referral report vs workbook) — user asked which is intended; not resolved.
- Offered: warning when uploading a workbook for an office with a connected sheet; warning when an uploaded office
  name doesn't match known offices; Undo button in the Changes tab; default column order matching the user's data
  dictionary (Office, Total Collection, Refunds, Net, %, Production, NP Coll, Existing Pt Coll, Insurance Coll, Ortho Coll).
- Rewrite README; commit + push (ask first).

## 13. Working conventions (for Claude)

- **Never print patient names/IDs.** Inspect workbooks with counts, value shapes, row numbers only. Hover details store
  sheet row numbers and short matched phrases only.
- Test changes on a **copy**: `cp` the app files + `data/reports.db` into the scratchpad, run with `PORT=8799`; for Google,
  a local stand-in server via `GOOGLE_TOKEN_URL` / `GOOGLE_SHEETS_API` env vars.
- Back up `data/reports.db` before any direct data fix.
- Careful with `pkill -f pattern` — it can kill the calling shell if the pattern appears in the command; find PIDs by
  checking `/proc/<pid>/cwd` instead.
- Don't enter credentials (e.g. the Google key) for the user; tell them where to click.
- User prefers asking clarifying questions via multiple-choice before building, then a concise summary with a table of
  real numbers and what changed.
