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
- **Runs as an always-on systemd user service** `monthly-report` (`~/.config/systemd/user/monthly-report.service`,
  Restart=always, no HOST override → listens on 127.0.0.1:8765 only). Reachable from this computer and, through
  **Tailscale Serve** (`sudo tailscale serve --bg 8765`, persists across restarts), at https://<node>.<tailnet>.ts.net (`tailscale serve status` prints the real address)
  for devices on the tailnet. Not on the office Wi-Fi (set `Environment=HOST=0.0.0.0` in the unit to allow that).
  App password required for everyone (salted PBKDF2 hash in `data/app_password.json`; change with
  `python3 app.py --set-password`, then restart). Admin parts (`LOCAL_ONLY_PATHS`: SQL, Google/Denticon keys,
  Audit, notes sheet, map setup) only from this computer's own browser; requests forwarded by Tailscale carry
  forwarding headers and are treated as other computers.
- **Restart / load code changes:** `systemctl --user restart monthly-report`; logs: `journalctl --user -u monthly-report`.
  Page-only changes (static/index.html) need just a hard refresh (Ctrl+Shift+R); served with `Cache-Control: no-store`.
- **Don't also run `./run.sh`** (same port; it was the cause of three outages after terminal closes / reboots).
  run.sh is only for running it by hand with the service stopped.
- Starts at login; to start at boot without logging in: `sudo loginctl enable-linger morg` (needs the user's sudo).
  The VM is VirtualBox: the host must stay on and the VM running.
- Python 3.13, **standard library only** (no pip on this machine). Needs `pdftotext` (poppler) for PDFs and `openssl`
  for Google sign-in.

- **Daily backups** (`backup_loop` thread, not in view-only copies): once a day, the first hourly check copies
  `data/reports.db` to `data/backups/reports-YYYY-MM-DD.db` with SQLite's backup API, runs `PRAGMA integrity_check`
  on the copy, and keeps the newest 30 (`KEEP_BACKUPS`). Days the app isn't running get no backup (nothing changes
  then). Restore: stop the app, copy a backup over `data/reports.db`, start it. Backups are on the same VM disk;
  an off-machine copy was suggested.

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
- **No future dates:** a Thru / as-of date after today is refused (uploads, refunds paste); the sync caps a typo'd
  future date on a tab at today. Date pickers stop at today.
- **Lower or same-date-changed numbers need confirming** (uploads + refunds paste, not the Google sync): the save
  returns `needs_review` with the list (was → now) and nothing is written; the UI shows "Check before saving" and
  re-sends with `confirm: true` on "Save anyway". `no_next_visit` and `collected_over_approved` are exempt (can go down).
- An older date than the log has is skipped per office/number, with a note to ask whoever manages the app if the
  later one was a mistake (fixed by hand from history, e.g. the 2026-10-30 NP collection mix-up on 10/05).
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

- **% TO GOAL sheet (refunds)**: adding a link that isn't a New Patient Analysis tries it as the % TO GOAL
  sheet (`sheet_sources.kind = 'refunds'`, office NULL). Sync reads the newest 2 month tabs with a `refunds`
  header (office names in column A, header there is just "."), stops at the **Total** row (a "Max Adwords" table
  below uses the same column), maps names through `office_alias` (unmatched names are listed in the status), and
  files each tab under **its own month**: as_of = min("Updated m/d/yy", month end, today); a tab whose Updated date
  is before its month (copied template) is skipped. Manual paste uses the same rules and shows "Will be logged for
  <month>".
- Word uploads: `.doc` (Word 97: UTF-16 text runs read straight from the file, no LibreOffice — it crashed this
  1-CPU VM) and `.docx` (zipped XML) go through the same row readers as PDFs.

- **Load past months** (Google Sheets tab): `POST /api/google/sync {"since": "2026-01"}` reads every month tab from
  that month on for every connected sheet (NP + refunds); the regular sync still reads the newest 2. Google allows
  ~60 reads/min, so `google_sheets._request` paces requests (1.1 s apart, `GOOGLE_MIN_GAP`) and retries 429/503.
  A full load from Jan 2026 takes ~3–6 minutes. Tabs for months that haven't started are skipped.
- Tab names: `APR'26` style years are read; a tab with **no year** ("October 🎃", "Feburary 💟") is taken as the
  latest such month up to today unless a tab with a year already covers it (Cerritos' old 2023 "SEPTEMBER" etc.).
- Tab order in the UI: Log (opens first), Charts (report + map), Calls, Changes, Upload, Google Sheets, SQL, Audit (local only).
  There is no Maps tab any more: the map is a card inside the Charts report (`#mp-section`), with the Map setup card
  (`#mp-admin`, local only) at the bottom of the Charts tab.
- **Charts tab** = the **collections report** (index.html `loadCharts` → `renderReport`/`renderReportTable`/
  `renderHeatmap`; server `maps_trends` → `GET /api/maps/trends`, also allowed on view-only links: ZIP-level counts
  only). Per ZIP × month × office × General/Ortho: patients, paid patients, $ collected, plus each month's latest
  Thru date (`thru`; a month whose Thru is before its last day is **partial**: shown dashed/italic, left out of
  counts and totals), office `pins` and ZIP `centers` (nearest office, computed on the page). Filters: Office
  (patients who visit it) + From/To month (`zipReportSel`). Sections: lead sentence + 4 tiles (collected, patients,
  ZIPs, ZIPs for half the $ / top-10 share); collected per month (bars, partial dashed); top 25 ZIPs by $ (bars);
  **the map** (see Maps below; its Office comes from the report's Office filter, its Month select stays inside the
  report's range); top 10 ZIPs month by month (small multiples, one scale); every ZIP as a sortable/searchable table
  with Copy; each office's top 5 ZIPs; footnote. A ZIP link scrolls to the map and zooms it on that ZIP (`mapFocusZip`). The
  same report exists as a shared artifact ("ZIP Collections Report", a snapshot). The earlier office charts
  (trend / ranking / heatmap by office) and the ZIP-trends view (bump chart, per-month lists) were removed at the
  user's request (Oct 2026).

- **Denticon lender check** (lender_journal.py + `denticon_checks` in app.py): a Daily Journal with Patient Type
  **Both** and Transactions **Payments, Adjustments** (Word or PDF) is read as `lender_journal`: per patient,
  each lender's payment and merchant fee (Alphaeon, Care Credit, PatientFi, Proceed = prime; Access Alph, HFD,
  Covered Care, Cherry, Sunbit = subprime). Detail must add up to Denticon's own "Summary for <office>" per lender
  or the file is refused. Ortho journals still go to ortho collection.
  - **Pat IDs are never stored or sent to the browser**: `pat_key` = HMAC-SHA256 with `data/patient_key` (0600);
    `/api/parse` scrubs; storage scrubs again. Tables: `np_patients` (each NP row's collected per lender, saved with
    every sync/upload of the NP sheet), `lender_payments`, `journal_uploads` (one per office+month, later Thru wins).
  - Comparison: sheet "$Amt Coll" (before the fee) vs Denticon payment + fee, summed over all uploaded months;
    tolerance $1. "Denticon more" always flagged; "sheet more" only when the journal reaches the sheet's date and
    the next month (if started) is uploaded. A fee booked under another lender with no payment of its own is
    paired with the one lender that paid without a fee (Alphaeon payment + "ACCESS ALPH" fee). Results are checks
    (`den:<row>:<lender|Subprime>`, kind = missing_sheet / missing_denticon / amount / lender / odd) with Mark OK.
  - The sheet's amount after the fee (= the payment alone) also counts as a match; a row whose total matches
    under different lenders is one "different lender" note; under $50 with no Denticon payment is "odd value".
  - **Admin only:** results are on the **Audit** tab (`/api/audit`), which, like saving a lender journal, only
    works from the computer running the app (`LOCAL_ONLY_PATHS`); the tab stays hidden elsewhere and in view-only
    links. They are not in the Log's ⚠ flags or the Checks panel. Sept 2026: 1,585 NPs checked, 124 differences.

- **Invoca API** (invoca_api.py, Audit tab card "Invoca API (call tracking)", local-only endpoints
  `/api/invoca/status|save|delete|test`): Transactions API v2020-10-01,
  `https://<site>.invoca.net/api/2020-10-01/<networks|advertisers>/transactions/<id>.json`, token in the
  `Authorization` header (never in URLs), `from`/`to` (YYYY-MM-DD, inclusive), `limit` ≤ 4000, paging with
  `start_after_transaction_id`, `include_columns` = `invoca_api.COLUMNS` (ids, times, durations, advertiser/campaign,
  promo line, city/region, result, `custom_data` for signals; **no caller phone numbers**). Config + token in
  `data/invoca_api.json` (0600). `invoca_test()` reads the last 7 days and summarizes calls by advertiser/campaign
  and the signal names seen. Docs: developers.invoca.net (Transactions API → Network/Brand, Advertiser/Merchant;
  Manage API Credentials: Integrations → Manage Integrations → Invoca APIs → + New API Credential). Nothing is
  stored from the test beyond its summary (in memory).
- **Calls tab** (`#calls`, hidden on view-only links): one day's calls from Invoca as a sheet. `POST /api/calls/pull`
  {date} (app password; any day up to today) → `pull_invoca_calls` → `all_transactions(day, day)` →
  `store_invoca_calls` keeps Call transactions in table `invoca_calls` (replace that day; signals = true Signal
  entries from the call's custom_data plus Signal transactions with the same complete_call_id; other marketing data
  as JSON `custom`; never phone numbers). `GET /api/calls?date=` → that day's rows + the list of pulled days.
  Page: Day (default yesterday) + Pull button + Office + Search + Copy; tiles (calls, answered %, avg talk, repeat
  callers, top signals); sortable sheet (time, office, campaign, promo line, city, state, phone type, repeat,
  rang+talked, talked or "not answered", outcome, signals/marketing chips); a pie of calls by promo line (≤ 8
  slices, rest "Other") and a callers' map (`drawCallsMap`: a second Leaflet map on the ZIP outlines, one circle per
  caller city sized/shaded by calls, the 12 busiest named; `calls_info` returns `points` = city → center from
  `city_names.json` for California cities by name, other cities listed in the note). Both follow the Office/Search
  filters.
- **Call → new patient matching** (privacy by fingerprint): `invoca_api.COLUMNS` includes `calling_phone_number`
  only so `store_invoca_calls` can store `phone_fp()` of it (HMAC with the patient_key secret over the 10 digits;
  the number itself is never kept or shown). Denticon's **New Patients Through Appointment Book** (Excel, one sheet
  per office: Patient Name, DOB, Home/Work/CellPhone, Email, Appt Date) is uploaded on the Calls tab
  (`POST /api/calls/appointments` → `summary_parser.parse_new_patient_appts` → `save_np_appts`): only the office,
  appointment date and phone fingerprints go into `np_appts`; names, DOB and emails are never read into the result;
  extensions ("x: 12") are stripped. `calls_info` joins each call's fingerprint to `np_appts`: an appointment on or
  after the call day = "converted" (shown green "booked m/d · office"), earlier = "booked earlier"; tile "became new
  patients (% of answered)". The report has no Pat ID, so the link is phone-only; run it for a wide date range.
  **Preferred source: Denticon's "New Patients List"** (one office per file: Pat ID, Patient Name (Chart No), DOB,
  Phone# (one cell "(H) : … (C) : … (W) : …"), Email, Prdr, Referred By, Created On, First visit, Type, General/Ortho)
  → `parse_new_patients_list` → `save_patient_phones` → table `patient_phones(fp, pat_id, office, first_visit,
  created_on, phone_kind)`. The same Calls-tab button (`POST /api/calls/upload`, `calls_upload_kind` tells the two
  reports apart) takes either report. `calls_info` match priority: **visited** (patient_phones first_visit ≥ call
  day; `collected` = that patient's Referral-listing collection for the first-visit month, from map_patients by
  pat_id + office) > **converted** (np_appts appointment ≥ call day) > **existing** (first_visit < call day) >
  **before** (appointment < call day). The sheet shows "came in m/d · office · $x" (Pat ID in the hover title),
  "booked m/d · office", "existing patient", "booked earlier"; the tile adds "N came in ($ collected), M booked".
  Placeholder numbers (one repeated digit, 000…, 123-456-7890) are skipped everywhere (`phone_fp`).
- **Funnel** (`calls_funnel(month, office)` → `GET /api/calls/funnel`; card `#fn-card` at the top of the Calls tab):
  for a call month: calls, answered, distinct callers, became new patients (first visit ≥ call day on a loaded
  patient list; each patient once, credited to their first qualifying call's promo line/campaign), started
  treatment (collected > 0 that first-visit month), $ collected, booked-not-yet-seen (np_appts); breakdown by
  promo line or campaign. The office filter applies from the patient step (calls are group-level in Invoca).
  `pull_invoca_range(frm, to)` backfills a date range in one paged fetch (Aug 1–Sep 30 2026 were backfilled).
  Next: a nightly pull and calls per office per month next to new patients.
- **Denticon API** (denticon_api.py, Audit tab, local-only): base `https://api.planetdds.com/denticon`, header
  `PDDS-Subscription-Key` (key in `data/denticon_api.json`, 0600, entered by the user in the app). Step 1 built:
  save key, Test connection (Practices `/practices/v0/offices`), match Denticon offices to log offices
  (`denticon_offices`). Step 2 planned: nightly RCM `/rcm/v0/ledgers/{OfficeId}` (LastChangedOn ≤ 30-day windows,
  PageSize 1000; ledgerType C/P/I/A, description, amount, patientId) → lender payments + fees per hashed patient,
  replacing Daily Journal uploads for the Audit tab. 429 responses say "Try again in N seconds" (handled).

- **Notes sheet** (Google Sheets tab › 4, admin/local-only): the app copies the Log (every office × month, the
  Log's visible columns read from `LOG_COLS`/`SECTIONS` in index.html, internal ones skipped) to **one tab per month**
  ("October 2026", newest first, Total row calculated like the page, rates from `ratio`) of a
  Google Sheet the user shares with the service account as **Editor** (write scope `spreadsheets`; office sheets
  stay Viewer, so still read-only). One way: month tabs are cleared and rewritten (batchGet/batchClear/batchUpdate,
  3–4 requests per update), except each tab's **Notes** column, read first and re-attached by office; notes for rows
  that disappear go to "Notes (orphaned)". The earlier one-tab "Log" layout is migrated and its tab removed.
  `notes_sheet_loop` checks a fingerprint every 2 minutes and writes only when the log changed; Write now forces.
  Table `notes_sheet`; endpoints `/api/google/notes-sheet[/write|/delete]`.

- **The map** (formerly the Maps tab; now the `#mp-section` card in the Charts report; everyone with the password,
  hidden on view-only links): new patients by ZIP for a month. Its Office comes from the report's Office filter; the
  Maps ZIP table was removed (the report's table covers it).
  - Source: the Referral Production Listing run **grouped by Zip Code** (Excel best; PDF only where page breaks
    don't scramble lines). `parse_referral` → `_referral_patients` (Pat ID, zip, referral type, visits,
    production, collection; never names), attached as `details._map` only if an office's patients add up to its
    total → `map_patients` (replaced per office × month, later as_of wins).
  - General/Ortho: Patient List – Address upload (`parse_patient_types` → `patient_types`, Pat ID → type only).
  - ZIP centers: `data/zip_centroids.json`, refreshed from the boundary download (Map setup step 1); the older
    gazetteer download (`/api/maps/zip-download`, `download_zip_centroids`) still exists as a fallback with no button.
    Office pins: `office_locations` (US Census geocoder for office addresses, or lat/lon typed in).
  - UI: Leaflet (cdnjs, loaded only on the tab). **No street tiles**: plain background, every ZIP in the region
    drawn as an area (US Census TIGERweb ZCTA GeoJSON cached in `data/zip_shapes.json` via Map setup step 1 →
    `/api/maps/shapes-download`, local-only; served to the page by `GET /api/maps/shapes`; ZCTA POP100 gives the
    population for "per 10k residents", so no Census key; ZIPs the Census has no area for, e.g. PO boxes, are
    remembered in `data/zip_no_zcta.json` so the page says "not on the map" rather than "run step 1 again"), ZIPs
    with patients filled by the measure (new patients,
    $ collected, per 10k, dominant office; light-yellow→dark-red ramp, top 5% capped), white = no patients; a
    labels-only Esri overlay (`Reference/World_Boundaries_and_Places`, pane "labels" z450) for city names, inverted
    in dark mode; ZIP numbers as permanent tooltips at zoom ≥ 11 (busiest 150 in view). Numbered office pins
    (short names from office_alias via `office_short_names`; offices < 0.7 mi apart share a "5·13" pin), legend,
    optional 3/5/8-mile rings (straight-line; nearest/outside measured against the same offices the rings are drawn
    for; "not nearest" flag when a ZIP's patients mostly visit a farther office), inset mini-map (light-gray base
    tiles, bottom right) with the main view's box, auto-fit to the ZIPs holding 80% of patients within 35 miles of
    the patient-weighted median (nearby pins within 1.1× that reach). ZIPs with patients but no boundary (outside the
    downloaded region) show as dots and are listed in the note. The map is limited to the **map area**
    (`map_region()`: the office pins padded by `MAP_PAD` = 0.5°, about 35 miles; before any pin, a box around the
    patients' median): zoom out stops where the area fits, zoom in at 13, panning can't leave it, and the inset shows
    exactly that area. Patients in ZIPs outside the area (other counties/states) stay in the totals and the ZIP table
    with an "outside map" flag and are listed in the note under the map; they are never drawn.
    Color classes (`classBreaks`): five data-driven classes (quantiles of the ZIPs, top class from the 95th
    percentile when ≥ 20 ZIPs, whole numbers for counts, fewer classes when values repeat); the legend (top-right,
    collapsible, folded by default on phones) lists each class's real range. Office colors (`assignOfficeColors`)
    are stable per office and chosen so offices within 10 miles get the most different hues; mixed ZIPs get a dark
    dashed outline. The numbered office key sits under the map (`#mp-key`). The map canvas stays light in dark mode.
    City names: Step 1 also downloads the area's cities (Census 2020 incorporated places + CDPs, layers 25/26 of
    TIGERweb Places, with population) and works out which city each ZIP center and office pin falls in (point in
    polygon, done once; outlines not kept) → `data/city_names.json` {"cities": {name: {c, pop, cdp}}, "where":
    {zip|"office:<name>": city}}. The page (`drawCityLabels`) names only cities with an office or with patients on
    show, thinning by population when zoomed out; without that file it falls back to Esri's labels-only tiles. ZIP table adds per-10k, nearest office (miles) and
    multi-office flags; hover lists Pat ID · office · $ (General/Ortho column only once a patient list is loaded).

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
