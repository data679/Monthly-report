# Product

<!-- impeccable:product-schema 1 -->

## Platform

web

## Users

- **The finance/data analyst at Care Dentistry Group (the owner of this app).** Works at a desk on this computer. Enters the monthly executive summary reports for 15 dental offices, checks the numbers, compares months and offices, pulls Invoca call data and Denticon patient lists, and answers questions from management and marketing.
- **Owners and executives.** Open the app through the Tailscale viewer link, often on a phone or tablet, to read totals, the Charts report (new-patient collections by ZIP) and the Calls funnel. View-only.
- **Office managers.** Look up their own office's months and audit results.
- **Marketing (Jerry).** Reads the ZIP collections report and the calls funnel to target Google Ads.

## Product Purpose

A local web app ("Executive Summary Logger") that turns Denticon/PlanetDDS reports, pasted spreadsheet rows and each office's New Patient Analysis sheet into one logged table per office per month, with history, checks, benchmarks and exports. Later additions: a Charts report of new-patient collections by home ZIP code with a ZIP-area map, and a Calls tab that joins Invoca call data to Denticon new-patient lists into a funnel (calls → answered → became new patients → started treatment → dollars collected). Success means the analyst can log and verify a month quickly, and management can read the state of the business without asking.

## Positioning

The only place where the group's monthly office numbers, new-patient geography and marketing call outcomes sit in one table with the audit trail of where each number came from (hover lists show which sheet rows were counted; every change is logged). Patient privacy is built in: phone numbers are never stored or shown (keyed fingerprints only), the repo holds no data.

## Operating Context

- Runs as an always-on systemd user service on the analyst's VM, reached locally and through Tailscale Serve by devices on the tailnet. App password for everyone; a separate viewer password gives a view-only mode that hides admin parts.
- Admin tools (SQL, Google Sheets, Audit, API keys, map setup) are visible only from this computer's own browser.
- Monthly ritual: the analyst uploads or pastes Denticon reports for each office, checks the Log, runs checks and benchmarks, exports CSV or copies rows into Google Sheets. Calls are pulled from Invoca per day or per month; Denticon patient lists are uploaded to match callers to patients.
- Reading ritual: executives open the Charts report and the Calls funnel on the viewer link, often on a phone.
- **Preview before commit (confirmed 2026-10-15):** the user wants to see a rebuilt UI as a preview that runs a step behind the live site, and only switch the live site over when they approve. The live service serves static/index.html from disk, so a new UI must live at a separate path until approved.

## Capabilities and Constraints

- Python 3.13 standard library only (no pip), SQLite, single-page UI in static/index.html with vanilla JS; page edits are live without restart, server edits need a service restart.
- Tabs today: Log, Charts, Calls, Changes, Upload, Google Sheets (local-only), SQL (local-only), Audit (local-only).
- Everything in the Log (section toggles, column picker, checks, benchmarks, office order, copy, CSV), the Charts report (tiles, month bars, top ZIPs, map card, ZIP table, office cards) and the Calls tab (funnel tiles, breakdown table, day sheet, promo-line pie, caller map, uploads, pulls) must keep working; layouts and controls may change.
- Leaflet 1.9.4 from cdnjs with SRI is the only external script; no other CDN dependencies, no build step.
- Terminology: Office, Period (half-month or month), Executive Summary, Collections, Ortho, New patients, Prime/Subprime (financing), Treatment, Data checks, Promo line (Invoca tracking number), Became new patients, Started treatment.
- Public GitHub repo: no patient data, credentials or office financial figures in code or docs.

## Brand Commitments

- Name: Executive Summary Logger, for Care Dentistry Group. No logo or brand assets exist in the repo.
- Confirmed preference (2026-10-15): a familiar, clean working-tool dashboard done at a high level of craft, in the category standard played straight. Flashy or decorative treatment would feel wrong. Peers whose craft sets the bar: Stripe Dashboard and Google Analytics / Looker Studio.

## Evidence on Hand

- Real data in data/reports.db (not in the repo): 15 offices, Jan–Oct 2026 logs, 17,500+ new patients with ZIPs, 25,000+ Invoca calls Aug–Oct 2026. The UI always renders real figures; nothing needs to be invented.
- No testimonials, screenshots or marketing copy exist and none should be fabricated.

## Product Principles

1. Numbers first: the figure and its source outrank any decoration.
2. One vocabulary everywhere: the same table, tile, control and state language on every tab.
3. Readable on a phone for executives, dense on a desk for the analyst.
4. Privacy by construction: never surface phone numbers or patient names.
5. Familiar over surprising: a category-fluent user should trust it at first glance.
