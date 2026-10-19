---
version: 1
slug: "static-index-html"
primary_target: "static/index.html"
related_targets: ["static/preview.html"]
---

# Surface brief: the whole app shell (Log, Charts, Calls, Changes, Upload, admin tabs)

Scope: replacement UI for the single-page app, built first at static/preview.html (served at /preview) and swapped into static/index.html when the user approves. Visitor mode: Operate.

Audience and job: the finance analyst at a desk logs and checks office months, pulls calls and uploads patient lists; owners, office managers and marketing read totals, the Charts report and the Calls funnel, often on a phone through the viewer link.

Task and content: every existing function stays (Log table with section toggles, column picker, checks, benchmarks, office order, copy, CSV; Charts report with tiles, month bars, top ZIPs, map card, ZIP table, office cards; Calls funnel, breakdown, day sheet, pie, caller map, uploads, pulls; Changes; Upload; Google Sheets, SQL, Audit local-only). All element ids and JavaScript behavior are preserved; only markup around them, CSS and small template strings change.

Constraints: standard library server, no build step, no new CDN scripts beyond Leaflet; no patient names or phone numbers anywhere; viewer mode hides admin and map setup; dark mode must keep working; the user chose the category standard (Stripe Dashboard, Google Analytics / Looker Studio as the craft bar) and rejected anything flashy.

Unresolved: none blocking. Sidebar collapse width and bottom-nav labels are mine to decide.

## Direction contract

THESIS: One working screen for a dental group's finance data in the Stripe / Looker canon: a quiet dark sidebar that says where you are, a page header that carries the scope (office, month) and the actions, and figures before anything else. It refuses the current arrangement of a text banner, eight flat tabs and a single scrolling column, and it refuses decoration of any kind.

OWN-WORLD: Cool slate-tinted canvas with white content cards and hairline borders; a committed deep-navy sidebar shell carrying the product mark and grouped navigation; one blue-violet accent for primary actions, current selection, links and focus; state colours do jobs only (green paid/ok, amber pending/warning, red error/refund); the existing geographic office colours remain the data series. One system sans family with tabular numerals, a fixed rem scale at a 1.2 ratio (12, 13, 14, 17, 20, 24), 8-px spacing grid, 6-px radii, 1-px borders, skeleton rows while loading, empty states that explain.

STORY: A reader knows the product and the group from the sidebar mark, sees which office and month are in scope from the header, reads the tiles, then the table, and trusts it at a glance because every tab speaks the same table, tile and button vocabulary.

FIRST VIEWPORT: 1440x900. Left: a 232-px navy sidebar with the mark at top ("Executive Summary Logger", "Care Dentistry Group" under it), then nav groups Work (Log, Changes, Upload), Reports (Charts, Calls), Admin (Google Sheets, SQL, Audit, local only), current item filled with the accent; the viewer and local-only states fold groups away. Right: a 56-px page header with the page title at left and the primary action group at right (Columns, Checks, Benchmarks, Office order as quiet buttons; Copy and CSV as secondary). Under it a filter bar with the office search and the month stepper, then section chips, then the Log table card filling the rest of the viewport with sticky column headers and sticky office column. Signature interaction: the month stepper and office filter in the header re-render the whole page in place with a 180-ms skeleton swap, never a reload or a blank. Motion grammar: 150 to 200 ms ease-out on state changes only; no entrance choreography.

FORM: the category standard, played straight (the standing exit, taken by the user); it sits outside the ordered grounded list. seed:76ed6933b00731a1

FINISH: unreviewed and undocumented is unfinished; this build ends with the finish review, the verdict, DESIGN.md, and every shipping raster carrying its provenance
