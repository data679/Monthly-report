---
name: Executive Summary Logger
description: One working screen for a dental group's finance data, in the Stripe / Looker canon, played straight.
colors:
  # Canvas and surfaces (light mode; dark values are listed in Colors and in .impeccable/design.json)
  bg: "#f4f6f9"
  surface: "#ffffff"
  surface-2: "#f8fafc"
  hover: "#f0f3f7"
  code-bg: "#f3f5f8"
  # Ink
  text: "#1a1f36"
  muted: "#5b6678"
  faint: "#687283"
  # Lines
  border: "#e2e7ee"
  border-strong: "#c9d2dd"
  # The one accent
  accent: "#3b5bdb"
  accent-hover: "#3350c4"
  accent-soft: "#e8edfb"
  accent-ink: "#2b45b3"
  accent-text: "#ffffff"
  # State colours that do jobs
  good: "#177245"
  good-soft: "#e3f4ea"
  warn: "#a24e0b"
  warn-soft: "#fcf0e2"
  bad: "#b92f23"
  bad-soft: "#fcebe9"
  # Chart series (series tracks the accent)
  series: "#3b5bdb"
  series-dim: "#c7d0f3"
  # Sidebar shell (identical in both modes)
  side: "#121a2d"
  side-2: "#1b2540"
  side-active: "#27345a"
  side-current: "#1f2f65"
  side-text: "#c4ccdd"
  side-muted: "#7f8aa4"
  side-bright: "#ffffff"
  # Log column-group bands
  sec-1: "#2f5fd4"
  sec-2: "#0b7a71"
  sec-3: "#b02bc4"
  sec-4: "#a85a05"
  sec-5: "#137a3c"
  sec-6: "#c92a2a"
  sec-7: "#6d3fd1"
  sec-8: "#086d8e"
  # Benchmark cell fills and their ink
  bm-green: "#dff5e7"
  bm-green-ink: "#14532d"
  bm-yellow: "#fcefc7"
  bm-yellow-ink: "#7a4a06"
  bm-red: "#fde3e1"
  bm-red-ink: "#8a1c13"
  # Data-check flag
  chk: "#c2410c"
  chk-ink: "#ffffff"
  # Map canvas (fixed; stays light in dark mode)
  map-canvas: "#dfe3e8"
  map-pin: "#111827"
  map-label: "#1f2937"
  map-outline: "#d1d5db"
  map-empty: "#f3f4f6"
  map-no-figure: "#9ca3af"
  map-inset-bg: "#e5e7eb"
  map-inset-rect: "#1d4ed8"
  # Dark mode (prefers-color-scheme: dark) values of the tokens that change; the sidebar and map tokens do not
  bg-dark: "#0f141b"
  surface-dark: "#171d27"
  surface-2-dark: "#1c232e"
  hover-dark: "#1f2733"
  code-bg-dark: "#121820"
  text-dark: "#e7ebf2"
  muted-dark: "#9aa5b6"
  faint-dark: "#8d98aa"
  border-dark: "#28313e"
  border-strong-dark: "#3a4556"
  accent-dark: "#7f93ff"
  accent-hover-dark: "#93a4ff"
  accent-soft-dark: "#232c4a"
  accent-ink-dark: "#a9b6ff"
  accent-text-dark: "#0b1220"
  good-dark: "#4ade80"
  good-soft-dark: "#15321f"
  warn-dark: "#fbbf24"
  warn-soft-dark: "#3a2a0c"
  bad-dark: "#f87171"
  bad-soft-dark: "#3d1a17"
  series-dark: "#7f93ff"
  series-dim-dark: "#2f3a63"
  sec-1-dark: "#6f9cff"
  sec-2-dark: "#3dd3c3"
  sec-3-dark: "#e68cf5"
  sec-4-dark: "#fbbf24"
  sec-5-dark: "#4ade80"
  sec-6-dark: "#f87171"
  sec-7-dark: "#b4a0ff"
  sec-8-dark: "#38d1f0"
  bm-green-dark: "#14532d99"
  bm-green-ink-dark: "#86efac"
  bm-yellow-dark: "#78350f99"
  bm-yellow-ink-dark: "#fcd34d"
  bm-red-dark: "#7f1d1d99"
  bm-red-ink-dark: "#fca5a5"
  chk-dark: "#fb923c"
  chk-ink-dark: "#1c1917"
typography:
  figure:
    fontFamily: "-apple-system, BlinkMacSystemFont, \"Segoe UI\", Roboto, \"Helvetica Neue\", Arial, sans-serif"
    fontSize: "24px"
    fontWeight: 600
    lineHeight: 1.15
    letterSpacing: "-0.01em"
    fontFeature: "tabular-nums"
  page-title:
    fontFamily: "-apple-system, BlinkMacSystemFont, \"Segoe UI\", Roboto, \"Helvetica Neue\", Arial, sans-serif"
    fontSize: "20px"
    fontWeight: 600
    lineHeight: "40px"
    letterSpacing: "-0.01em"
  card-title:
    fontFamily: "-apple-system, BlinkMacSystemFont, \"Segoe UI\", Roboto, \"Helvetica Neue\", Arial, sans-serif"
    fontSize: "16px"
    fontWeight: 600
    letterSpacing: "-0.01em"
  body:
    fontFamily: "-apple-system, BlinkMacSystemFont, \"Segoe UI\", Roboto, \"Helvetica Neue\", Arial, sans-serif"
    fontSize: "14px"
    fontWeight: 400
    lineHeight: 1.45
    fontFeature: "tabular-nums"
  nav:
    fontFamily: "-apple-system, BlinkMacSystemFont, \"Segoe UI\", Roboto, \"Helvetica Neue\", Arial, sans-serif"
    fontSize: "14px"
    fontWeight: 500
  control:
    fontFamily: "-apple-system, BlinkMacSystemFont, \"Segoe UI\", Roboto, \"Helvetica Neue\", Arial, sans-serif"
    fontSize: "13px"
    fontWeight: 500
    lineHeight: 1
  table:
    fontFamily: "-apple-system, BlinkMacSystemFont, \"Segoe UI\", Roboto, \"Helvetica Neue\", Arial, sans-serif"
    fontSize: "13px"
    fontWeight: 400
    lineHeight: 1.45
    fontFeature: "tabular-nums"
  label:
    fontFamily: "-apple-system, BlinkMacSystemFont, \"Segoe UI\", Roboto, \"Helvetica Neue\", Arial, sans-serif"
    fontSize: "12px"
    fontWeight: 500
  column-head:
    fontFamily: "-apple-system, BlinkMacSystemFont, \"Segoe UI\", Roboto, \"Helvetica Neue\", Arial, sans-serif"
    fontSize: "12px"
    fontWeight: 600
    letterSpacing: "0.05em"
  mono:
    fontFamily: "ui-monospace, \"SF Mono\", Menlo, Consolas, monospace"
    fontSize: "12.5px"
    fontWeight: 400
    lineHeight: 1.5
rounded:
  tick: "1px"
  micro: "2px"
  mark: "3px"
  code: "4px"
  stepper-inner: "5px"
  control: "6px"
  inset: "8px"
  card: "10px"
  pill: "999px"
spacing:
  xs: "4px"
  icon-gap: "6px"
  sm: "8px"
  toolbar-gap: "10px"
  md: "12px"
  lg: "16px"
  xl: "24px"
  page-bottom: "48px"
components:
  sidebar:
    backgroundColor: "{colors.side}"
    textColor: "{colors.side-text}"
    width: "232px"
    padding: "16px 8px 8px"
  brand-mark:
    backgroundColor: "{colors.accent}"
    textColor: "{colors.accent-text}"
    rounded: "{rounded.inset}"
    size: "30px"
  nav-group:
    textColor: "{colors.side-muted}"
    typography: "{typography.column-head}"
    padding: "16px 12px 4px"
  nav-item:
    backgroundColor: "transparent"
    textColor: "{colors.side-text}"
    typography: "{typography.nav}"
    rounded: "{rounded.control}"
    padding: "8px 12px"
  nav-item-hover:
    backgroundColor: "{colors.side-2}"
    textColor: "{colors.side-bright}"
  nav-item-active:
    backgroundColor: "{colors.side-current}"
    textColor: "{colors.side-bright}"
  page-title:
    textColor: "{colors.text}"
    typography: "{typography.page-title}"
    height: "40px"
  card:
    backgroundColor: "{colors.surface}"
    textColor: "{colors.text}"
    rounded: "{rounded.card}"
    padding: "16px"
  table-card:
    backgroundColor: "{colors.surface}"
    rounded: "{rounded.card}"
    padding: "0"
  button:
    backgroundColor: "{colors.surface}"
    textColor: "{colors.text}"
    typography: "{typography.control}"
    rounded: "{rounded.control}"
    padding: "0 12px"
    height: "32px"
  button-hover:
    backgroundColor: "{colors.hover}"
  button-active:
    backgroundColor: "{colors.border}"
  button-primary:
    backgroundColor: "{colors.accent}"
    textColor: "{colors.accent-text}"
    typography: "{typography.control}"
    rounded: "{rounded.control}"
    padding: "0 12px"
    height: "32px"
  button-primary-hover:
    backgroundColor: "{colors.accent-hover}"
  button-danger:
    backgroundColor: "{colors.surface}"
    textColor: "{colors.bad}"
  button-danger-hover:
    backgroundColor: "{colors.bad-soft}"
  button-quiet:
    backgroundColor: "transparent"
    textColor: "{colors.accent-ink}"
    typography: "{typography.control}"
    rounded: "{rounded.control}"
    padding: "0 12px"
    height: "32px"
  button-quiet-hover:
    backgroundColor: "{colors.accent-soft}"
  button-expanded:
    backgroundColor: "{colors.accent-soft}"
    textColor: "{colors.accent-ink}"
  button-compact:
    height: "28px"
    padding: "0 10px"
  link-button:
    backgroundColor: "transparent"
    textColor: "{colors.accent-ink}"
    typography: "{typography.control}"
    padding: "0"
  input:
    backgroundColor: "{colors.surface}"
    textColor: "{colors.text}"
    typography: "{typography.body}"
    rounded: "{rounded.control}"
    padding: "0 10px"
    height: "32px"
  textarea:
    backgroundColor: "{colors.code-bg}"
    textColor: "{colors.text}"
    typography: "{typography.mono}"
    rounded: "{rounded.control}"
    padding: "10px"
  field-label:
    textColor: "{colors.muted}"
    typography: "{typography.label}"
  month-stepper:
    backgroundColor: "{colors.surface}"
    textColor: "{colors.text}"
    rounded: "{rounded.control}"
    height: "32px"
  popover:
    backgroundColor: "{colors.surface}"
    textColor: "{colors.text}"
    rounded: "{rounded.card}"
    padding: "10px"
    width: "260px"
  column-menu:
    backgroundColor: "{colors.surface}"
    rounded: "{rounded.card}"
    padding: "10px"
    width: "300px"
  month-grid:
    backgroundColor: "{colors.surface}"
    rounded: "{rounded.card}"
    padding: "10px"
    width: "250px"
  month-grid-selected:
    backgroundColor: "{colors.accent}"
    textColor: "{colors.accent-text}"
  chip-section:
    backgroundColor: "{colors.surface}"
    textColor: "{colors.muted}"
    typography: "{typography.control}"
    rounded: "{rounded.pill}"
    padding: "0 10px 0 8px"
    height: "26px"
  chip-section-on:
    textColor: "{colors.text}"
  tag:
    textColor: "{colors.muted}"
    typography: "{typography.label}"
    rounded: "{rounded.pill}"
    padding: "1px 7px"
  table-head:
    backgroundColor: "{colors.surface}"
    textColor: "{colors.muted}"
    typography: "{typography.column-head}"
    padding: "8px 12px"
  table-cell:
    textColor: "{colors.text}"
    typography: "{typography.table}"
    padding: "8px 12px"
  table-row-hover:
    backgroundColor: "{colors.hover}"
  log-cell:
    typography: "{typography.table}"
    padding: "6px 8px"
  benchmark-cell-green:
    backgroundColor: "{colors.bm-green}"
    textColor: "{colors.bm-green-ink}"
  benchmark-cell-yellow:
    backgroundColor: "{colors.bm-yellow}"
    textColor: "{colors.bm-yellow-ink}"
  benchmark-cell-red:
    backgroundColor: "{colors.bm-red}"
    textColor: "{colors.bm-red-ink}"
  check-count:
    backgroundColor: "{colors.chk}"
    textColor: "{colors.chk-ink}"
    typography: "{typography.label}"
    rounded: "{rounded.pill}"
    padding: "0 5px"
    height: "18px"
  stat-tile:
    backgroundColor: "{colors.surface}"
    rounded: "0"
    padding: "12px 16px"
  stat-tile-label:
    textColor: "{colors.muted}"
    typography: "{typography.control}"
  stat-tile-figure:
    textColor: "{colors.text}"
    typography: "{typography.figure}"
  skeleton:
    backgroundColor: "{colors.hover}"
    rounded: "{rounded.code}"
    height: "0.85em"
    width: "62%"
  empty-state:
    textColor: "{colors.muted}"
    padding: "44px 16px"
  tooltip:
    backgroundColor: "{colors.surface}"
    textColor: "{colors.text}"
    typography: "{typography.table}"
    rounded: "{rounded.inset}"
    padding: "10px 12px"
  drop-zone:
    textColor: "{colors.muted}"
    rounded: "{rounded.card}"
    padding: "36px 16px"
  help-dialog:
    backgroundColor: "{colors.surface}"
    textColor: "{colors.text}"
    rounded: "{rounded.card}"
    width: "min(760px, calc(100vw - 32px))"
  help-step:
    backgroundColor: "transparent"
    textColor: "{colors.text}"
    typography: "{typography.control}"
    rounded: "{rounded.pill}"
    padding: "5px 12px"
  help-step-selected:
    backgroundColor: "{colors.accent}"
    textColor: "{colors.accent-text}"
---

# Design System: Executive Summary Logger

## Overview

**Creative North Star: "One Working Screen"**

This is the finance desk of a dental group rendered as a single working screen in the Stripe Dashboard / Looker Studio canon, played straight. A committed navy sidebar says where you are; a page header carries the scope (office, month) and the actions; everything below is figures first: stat strips, then tables, then charts and maps. Density is desk-grade for the analyst (a 13-px Log with 6/8-px cells, sticky headers and a pinned office column) and still readable on a phone for the owners who open the viewer link, where the sidebar folds into a top bar with a scrolling tab row and the same cards stack.

The world is restrained on purpose. One system sans family in six sizes with tabular numerals; a cool slate canvas; white cards drawn with a single hairline; one blue-violet accent that marks the primary action, the current place in the sidebar, links and focus, and nothing else. State colours do jobs (green ok, amber warning, red error) and are paired with a shape or a word, never used alone. Nothing is decorated: no gradients beyond a flat 7% band tint, no drop shadows on resting surfaces, no entrance choreography. Motion is a 150-ms state change and a 180-ms fade when the month changes.

The rejections are confirmed by the user and by the build: nothing flashy, no second accent, no nested cards, no decorative shadows or gradients, no glyph icons, and never a patient name or phone number on screen.

**Key Characteristics:**
- A 232-px flat navy sidebar with grouped navigation (Work / Reports / Admin), the brand mark at top, help at the foot; it becomes a sticky top bar under 900 px.
- One accent (`accent`) for primary fills, the current nav item (a 32% accent tint over the navy), focus outlines and selection; its deeper text form (`accent-ink`) for links and quiet actions.
- Six type sizes (12 / 13 / 14 / 16 / 20 / 24 px) in one system sans with `tabular-nums` everywhere; monospace only for code and the paste / SQL textareas.
- Cards are a 1-px hairline on white; floating surfaces (menus, tooltips, the help dialog) carry the one `--shadow` and no border; in dark mode the shadow grows a 1-px ring.
- Every control is 32 px tall with a 6-px radius; compact variants are 28 px; chips and counts are pills.
- Tables with quiet uppercase 12-px headers, hairline rows, right-aligned tabular numbers; the Log pins its headers, its office column and its total row.
- Stat strips are one card divided by hairlines: a 13-px muted label over a 24-px figure, no tile radius, no tile shadow.
- Icons are an authored 16-px SVG set (inline for nav and a few controls, `mask-image` custom properties for the rest), drawn at 1.5 / 1.8 stroke, painted in `currentColor`.
- Light and dark token sets are both normative; the sidebar tokens and the map canvas do not change with the scheme.

## Colors

A cool slate canvas with white cards and hairlines, a constant deep-navy sidebar, one indigo accent, and state colours that only ever do a job.

### Primary
- **Indigo accent** (`accent`): the primary button fill, the selected month in the month grid, the selected help step, the brand mark, the global `:focus-visible` outline (2 px, offset 2), the text caret, native `accent-color`, and the text selection (a 22% tint). The current sidebar item is this colour mixed at 32% over the sidebar navy (`side-current`).
- **Accent hover** (`accent-hover`): the primary button's hover fill only.
- **Accent tint** (`accent-soft`): hover fill of quiet buttons and month-grid footer links, the fill of a button whose popover is open (`aria-expanded="true"`), and the drop zone when a file is over it.
- **Accent ink** (`accent-ink`): the text form of the accent on white. Links, `.btn.quiet`, `.link-btn`, `details > summary`, the benchmark target flag in a column header, and the "booked" state in the calls sheet. Links underline on hover only (offset 3 px).
- **Accent text** (`accent-text`): the ink on an accent fill. White in light mode; near-black (`#0b1220`) in dark mode because the dark accent is a pale periwinkle.

### Secondary (chart series)
- **Series** (`series`): the one data colour for bars; it tracks the accent in light mode (`var(--accent)`) and is the dark accent in dark mode. **Series dim** (`series-dim`) is the muted bar for non-top rows and the ZIP table's inline bars. Partial months draw the series at 45% opacity with a 4 / 3 dash.

### Tertiary (state and data colours)
- **Good** (`good`, `good-soft`): ok messages, "connected", the up arrow, the "ok" cell in the calls sheet. The soft fill is unused for cells today; it exists for the same job.
- **Warn** (`warn`, `warn-soft`): the warning tag and the confirm box (soft fill with a 45% warn border).
- **Bad** (`bad`, `bad-soft`): error messages, the down arrow, the danger button's text and hover fill, a `select.needs` border and its 18% ring, missing values in the calls sheet.
- **Data-check flag** (`chk` on `chk-ink`): the one orange in the system; the warning triangle beside a value that looks wrong and the count pill on the Checks button. Deliberately not `warn`, so a data check is never mistaken for a benchmark.
- **Benchmark fills** (`bm-green` / `bm-yellow` / `bm-red` with their `-ink`): soft cell fills with a dark ink of the same hue, always paired with a mark (a dot, an up triangle, a down triangle) so the state reads without colour. In dark mode the fills are 60%-alpha hex over the surface.
- **Section bands** (`sec-1` … `sec-8`): the Log's column-group colours, set as `--sec` on each group. They draw the 2-px top rule and 7% tint of a band header, the dot and tint of a section chip, and the tag in the column picker. They are labels, not states.

### Neutral
- **Canvas** (`bg`): the page behind the cards.
- **Surface** (`surface`): cards, popovers, inputs, the sticky table headers and the pinned Log column. **Surface 2** (`surface-2`): the benchmark panel, the path chips, the "sig" pills and the even rows of the calls sheet (at 70%).
- **Hover** (`hover`): row hover, button hover, menu-item hover, the skeleton bar.
- **Code background** (`code-bg`): inline code and the mono textareas.
- **Text** (`text`), **Muted** (`muted`, labels, table headers, hints, tile labels), **Faint** (`faint`, placeholders, as-of dates, tile footnotes, the "none" value, input hover borders).
- **Border** (`border`, every hairline: cards, rows, dividers) and **Border strong** (`border-strong`, the control border, total-row rule, chart axis, scrollbar thumb).
- **Sidebar** (`side`, `side-2` hover, `side-current` current item, `side-text`, `side-muted` group labels, `side-bright` for the brand name and hovered/current item text; `side-active` is the sidebar's scrollbar thumb). These do not change with the colour scheme.
- **Map** (`map-canvas`, `map-pin`, `map-label`): fixed hex on the Leaflet canvas, pins and city labels. The map stays light in dark mode because its colour ramp needs a light ground.

### Dark mode
Dark values are applied by `@media (prefers-color-scheme: dark)` only; there is no theme attribute. The sidebar and map tokens are unchanged. Everything else re-maps as follows (light values are the frontmatter):

| Token | Dark | Token | Dark |
|---|---|---|---|
| bg | `#0f141b` | accent | `#7f93ff` |
| surface | `#171d27` | accent-hover | `#93a4ff` |
| surface-2 | `#1c232e` | accent-soft | `#232c4a` |
| hover | `#1f2733` | accent-ink | `#a9b6ff` |
| code-bg | `#121820` | accent-text | `#0b1220` |
| text | `#e7ebf2` | good / good-soft | `#4ade80` / `#15321f` |
| muted | `#9aa5b6` | warn / warn-soft | `#fbbf24` / `#3a2a0c` |
| faint | `#8d98aa` | bad / bad-soft | `#f87171` / `#3d1a17` |
| border | `#28313e` | series / series-dim | `#7f93ff` / `#2f3a63` |
| border-strong | `#3a4556` | chk / chk-ink | `#fb923c` / `#1c1917` |
| sec-1 … sec-8 | `#6f9cff` `#3dd3c3` `#e68cf5` `#fbbf24` `#4ade80` `#f87171` `#b4a0ff` `#38d1f0` | bm-green / -ink | `#14532d99` / `#86efac` |
| side-current | `#354170` (32% dark accent over `side`) | bm-yellow / -ink | `#78350f99` / `#fcd34d` |
| | | bm-red / -ink | `#7f1d1d99` / `#fca5a5` |

### Named Rules
**The One Accent Rule.** Indigo is the only accent. It marks the primary action, the current place, focus, selection and links; it never colours data except as the single chart series, and no second accent is introduced for emphasis.

**The Jobs-Only Rule.** Green, amber and red appear only when they report a state (ok, warning, error / benchmark on target, close, below), and the state is also carried by a word, a mark or a shape. Decorative use of a state colour is a defect.

**The Constant Shell Rule.** The sidebar's five tokens and the map's three are the same in light and dark mode. Everything else must be expressed through a token that the dark block re-maps; a fixed hex outside the shell and the map is a defect.

## Typography

**Display Font:** system sans (`-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif`), used for everything.
**Body Font:** the same stack.
**Label/Mono Font:** `ui-monospace, "SF Mono", Menlo, Consolas, monospace` at 12.5 px, for inline code and the paste / SQL textareas only.

**Character:** the platform's own sans, antialiased, with `font-variant-numeric: tabular-nums` set on the body so every column of figures lines up. Headings and figures tighten to -0.01em and sit at 600; nothing goes heavier than 700 (the brand mark and map pins). There are six sizes and no fluid type.

### Hierarchy
- **Figure** (600, 24px / 1.15, -0.01em): the number in a stat tile (`.rp-tile .v`, `.fn-step .v`); drops to 20 px under 480 px. The calls strip's figure (`.cl-tiles span b`) is the 20-px step.
- **Page title** (600, 20px / 40px, -0.01em): the one `h1` on each page, and the help dialog's `h2`.
- **Card title** (600, 16px, -0.01em): `.card-title` / `.tr-title`, the empty-state heading, the drop zone's lead line, `h3` inside the help body.
- **Body** (400, 14px / 1.45): copy, inputs, selects, the month stepper's current month (600), the sidebar items (500), help prose (14 / 1.55).
- **Table / control** (13px): tables, buttons (500, line-height 1), section chips (500), hints, page sub-lines, messages, tile labels (500, muted), the chart tooltip.
- **Label** (500, 12px): form labels (muted), tags, as-of dates, tile footnotes, chart axis text, legends, the sidebar brand line.
- **Column head** (600, 12px, +0.05em, uppercase): every `th`, the Log's band row (same spacing), and the sidebar group labels at +0.06em.
- **Mono** (400, 12.5px / 1.5): `code` and `textarea`.

### Named Rules
**The Six Sizes Rule.** Type is set at 12, 13, 14, 16, 20 or 24 px (plus 12.5 px mono). A new element takes the nearest step; it does not add a seventh.

**The Tabular Figures Rule.** Numbers are right-aligned (`td.num`, `th.num`) and tabular everywhere, including tooltips and tiles, so differences between rows are visible without reading.

## Layout

The shell is a two-column CSS grid on `body`: a 232-px sidebar and a `minmax(0, 1fr)` main column, both `100vh` and scrolling independently (`overflow-y: auto`, `overscroll-behavior: contain`). The sidebar is padded 16 / 8 / 8; the main column is padded 24 px on the sides and top with 48 px at the bottom. The Log page is a flex column so its table card fills the viewport and scrolls inside its own box (`.log-wrap`, minimum 220 px) with headers, office column and total row pinned.

Vertical rhythm is an 8-px grid with 4 and 12 as half steps: card padding 16, card gap 16, page-header margin 16, nav items 8 / 12, nav group labels 16 / 12 / 4, stat-tile and table-card toolbar padding 12 / 16, table cells 8 / 12, Log cells 6 / 8, popover padding 10, section-chip row 8 / 16, card foot 8 / 16. Inline gaps inside a control or a toolbar are 6 and 10 px (icon gap 6; toolbar and row gaps 10 × 12). The toolbar's bottom margin (14 px) is the one off-grid value, an optical gap between a 32-px control row and the card below it.

Every page opens with a page header: `h1` on a 40-px line at left, the action group (`.page-actions`, 8-px gaps) at right, and a 13-px muted sub-line on its own row beneath. The table card then owns a toolbar (office search, month stepper), an optional section-chip row and an optional benchmark panel, each separated by a hairline.

Controls are 32 px tall; compact ones (example chips, the data-check "mark ok" form, map office rows) are 28 px. Report grids use `auto-fit` columns: tiles at `minmax(12rem, 1fr)`, funnel steps at `10rem`, paired cards at `24rem`, office cards at `23rem`, small charts at `10rem`.

Breakpoints:
- **≤ 900 px**: the grid collapses to a block flow. The sidebar becomes a sticky top bar (`grid-template-areas: "brand help" / "note note" / "nav nav"`); groups and icons hide, items become 13-px tabs with a 2-px bottom rule (white when current) in a horizontally scrolling row with a navy fade at the right edge. Main padding drops to 16 px, the `h1` line to 32 px, page actions and section chips scroll horizontally, the Log box caps at `calc(100vh - 220px)`.
- **≤ 760 px**: the calls page's two-column top (`.cl-top`) becomes one column.
- **≤ 640 px**: the map legend narrows to 170 px and the inset map shrinks.
- **≤ 480 px**: stat strips go to two columns, figures drop to 20 px, page-action buttons to 30 px.

## Elevation & Depth

Depth is a hairline or a shadow, never both, and never on a resting surface. Cards, inputs and the month stepper are flat: a 1-px `border` (or `border-strong` on controls) on white. The sidebar is a flat navy plane. Only surfaces that float above the page carry the single shadow token: the popover menus (`.cal`, `.colmenu`, `.mgrid`), the detail tooltip, the chart tooltip, the map legend and the help dialog, each with no border. In dark mode the shadow token adds a 1-px ring of `border` so the floating edge still reads against a dark canvas. The help dialog's backdrop is `rgba(10, 14, 24, .5)`.

`box-shadow` is also used as a drawing tool for hairlines where a border would misalign: the stat strip's cell dividers (`1px 0 0 0 border, 0 1px 0 0 border`), the sticky table headers (`inset 0 -1px 0 border`), the pinned Log column (`1px 0 0 border`) and the sticky total row (`inset 0 1px 0 border-strong`). These are lines, not depth.

### Shadow Vocabulary
- **Floating** (`box-shadow: 0 12px 32px -12px rgba(16, 24, 40, .30), 0 2px 6px rgba(16, 24, 40, .08)`; dark: `0 12px 32px -12px rgba(0, 0, 0, .7), 0 0 0 1px var(--border)`): popovers, tooltips, the map legend, the help dialog.
- **Focus ring** (`box-shadow: 0 0 0 3px color-mix(in srgb, var(--accent) 26%, transparent)` with the border turned `accent`): inputs, selects and textareas on `:focus-visible`. Everything else focuses with a 2-px `accent` outline offset 2 px (inset -2 px inside the sidebar, tables and the stepper; white inside the sidebar).
- **Map only**: pins (`0 1px 4px rgba(0,0,0,.5)`) and the inset map (`0 2px 8px rgba(0,0,0,.2)`) carry fixed shadows on the light map canvas; they do not leave the map.

### Named Rules
**The Hairline-or-Shadow Rule.** A surface has a 1-px border or the floating shadow, never both and never a second shadow value. Cards and controls get the line; menus, tooltips and dialogs get the shadow.

**The Flat Plane Rule.** The sidebar, the page canvas and the stat strips are flat. Depth is reserved for things that are temporarily above the page.

## Shapes

Corners are small and consistent: 6 px on every control (buttons, inputs, selects, nav items, menu rows, month-grid cells; the month stepper's inner buttons take 5 px to sit inside the 6-px frame), 10 px on cards, the popover menus, the drop zone and the help dialog, 8 px on tooltips, the map canvas, the confirm box, the help-report boxes, the brand mark and the dialog's close button, and 4 px on inline code, the skeleton bar and small key swatches. Chips, tags, the check count, the help steps and the calls "sig" pills are full pills (999 px). Stat tiles have no radius; they are cells of one card divided by hairlines, and the strip clips them with `overflow: hidden`. Rounded dots are circles: the section chip's 9-px track, the benchmark dot, map pins.

Borders are 1 px everywhere, with three exceptions that are not decoration: the drop zone's 1.5-px dashed `border-strong`, the section chip's 1.5-px track, and the Log band header's 2-px top rule in its section colour (also the 2-px bottom rule on mobile tabs). Nothing is clipped diagonally, skewed or outlined twice.

### Named Rules
**The Four Radii Rule.** 4 / 6 / 8 / 10 px and the pill. A new container takes 10, a new control 6, a new floating note 8; tiles inside a strip take none.

## Components

### Sidebar navigation
A flat navy plane (`side`) that says where you are. The brand block is a 30-px accent mark ("ES", 13 px, 700, 8-px radius) beside the product name (14 px, 600, white) and the group name (12 px, `side-muted`). Groups (Work / Reports / Admin · this computer) are 12-px uppercase labels at +0.06em in `side-muted`; the Admin group only renders when an admin tab is visible. Items are full-width 14-px / 500 buttons, 8 / 12 padding, 6-px radius, `side-text`, a 16-px inline SVG at 80% opacity; hover fills `side-2` and turns the text white; the current item fills with a 32% accent tint over the navy (`side-current`), turns white and 600, and its icon goes `#b7c4ff` at full opacity. Focus is a 2-px white outline inset. A viewer-mode note sits under the brand as a 12-px `side-2` pill-box. The help button lives in the foot (`margin-top: auto`). Under 900 px the whole thing is a sticky top bar with a scrolling row of underlined tabs (see Layout).

### Page header
- **Title row:** `h1` at 20 px on a 40-px line; actions to the right in an 8-px-gapped wrapping row; the 13-px muted sub-line wraps to its own full-width row beneath.
- **Actions:** quiet buttons for secondary actions (Copy, Download CSV), default buttons for pickers (Columns, Checks, Benchmarks, Office order), a danger button for Delete selected. A picker's button shows `button-expanded` while its menu is open.
- **Mobile:** the title line shrinks to 32 px, the sub-line moves above the actions, and the actions scroll horizontally without a scrollbar.

### Toolbar
A wrapping flex row (gap 10 × 12, `align-items: end`) of labelled controls. Inside a table card it is padded 12 / 16 and separated by a hairline. Labels are 12-px / 500 muted text stacked 4 px above their control.

### Cards
- **Corner Style:** 10 px.
- **Background:** `surface`; the benchmark panel inside a table card uses `surface-2`.
- **Shadow Strategy:** none; a 1-px `border` hairline (see Elevation & Depth).
- **Border:** 1 px `border`.
- **Internal Padding:** 16 px; 16-px gap between cards. A table card has zero padding and composes bands (toolbar, chip row, panel, table, foot) each separated by a hairline, the table's bottom corners following the card's radius.
- **Title:** 16 px / 600 with a 4-px gap to a 13-px muted sub-line (`.tr-sub` then 12 px below). Cards never nest.

### Buttons
- **Shape:** 32 px tall, 6-px radius, 13-px / 500 text, 0 / 12 padding, 6-px icon gap; compact variant 28 px with 0 / 10 padding; 30 px in the mobile page-actions row.
- **Default (`.btn`):** `surface` fill, 1-px `border-strong`, `text`; hover `hover` fill with a `faint` border; active `border` fill; disabled at 45% opacity.
- **Primary (`.btn.primary`):** `accent` fill, `accent-text`, 600, transparent border; hover `accent-hover`.
- **Danger (`.btn.danger`):** default button with `bad` text; hover `bad-soft` fill and a 40% `bad` border.
- **Quiet (`.btn.quiet`):** no border, no fill, `accent-ink` text; hover `accent-soft`.
- **Expanded (`aria-expanded="true"`):** `accent-soft` fill, 45% accent border, `accent-ink` text, for a button whose popover is open.
- **Link button (`.link-btn`):** 13-px / 500 `accent-ink` text with no box; underline on hover.
- **Transitions:** background, border-color and color at 150 ms.

### Inputs / Fields
- **Style:** 32 px tall, 14-px text, `surface` fill, 1-px `border-strong`, 6-px radius, 0 / 10 padding; placeholders in `faint`; selects hide the native arrow and draw a 12-px chevron (fixed `#6b7280` stroke) 9 px from the right edge, max-width 280 px; search fields are at least 180 px wide.
- **Hover:** border turns `faint`.
- **Focus:** border turns `accent` and the 3-px 26% accent ring appears; no outline.
- **Error:** `select.needs` takes a `bad` border and an 18% `bad` ring.
- **Textarea:** full width, 120-px minimum, 10-px padding, mono 12.5 / 1.5 on `code-bg`, vertical resize.
- **Label:** a column with a 4-px gap, 12 px / 500 `muted`.

### Month stepper
An inline group 32 px tall in a `border-strong` frame with a 6-px radius: two 30-px step buttons (14-px inline chevrons, `muted`, hover `hover` fill, 30% opacity when disabled) flanking the current month (14 px / 600, 0 / 12 padding, 150-px minimum, hairline left and right). The current-month button opens the month grid.

### Popovers (`.cal`, `.colmenu`, `.mgrid`)
Absolute, 6 px below their trigger, `surface`, 10-px radius, 10-px padding, the floating shadow and no border. The calendar is 260 px wide with a 600 header, 30-px year buttons and a 12-px muted foot of link-style buttons (hover `accent-soft`). The column / office menu is 300 px, right-aligned, capped at 70vh, a list of 13-px rows with a grip, a checkbox label, a 12-px section tag in its band colour and 26-px move buttons; a dragging row dims to 40%, drop targets draw a 1-px accent edge. The month grid is 250 px, centred, four 34-px cells per row (13 px / 500, hover `hover`, selected `accent` / `accent-text`), with a full-width "all months" button beneath.

### Section chips (`.switch`)
26-px pills with a 1-px `border`, `surface` fill, 13-px / 500 `muted` text, 0 / 10 / 0 / 8 padding and a 9-px ring track in the section colour. Checked: the track fills, the chip tints 9% in its section colour with a 40% section border and the text turns `text`; hover deepens the border to 55%. The 12-px count beside the name is `faint`. Role `switch` with `aria-checked`. Under 900 px the row scrolls horizontally.

### Tags and counts
`.tag` is a 12-px / 500 pill with a hairline and `muted` text; its warn variant takes `warn` text, a 45% warn border and `warn-soft` fill. `.chk-count` is an 18-px `chk` pill with white 12-px / 600 text.

### Tables
- **Style:** 13 px, `border-collapse`, cells 8 / 12 with a hairline beneath, `white-space: nowrap`, vertically centred; headers 12 px uppercase +0.05em 600 `muted`; numeric cells and headers right-aligned; the total row 600 with a `border-strong` rule above; rows highlight `hover` on hover; sticky headers in every scrolling box with an inset hairline.
- **Sorting:** header buttons inherit the header style, show a 12-px sort icon at 40% that becomes a full-strength up / down arrow when active, and turn `text` on hover or when sorted.
- **Detail values:** a count that carries a list is underlined dotted in `faint` (`.has-detail`) and opens the detail tooltip.
- **Log:** 6 / 8 cells; headers wrap to three lines (64–150 px); the band row above the headers is 12-px uppercase in the section colour with a 2-px top rule and a 7% tint; the office column (second) is pinned left on `surface` with a 1-px right hairline and 500 weight; the total row is pinned at the bottom; source cells truncate at 130 px in `muted`.
- **Benchmark cells:** `bm-*` fill and ink with an 8-px mark before the figure (dot / up triangle / down triangle); the as-of date inherits the ink at 75%; row hover darkens them 4%.
- **Data-check flag:** a 14-px `chk` warning triangle before the value, `cursor: help`, with an sr-only "(looks wrong)".
- **Calls sheet:** 13 px, 5 / 8 cells, even rows tinted `surface-2` at 70%, ok values `good` 600, booked values `accent-ink` 600, missing values `bad`.

### Stat strips (`.rp-tiles`, `.fn-steps`, `.cl-tiles`)
One card divided by hairlines: a `border` rule above the strip and each cell drawing its own right and bottom hairline via `box-shadow`; no radius, no shadow, no gap. Cells are padded 12 / 16 on `surface`: a 13-px / 500 `muted` label, a 24-px / 600 figure 4 px below, a 12-px `faint` footnote 4 px under that. The calls strip inverts the order (figure 20 px / 600 over a 12-px muted label) in a wrapping flex row. Under 480 px strips are two columns and figures are 20 px. While loading, tiles render skeleton blocks (label 50% × 11 px, figure 70% × 22 px, note 55% × 10 px).

### Skeleton and empty state
`.skel` is an inline block 0.85em tall (62% wide; 48% in numeric cells), 4-px radius, `hover` fill, pulsing to 40% opacity on a 1.3-s ease-in-out loop. `.empty` centres a 16-px / 600 heading over 13-px muted copy with 44 / 16 padding, and explains what to do next.

### Tooltips
- **Detail tooltip (`#detail-tip`):** fixed, up to `min(640px, 100vw - 24px)` wide and 60vh tall, `surface`, 8-px radius, floating shadow, 10 / 12 padding, 13 px / 1.5, `pre-line`; a 600 head with a 12-px muted right-hand span and a 12-px `accent-ink` copy button; an inner 12-px table with a sticky header.
- **Chart tooltip (`.tr-tip`):** absolute over the chart, 8-px radius, 8 / 10 padding, 13 px, floating shadow, 150-px minimum; a 12-px muted head, then rows of a 12 × 2 px series swatch, a 600 value and a muted span.

### Help dialog
A native `<dialog>` at `min(760px, 100vw - 32px)` by `min(88vh, 900px)`, `surface`, 10-px radius, floating shadow, dimmed backdrop. Head: 20-px `h2` with a 34-px icon close button (hover `hover`). A row of 13-px pill steps (selected `accent` / `accent-text`) under a hairline; a scrolling body of 14 / 1.55 prose with 16-px `h3`, 14-px `h4`, bordered 8-px "help-report" boxes and a `.path` breadcrumb of `surface-2` chips separated by a `faint` "›"; a hairline foot.

### Charts
Inline SVG: 12-px text in `muted` (values in `text`), `border` grid lines, `border-strong` axes, `series` bars with `series-dim` for secondary bars, 45%-opacity dashed outlines for partial months, a 2-px accent outline on focused groups. Small multiples sit in a `10rem` auto-fit grid with a 12-px title row. The ZIP table's inline bars are 8-px rounded (3 px) `series-dim` bars, `series` for the top row.

### Maps
Leaflet canvases 8-px rounded with a `border` hairline on the fixed light `map-canvas`, `min(70vh, 640px)` tall (340 px on the calls page). Pins are 22-px `map-pin` circles with a 2-px white ring and 11-px / 700 white numerals (16 px when zoomed out); city labels are 10–14 px / 500–700 `map-label` on a white text-shadow halo. The legend is a `surface` 8-px box with the floating shadow, 12-px text, a colour ramp strip and dot keys. Map text uses `system-ui, sans-serif` rather than the body stack, and its colours are fixed hex; this is the one place the scheme does not re-map.

### Icons
An authored 16 × 16 set, 1.5-px stroke with round caps and joins, `currentColor`, no fill: inline SVG for the sidebar items (16 px, 80% opacity, full on the current item), the help button and the stepper arrows (14 px, 1.8 stroke). The rest are `mask-image` custom properties on `:root` (`--i-up`, `--i-down`, `--i-left`, `--i-right`, `--i-sort`, `--i-x`, `--i-target`, `--i-warn` at 1.8 stroke; `--i-dot`, `--i-tri-up`, `--i-tri-down`, `--i-grip` filled), painted with `background: currentColor` through `mask` on `.ico`, `.arrow`, `.bm-flag`, `.bm-mark`, `.chk-flag`, `.grip`, the move buttons and the year buttons, rendered at 8–14 px. Where the script still writes a glyph into one of these elements the stylesheet zeroes its font-size so only the mask shows. The select chevron is a separate 12-px data-URI background.

### Motion grammar
State changes only, 150 ms (`background`, `border-color`, `color`, `box-shadow`; the browser's default curve) on nav items, buttons, inputs, chips and the drop zone. The Log swap: when the month or office filter changes, `.log-wrap.swap` drops the box to 40% opacity instantly and the class's removal fades it back over 180 ms ease-out. The skeleton pulses on a 1.3-s ease-in-out loop. `prefers-reduced-motion: reduce` removes every transition and stops the pulse. There are no entrance animations, no transforms on hover and no parallax.

## Do's and Don'ts

### Do:
- **Do** build a familiar working-tool dashboard in the Stripe / Looker canon: sidebar, page header with scope and actions, tiles, then tables.
- **Do** put the figure first: a 24-px tabular number over a 13-px muted label, right-aligned numbers in every table, source and as-of beside the value.
- **Do** use the same vocabulary on every tab: 32-px controls with a 6-px radius, 10-px cards with one hairline, 12-px uppercase column heads, pills for chips and counts.
- **Do** express colour through the tokens so the `prefers-color-scheme: dark` block re-maps it; only the sidebar and the map keep fixed values.
- **Do** pair every state colour with a word, a mark or a shape (the benchmark dot / triangles, the data-check triangle, "ok" / "missing" text).
- **Do** draw icons from the authored 16-px SVG set, inline or through the `--i-*` masks, in `currentColor`.
- **Do** keep motion to 150-ms state changes, the 180-ms Log swap and the skeleton pulse, and honour `prefers-reduced-motion`.
- **Do** give every scrolling table a sticky header, and the Log its pinned office column and total row.
- **Do** write empty states that explain what to do next, and show skeleton rows while data loads.

### Don't:
- **Don't** add anything flashy: no gradients beyond the flat 7% band tint and the mobile nav's scroll fade, no decorative shadows, no glows, no entrance choreography.
- **Don't** introduce a second accent; emphasis is weight, size or the one indigo.
- **Don't** nest a card inside a card; compose a table card from hairline-separated bands instead.
- **Don't** put a shadow on a resting surface or a border on a floating one.
- **Don't** add a seventh type size, a display face or a second sans; the system stack at 12 / 13 / 14 / 16 / 20 / 24 px is the whole ramp.
- **Don't** render an icon as a Unicode glyph or an icon font; the mask set exists for exactly that.
- **Don't** use `warn` for a data check or `chk` for a benchmark; the orange belongs to the checks alone.
- **Don't** surface a patient name or a phone number anywhere in the interface, including tooltips, exports previews and map popups.
- **Don't** darken the map canvas in dark mode; its colour ramp needs the light ground.
