"""Report builder output: one report (long rows plus how to pivot them) as a preview table, a CSV, or a formatted
.xlsx whose Summary and office tabs are formulas over a Data tab. A share copy (a file that may leave the group) is
only the Summary table, as plain values, and a Notes tab.

A report is a dict:
    title, subtitle, notes[]         text for the top of the Summary tab and the Notes tab
    row_label, extra_label           what a row is ("ZIP") and an optional second label column ("City")
    columns                          "office", "month" or None: what the Summary's columns are
    offices[], months[]              column order (short office names, "YYYY-MM")
    fields[]                         {key, label, kind}: the number columns on the Data tab
    measures[]                       {key, label, kind, num, den}: num/den are field keys for rates and averages
    data[]                           {office, month, group, extra, <field key>: number}
    order[]                          optional row order (else biggest first, by the first measure)
    office_tabs                      True to add one tab per office, by month
    last                             optional row label kept as the last row whatever its size ("Other")
    share, marks, mark               a share copy: marks = {(row label, column)} cells that show the text mark ("*")
                                     because their numbers are counted in the "last" row instead

The .xlsx is written as a flat ODF spreadsheet and converted by LibreOffice (soffice), which also calculates the
formulas: Python has no Excel writer here, and LibreOffice is already on the machine."""
import csv
import io
import re
import shutil
import subprocess
import tempfile
import threading
from datetime import date
from pathlib import Path
from xml.etree import ElementTree
from xml.sax.saxutils import escape
import zipfile


class ReportError(Exception):
    pass


MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
_convert_lock = threading.Lock()


def month_label(m, year=True):
    """'2026-03' -> 'Mar 2026' (or 'Mar')."""
    name = MONTH_NAMES[int(m[5:7]) - 1]
    return f"{name} {m[:4]}" if year else name


def safe_label(s, strict=False):
    """A row label that SUMIFS can match exactly: no wildcards, and no leading =, < or > (which would read as a test).
    strict (a share copy, whose CSV an outside reader opens): no leading +, - or @ either, which a spreadsheet reads as
    the start of a formula."""
    s = " ".join(str(s).split()).translate({42: "∗", 63: "？", 126: "∼"}) or "(not set)"
    swap = {"=": "＝", "<": "＜", ">": "＞", **({"+": "＋", "-": "－", "@": "＠"} if strict else {})}
    return swap[s[0]] + s[1:] if s[0] in swap else s


def _months_head(months):
    year = len({m[:4] for m in months}) > 1
    return [month_label(m, year) for m in months]


def _total_head(m):
    """The heading of the main measure's total column, beside its per-office or per-month columns."""
    label = m["label"]
    if m.get("num"):
        return f"{label} overall"
    return label if label.lower().startswith("total ") else "Total " + (label if label[1:2].isupper() else label[0].lower() + label[1:])


def _value(rows, m):
    """A measure over some data rows: added up, or top ÷ bottom for rates and averages (None when there's no bottom)."""
    if m.get("num"):
        den = sum(r.get(m["den"]) or 0 for r in rows)
        return sum(sum(r.get(k) or 0 for k in m["num"]) for r in rows) / den if den else None
    return sum(r.get(m["key"]) or 0 for r in rows)


def _groups(rep, data):
    """Row labels in display order, and each one's second label."""
    main = rep["measures"][0]
    by = {}
    for r in data:
        by.setdefault(r["group"], []).append(r)
    if rep.get("order"):
        rank = {g: i for i, g in enumerate(rep["order"])}
        names = sorted(by, key=lambda g: rank.get(g, len(rank)))
    else:
        names = sorted(by, key=lambda g: (-(_value(by[g], main) or 0), g))
    if rep.get("last") in by:   # "Other" is not one of the rows: it never ranks among them
        names.remove(rep["last"]); names.append(rep["last"])
    return names, by


def table(rep, limit=None):
    """The Summary tab as values, for the page's preview: {head, kinds, rows, total, n_rows}."""
    measures, main = rep["measures"], rep["measures"][0]
    names, by = _groups(rep, rep["data"])
    dim = rep.get("columns")
    cols = rep["offices"] if dim == "office" else rep["months"] if dim == "month" else []
    head = [rep["row_label"]] + ([rep["extra_label"]] if rep.get("extra_label") else [])
    kinds = ["text"] * len(head)
    head += cols if dim == "office" else _months_head(cols)
    kinds += [main["kind"]] * len(cols)
    # In a share copy with columns, a row's totals add up only the cells shown on that row: say so in the heading.
    shown = " (numbers shown)" if rep.get("share") and cols else ""
    head += [(_total_head(m) if cols and m is main else m["label"]) + shown for m in measures]
    kinds += [m["kind"] for m in measures]
    marks = rep.get("marks") or ()

    def line(rows, first, g=None):
        out = list(first)
        out += [rep["mark"] if (g, c) in marks else _value([r for r in rows if r[dim] == c], main) for c in cols]
        return out + [_value(rows, m) for m in measures]

    extra = {r["group"]: r.get("extra") or "" for r in rep["data"]}
    body = [line(by[g], [g] + ([extra[g]] if rep.get("extra_label") else []), g) for g in names]
    total = line(rep["data"], ["Total"] + ([""] if rep.get("extra_label") else []))
    small = set(rep.get("small") or [])
    if limit and len(body) > limit:   # a preview: the first rows, and always the "last" row ("Other") when there is one
        cut = limit - 1 if names[-1] == rep.get("last") else limit
        body, names = body[:cut] + body[len(body) - (limit - cut):], names[:cut] + names[len(names) - (limit - cut):]
    return {"head": head, "kinds": kinds, "rows": body, "total": total, "n_rows": len(by), "small": [g for g in names if g in small]}


def _stored(v, kind):
    """A share copy's number as it is kept in the file: no more exact than what is shown (whole dollars and counts)."""
    return round(v, 3) if kind == "pct" else int(round(v))


def to_csv(rep):
    """The Data tab as CSV: one line per office, month and row. A share copy is the Summary table instead, with
    nothing behind it."""
    out = io.StringIO()
    w = csv.writer(out)
    if rep.get("share"):
        t = table(rep)
        w.writerow(t["head"])
        for row in t["rows"] + [t["total"]]:
            w.writerow(["" if v is None else v if isinstance(v, str) else _stored(v, k) for v, k in zip(row, t["kinds"])])
        return out.getvalue().encode("utf-8-sig")
    has_extra = bool(rep.get("extra_label"))
    w.writerow(["Office", "Month", rep["row_label"]] + ([rep["extra_label"]] if has_extra else []) + [m["label"] for m in rep["measures"]])
    for r in rep["data"]:
        vals = [_value([r], m) for m in rep["measures"]]
        w.writerow([r["office"], r["month"], r["group"]] + ([r.get("extra") or ""] if has_extra else [])
                   + ["" if v is None else round(v, 4) for v in vals])
    return out.getvalue().encode("utf-8-sig")


# ------------------------------------------------------------------ the workbook

def _col(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _attr(s):
    return escape(str(s), {'"': "&quot;"})


NO_VALUE = object()   # a formula cell left for LibreOffice to calculate


def _cell(v=None, style="txt", formula=None, value=NO_VALUE):
    """A cell. A formula can carry its answer (value; None = empty text), which saves LibreOffice working it out."""
    st = f' table:style-name="{style}"'
    if formula:
        known = "" if value is NO_VALUE else ' office:value-type="string" office:string-value=""' if value is None \
            else f' office:value-type="float" office:value="{round(value, 6)}"'
        return f'<table:table-cell{st} table:formula="of:={_attr(formula)}"{known}/>'
    if v is None or v == "":
        return f"<table:table-cell{st}/>"
    if isinstance(v, (int, float)):
        return f'<table:table-cell{st} office:value-type="float" office:value="{v}"><text:p>{v}</text:p></table:table-cell>'
    return f'<table:table-cell{st} office:value-type="string"><text:p>{escape(str(v))}</text:p></table:table-cell>'


def _row(cells, style=None):
    return f"<table:table-row{f' table:style-name={chr(34)}{style}{chr(34)}' if style else ''}>" + "".join(cells) + "</table:table-row>"


def _sheet_name(name, taken):
    """A tab name Excel accepts: no []:*?/\\, at most 31 characters, not used yet."""
    base = "".join(" " if c in "[]:*?/\\'" else c for c in str(name)).strip()[:31] or "Sheet"
    out, i = base, 2
    while out.lower() in taken:
        out = f"{base[:28]} {i}"; i += 1
    taken.add(out.lower())
    return out


STYLE = {"money": "money", "count": "count", "pct": "pct"}
WIDTH = {"money": 15, "count": 12, "pct": 11}


def _document(tables, widths, panes):
    """The flat ODF file around its tabs: column widths, frozen panes {tab name: (columns, rows)} and the cell styles."""
    col_styles = "".join(f'<style:style style:name="w{w}" style:family="table-column"><style:table-column-properties '
                         f'style:column-width="{w * 0.19:.2f}cm"/></style:style>' for w in sorted(widths))
    item = lambda k, t, v: f'<config:config-item config:name="{k}" config:type="{t}">{v}</config:config-item>'
    frozen = "".join(
        f'<config:config-item-map-entry config:name="{_attr(name)}">'
        + item("HorizontalSplitMode", "short", 2 if x else 0) + item("VerticalSplitMode", "short", 2)
        + item("HorizontalSplitPosition", "int", x) + item("VerticalSplitPosition", "int", y)
        + item("PositionRight", "int", x) + item("PositionBottom", "int", y) + "</config:config-item-map-entry>"
        for name, (x, y) in panes.items())
    font = 'style:font-name="Arial" fo:font-family="Arial"'
    total_cell = '<style:table-cell-properties fo:background-color="#e8eef5" fo:border-top="0.5pt solid #b7c3d0"/>'
    total_text = f'<style:text-properties {font} fo:font-weight="bold"/>'
    total_box = total_cell + total_text
    number = '<number:number number:decimal-places="0" number:min-integer-digits="1" number:grouping="true"/>'
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
 xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
 xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2"
 xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" xmlns:ooo="http://openoffice.org/2004/office"
 xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
 office:version="1.2" office:mimetype="application/vnd.oasis.opendocument.spreadsheet">
<office:settings><config:config-item-set config:name="ooo:view-settings"><config:config-item-map-indexed config:name="Views">
 <config:config-item-map-entry><config:config-item-map-named config:name="Tables">{frozen}</config:config-item-map-named>
 </config:config-item-map-entry></config:config-item-map-indexed></config:config-item-set></office:settings>
<office:font-face-decls><style:font-face style:name="Arial" svg:font-family="Arial"/></office:font-face-decls>
<office:styles>
 <style:style style:name="Default" style:family="table-cell"><style:text-properties {font} fo:font-size="10pt"/></style:style>
</office:styles>
<office:automatic-styles>
 <number:currency-style style:name="cur0P" style:volatile="true"><number:currency-symbol>$</number:currency-symbol>{number}</number:currency-style>
 <number:currency-style style:name="cur0N" style:volatile="true"><number:text>-</number:text><number:currency-symbol>$</number:currency-symbol>{number}</number:currency-style>
 <number:currency-style style:name="cur0"><number:text>–</number:text><style:map style:condition="value()&gt;0" style:apply-style-name="cur0P"/><style:map style:condition="value()&lt;0" style:apply-style-name="cur0N"/></number:currency-style>
 <number:number-style style:name="int0">{number}</number:number-style>
 <number:percentage-style style:name="pct1"><number:number number:decimal-places="1" number:min-integer-digits="1"/><number:text>%</number:text></number:percentage-style>
 {col_styles}
 <style:style style:name="hrow" style:family="table-row"><style:table-row-properties style:row-height="0.9cm"/></style:style>
 <style:style style:name="txt" style:family="table-cell" style:parent-style-name="Default"/>
 <style:style style:name="wrap" style:family="table-cell" style:parent-style-name="Default"><style:table-cell-properties fo:wrap-option="wrap" style:vertical-align="top"/></style:style>
 <style:style style:name="title" style:family="table-cell" style:parent-style-name="Default"><style:text-properties {font} fo:font-size="13pt" fo:font-weight="bold"/></style:style>
 <style:style style:name="note" style:family="table-cell" style:parent-style-name="Default"><style:text-properties {font} fo:font-size="9pt" fo:font-style="italic" fo:color="#555555"/></style:style>
 <style:style style:name="head" style:family="table-cell" style:parent-style-name="Default"><style:table-cell-properties fo:background-color="#1f3a5f" fo:wrap-option="wrap" style:vertical-align="middle"/><style:paragraph-properties fo:text-align="center"/><style:text-properties {font} fo:font-weight="bold" fo:color="#ffffff"/></style:style>
 <style:style style:name="money" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="cur0"/>
 <style:style style:name="count" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="int0"/>
 <style:style style:name="pct" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="pct1"/>
 <style:style style:name="mark" style:family="table-cell" style:parent-style-name="Default"><style:paragraph-properties fo:text-align="end"/></style:style>
 <style:style style:name="tmark" style:family="table-cell" style:parent-style-name="Default">{total_cell}<style:paragraph-properties fo:text-align="end"/>{total_text}</style:style>
 <style:style style:name="totl" style:family="table-cell" style:parent-style-name="Default">{total_box}</style:style>
 <style:style style:name="tmoney" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="cur0">{total_box}</style:style>
 <style:style style:name="tcount" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="int0">{total_box}</style:style>
 <style:style style:name="tpct" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="pct1">{total_box}</style:style>
</office:automatic-styles>
<office:body><office:spreadsheet>
 <table:calculation-settings table:use-regular-expressions="false" table:use-wildcards="true" table:search-criteria-must-apply-to-whole-cell="true"/>
 {"".join(tables)}</office:spreadsheet></office:body>
</office:document>"""


def _fods(rep, cached=True):
    """The workbook as flat ODF. cached=False leaves every formula for LibreOffice to calculate (slow; used to check
    that the formulas give the same numbers as _value)."""
    measures, main = rep["measures"], rep["measures"][0]
    data, has_extra = rep["data"], bool(rep.get("extra_label"))
    names, by = _groups(rep, data)
    extra = {r["group"]: r.get("extra") or "" for r in data}
    n_last = len(data) + 1
    # Data tab columns: A Office, B Month, C the row label, (D the second label), then the number fields.
    first_field = 5 if has_extra else 4
    field_col = {f["key"]: _col(first_field + i) for i, f in enumerate(rep["fields"])}
    rng = lambda letter: f"[$Data.${letter}$2:.${letter}${n_last}]"
    quote = lambda s: '"' + str(s).replace('"', '""') + '"'

    def agg(m, crit):
        """A measure as a formula over the Data tab, for the rows meeting crit [(Data column letter, expression)]."""
        def total(key):
            if not crit:
                return f"SUM({rng(field_col[key])})"
            return f"SUMIFS({rng(field_col[key])};" + ";".join(f"{rng(c)};{e}" for c, e in crit) + ")"
        if m.get("num"):
            return f'IFERROR(({"+".join(total(k) for k in m["num"])})/{total(m["den"])};"")'
        return total(m["key"])

    tables, widths, panes, taken = [], set(), {}, {"data", "notes"}

    def add(name, w, rows, freeze=None):
        widths.update(w)
        cols = "".join(f'<table:table-column table:style-name="w{x}"/>' for x in w)
        tables.append(f'<table:table table:name="{_attr(name)}">{cols}{"".join(rows)}</table:table>')
        if freeze:
            panes[name] = freeze

    def pivot(name, title, sub, groups, dim, cols, col_heads, crit0, share, rows_in):
        """A tab: one row per group, a column of the main measure per office or month, then each measure's total.
        rows_in: the Data rows this tab adds up (what crit0 selects), for the formulas' answers."""
        mine, cut = {}, {}
        for r in rows_in:
            mine.setdefault(r["group"], []).append(r)
            if dim:
                cut.setdefault((r["group"], r[dim]), []).append(r); cut.setdefault((None, r[dim]), []).append(r)
        val = (lambda rows, m: _value(rows, m)) if cached else (lambda rows, m: NO_VALUE)
        grand = _value(rows_in, main) if cached else None
        label_cols = 2 if has_extra else 1
        head = [rep["row_label"]] + ([rep["extra_label"]] if has_extra else []) + col_heads
        head += [_total_head(m) if cols and m is main else m["label"] for m in measures]
        if share:
            head.append("Share of total")
        rows = [_row([_cell(title, "title")]), _row([_cell(sub, "note")]), _row([]), _row([_cell(h, "head") for h in head], "hrow")]
        first = 5
        t = first + len(groups)
        main_total = _col(label_cols + len(cols) + 1)
        data_dim = {"office": "A", "month": "B"}.get(dim)
        for i, g in enumerate(groups, first):
            here = crit0 + [("C", f"[.$A{i}]")]
            cs = [_cell(g)] + ([_cell(extra.get(g, ""))] if has_extra else [])
            cs += [_cell(style=STYLE[main["kind"]], formula=agg(main, here + [(data_dim, quote(c))]), value=val(cut.get((g, c), []), main))
                   for c in cols]
            cs += [_cell(style=STYLE[m["kind"]], formula=agg(m, here), value=val(mine.get(g, []), m)) for m in measures]
            if share:
                cs.append(_cell(style="pct", formula=f'IFERROR([.{main_total}{i}]/[.${main_total}${t}];"")',
                                value=(_value(mine.get(g, []), main) / grand if grand else None) if cached else NO_VALUE))
            rows.append(_row(cs))
        tc = [_cell("Total", "totl")] + ([_cell("", "totl")] if has_extra else [])
        tc += [_cell(style="t" + STYLE[main["kind"]], formula=agg(main, crit0 + [(data_dim, quote(c))]), value=val(cut.get((None, c), []), main))
               for c in cols]
        tc += [_cell(style="t" + STYLE[m["kind"]], formula=agg(m, crit0), value=val(rows_in, m)) for m in measures]
        if share:
            tc.append(_cell("", "totl"))
        rows.append(_row(tc))
        w = [max(9, min(34, max((len(str(g)) for g in groups), default=9) + 2))] + ([24] if has_extra else [])
        w += [max(WIDTH[main["kind"]] - 2, 11)] * len(cols) + [WIDTH[m["kind"]] + 2 for m in measures] + ([13] if share else [])
        add(name, w, rows, (label_cols, 4))

    dim = rep.get("columns")
    cols = rep["offices"] if dim == "office" else rep["months"] if dim == "month" else []
    heads = cols if dim == "office" else _months_head(cols)
    taken.add("summary")
    pivot("Summary", rep["title"], rep.get("subtitle") or "", names, dim, cols, heads, [], False, data)

    if rep.get("office_tabs") and len(rep["offices"]) > 1:
        for office in rep["offices"]:
            mine = [r for r in data if r["office"] == office]
            if not mine:
                continue
            groups, _ = _groups(rep, mine)
            pivot(_sheet_name(office, taken), f"{office}: {rep['title']}", "By month. Each row adds up the Data tab for this office.",
                  groups, "month", rep["months"], _months_head(rep["months"]), [("A", quote(office))], not main.get("num"), mine)

    head = ["Office", "Month", rep["row_label"]] + ([rep["extra_label"]] if has_extra else []) + [f["label"] for f in rep["fields"]]
    rows = [_row([_cell(h, "head") for h in head], "hrow")]
    for r in data:
        cs = [_cell(r["office"]), _cell(r["month"]), _cell(r["group"])] + ([_cell(r.get("extra") or "")] if has_extra else [])
        cs += [_cell(r.get(f["key"]), STYLE[f["kind"]]) for f in rep["fields"]]
        rows.append(_row(cs))
    add("Data", [16, 10, 22] + ([24] if has_extra else []) + [WIDTH[f["kind"]] + 2 for f in rep["fields"]], rows, (0, 1))

    notes = list(rep.get("notes") or []) + [
        "The Data tab holds the numbers, one row per office, month and " + rep["row_label"].lower() + ", with no patient details. "
        "The other tabs add them up with formulas, so a corrected Data row updates every tab.",
        f"Made {date.today().isoformat()} with the Executive Summary Logger's report builder."]
    add("Notes", [110], [_row([_cell("About this report", "title")])] + [_row([_cell(n, "wrap")]) for n in notes])

    return _document(tables, widths, panes)


def _fods_share(rep):
    """A share copy as flat ODF: the Summary table and a Notes tab, as plain text and numbers. No Data tab, no office
    tabs and no formulas, so the file holds nothing the table doesn't show. Returns the file and every piece of text
    written to it, for _scrub_share to check the finished workbook against."""
    t, texts = table(rep), []

    def text(v, style):
        texts.append(str(v))
        return _cell(v, style)

    def cell(v, kind, total=False):
        if kind == "text":
            return text(v, "totl" if total else "txt") if v else _cell("", "totl" if total else "txt")
        if isinstance(v, str):   # the mark for a number that isn't shown
            return text(v, "tmark" if total else "mark")
        if v is None:
            return _cell("", "totl" if total else "txt")
        v = _stored(v, kind)
        if kind == "count" and 0 < v < rep["small_limit"]:   # can't happen: every count here is a patient count already tested
            raise ReportError("This share copy would have shown a number that is too small, so it wasn't made.")
        return _cell(v, ("t" if total else "") + STYLE[kind])

    rows = [_row([text(rep["title"], "title")]), _row([text(rep.get("subtitle") or "", "note")]), _row([]),
            _row([text(h, "head") for h in t["head"]], "hrow")]
    rows += [_row([cell(v, k) for v, k in zip(r, t["kinds"])]) for r in t["rows"]]
    rows.append(_row([cell(v, k, True) for v, k in zip(t["total"], t["kinds"])]))
    labels = sum(1 for k in t["kinds"] if k == "text")
    widths = [max(9, min(44, max((len(str(r[0])) for r in t["rows"]), default=9) + 2))] + [24] * (labels - 1)
    widths += [WIDTH[k] + 2 for k in t["kinds"][labels:]]
    notes = [_row([text("About this file", "title")])] + [_row([text(n, "wrap")]) for n in rep.get("notes") or []]
    tabs = [("Summary", widths, rows), ("Notes", [110], notes)]
    tables = [f'<table:table table:name="{name}">' + "".join(f'<table:table-column table:style-name="w{x}"/>' for x in w)
              + "".join(r) + "</table:table>" for name, w, r in tabs]
    return _document(tables, set(widths) | {110}, {"Summary": (labels, 4)}), texts


SHARE_PARTS = re.compile(r"^(\[Content_Types\]\.xml|_rels/\.rels|docProps/(core|app)\.xml|xl/workbook\.xml|xl/_rels/workbook\.xml\.rels|"
                         r"xl/styles\.xml|xl/theme/theme1\.xml|xl/sharedStrings\.xml|xl/worksheets/sheet[12]\.xml)$")
# What replaces LibreOffice's own document properties (its exact build and the server's system) in a share copy.
SHARE_APP = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties '
             'xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>LibreOffice</Application></Properties>')
SHARE_CORE = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties '
              'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" '
              'xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"/>')


def _scrub_share(raw, texts):
    """Check a share copy's finished .xlsx and take LibreOffice's own details out of it. The file must be exactly the
    two tabs that were written, as plain values. A formula, another part, a hidden tab or any text that wasn't written
    means something went wrong, and then no file is given out."""
    def refuse(why):
        raise ReportError(f"The share copy didn't pass its last check ({why}), so no file was made.")
    src, out = zipfile.ZipFile(io.BytesIO(raw)), io.BytesIO()
    names, want = src.namelist(), {" ".join(s.split()) for s in texts}
    for part in ("xl/workbook.xml", "xl/sharedStrings.xml", "xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"):
        if part not in names:
            refuse(f"{part} is missing")
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
        for name in names:
            if not SHARE_PARTS.match(name):
                refuse(f"unexpected part {name}")
            data = src.read(name)
            if name == "docProps/app.xml":
                data = SHARE_APP.encode()
            elif name == "docProps/core.xml":
                data = SHARE_CORE.encode()
            elif name.startswith("xl/worksheets/") and re.search(rb"<f[ >/]", data):
                refuse("a formula")
            elif name == "xl/workbook.xml":
                xml = data.decode("utf-8")
                if re.findall(r'<sheet [^>]*?name="([^"]*)"', xml) != ["Summary", "Notes"] or "idden" in xml \
                        or "definedName" in xml or "externalReference" in xml:
                    refuse("its tabs")
            elif name == "xl/sharedStrings.xml":
                for si in ElementTree.fromstring(data):
                    if " ".join("".join(si.itertext()).split()) not in want | {""}:
                        refuse("text that wasn't written")
            info = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))   # no clock time either: the Notes tab gives the day
            info.compress_type = zipfile.ZIP_DEFLATED
            dst.writestr(info, data)
    return out.getvalue()


def to_xlsx(rep, profile_dir, cached=True):
    """The report as .xlsx bytes. profile_dir: a folder LibreOffice may keep its own settings in (so a copy the user
    has open isn't disturbed)."""
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if not exe:
        raise ReportError("Excel files need LibreOffice on the computer that runs the app, and it isn't installed. Download the CSV instead.")
    if not rep["data"]:
        raise ReportError("Nothing to download: no data for these choices.")
    Path(profile_dir).mkdir(parents=True, exist_ok=True)
    xml, texts = _fods_share(rep) if rep.get("share") else (_fods(rep, cached), None)
    with tempfile.TemporaryDirectory() as tmp, _convert_lock:
        src = Path(tmp) / "report.fods"
        src.write_text(xml, encoding="utf-8")
        try:
            done = subprocess.run([exe, f"-env:UserInstallation={Path(profile_dir).resolve().as_uri()}", "--headless", "--calc",
                                   "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", tmp, str(src)],
                                  capture_output=True, text=True, timeout=180)
        except subprocess.TimeoutExpired:
            raise ReportError("The Excel file took too long to make. Pick fewer offices or months, or download the CSV.")
        out = Path(tmp) / "report.xlsx"
        if not out.exists():
            raise ReportError(f"LibreOffice couldn't make the Excel file ({(done.stderr or done.stdout or 'no message').strip()[-200:]}). "
                              "Download the CSV instead.")
        raw = out.read_bytes()
    return raw if texts is None else _scrub_share(raw, texts)
