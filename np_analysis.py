"""Read an office's "New Patient Analysis" workbook (.xlsm/.xlsx).

The workbook has one sheet per month. Each month sheet is a Denticon Referral
Production Listing pasted in (often several pastes as the month goes on), with the
team's own columns added: Insurance Type, financing, dentist, etc.

Only counts are returned; patient rows are read to count them and never kept.
Macros in .xlsm files are never run -- only cell values are read.
"""

import re
import zipfile
import xml.etree.ElementTree as ET
from datetime import date, timedelta

M = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
MONTHS = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]
SKIP_SHEETS = re.compile(r"master|original|do not use", re.I)
NUMBER_RE = re.compile(r"-?\d+(\.\d+)?(E[-+]?\d+)?$")

# Insurance Type values counted for each metric (compared case-insensitively, ignoring spaces/dashes).
INSURANCE_METRICS = {
    "medi_cal_patients": {"medical", "medi-cal", "medi cal", "medcal"},
    "hmo_patients": {"hmo"},
}


class NPAnalysisError(Exception):
    pass


class NoOfficeError(NPAnalysisError):
    pass


# "Office: Dentist Of Gardena" (Denticon's header) or "Total for Office,Gardena Dental Care [ 3 ] :"
OFFICE_LINE_RE = re.compile(r"(?:Office\s*:|Total for Office\s*,)\s*(.+?)(?:\s*\[.*)?$")


def find_office(wb):
    """The office name from the newest month tab that has an "Office: ..." line. Some tabs are pasted
    without Denticon's header, so a tab may have no office line of its own."""
    for month in dated_months(wb):
        for cells in wb.rows(month["sheet"]).values():
            m = OFFICE_LINE_RE.match(cells.get("A", ""))
            if m:
                return clean_office(m.group(1))
    return None


class EmptyMonthError(NPAnalysisError):
    """A month tab with no patients yet (e.g. next month's tab set up early)."""


def _norm(text):
    return re.sub(r"[\s\-]+", "", (text or "").lower())


INSURANCE_KEYS = {m: {_norm(v) for v in vals} for m, vals in INSURANCE_METRICS.items()}
NOT_RAN = {"", "no", "none", "n/a", "na", "-"}   # "Subprime Co. Ran" values meaning nothing was run

# Notes saying the patient already has a financing account: "pt had an existing care credit account",
# "existing alphaeon account", "pre existing care credit", "already has cherry". Not "existing pt",
# "existing braces", or "ran on ortho account".
LENDERS = {
    "Care Credit": r"care\s*credit", "Alphaeon": r"alphaeon", "Cherry": r"cherry", "Sunbit": r"sunbit",
    "Patient FI": r"patient\s*fi\b|\bpfi\b", "Proceed": r"proceed", "HFD": r"\bhfd\b",
    "Covered Care": r"covered\s*care", "Power Pay": r"power\s*pay", "Clarity Pay": r"clarity\s*pay",
    "Access": r"\baccess\b", "Financing": r"financ\w*",
}
_LENDER = "|".join(f"(?:{p})" for p in LENDERS.values())
EXISTING_FIN_RE = re.compile(
    rf"\b(?:pre[\s-]?)?existing\s+(?:[a-z]+\s+){{0,2}}?(?:{_LENDER})"
    rf"|\balready\s+(?:has|had|have)\s+(?:an?\s+)?(?:{_LENDER})", re.I)


def lender_named(text):
    for name, pat in LENDERS.items():
        if re.search(pat, text, re.I):
            return name
    return "Financing"

# Notes saying the patient has no SSN / no ITIN, in the many ways staff write it:
# "no itin or social", "no social and no itin", "no ss#", "patient has no itin or ssn",
# "no ssn only itin", "itin only". A plain "no id" does not count.
_ID = r"(?:s\.?\s*s\.?\s*n\b|ss\s*#|social(?:\s+security)?|itin\b)"
# The trailing part takes in the rest of the phrase ("no social and no itin") so the hover list
# shows what staff wrote; it doesn't change which patients match.
_MORE_IDS = rf"(?:\s*(?:and|or|,|/|&)\s*(?:no\s+)?{_ID})*"
NO_SSN_ITIN_RE = re.compile(rf"\bno\s+(?:(?:valid|an?)\s+)?{_ID}{_MORE_IDS}|\bitin\s+only\b|\bonly\s+itin\b", re.I)


class Workbook:
    def __init__(self, path):
        try:
            self.z = zipfile.ZipFile(path)
        except zipfile.BadZipFile:
            raise NPAnalysisError("Not a valid Excel workbook")
        names = self.z.namelist()
        self.shared = []
        if "xl/sharedStrings.xml" in names:
            for si in ET.fromstring(self.z.read("xl/sharedStrings.xml")).findall(f"{M}si"):
                self.shared.append("".join(t.text or "" for t in si.iter(f"{M}t")))
        wb = ET.fromstring(self.z.read("xl/workbook.xml"))
        rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(self.z.read("xl/_rels/workbook.xml.rels"))}
        self.sheets = {}
        for s in wb.find(f"{M}sheets"):
            t = rels[s.get(RID)].lstrip("/")
            self.sheets[s.get("name")] = t if t.startswith("xl/") else "xl/" + t

    def rows(self, sheet):
        """{row number: {column letters: text}} for one sheet."""
        out = {}
        for c in ET.fromstring(self.z.read(self.sheets[sheet])).iter(f"{M}c"):
            kind, v = c.get("t"), c.find(f"{M}v")
            if kind == "s" and v is not None:
                val = self.shared[int(v.text)]
            elif kind == "inlineStr":
                val = "".join(t.text or "" for t in c.iter(f"{M}t"))
            else:
                val = v.text if v is not None else None
            if val not in (None, ""):
                ref = c.get("r")
                out.setdefault(int(re.search(r"\d+", ref).group(0)), {})[re.match(r"[A-Z]+", ref).group(0)] = val.strip()
        return out


def sheet_month(name):
    """(year or None, month) from a sheet name like 'September 2026📚' or 'Feburary 💟'."""
    if SKIP_SHEETS.search(name):
        return None
    # Also allows a day between month and year: "August 1 2026" (but "November 25" is still Nov 2025).
    # and an apostrophe year: "APR'26".
    m = re.match(r"\s*([a-z]{3})[a-z]*\.?\s*['’]?\s*(?:\d{1,2}(?:st|nd|rd|th)?,?\s+(?=\d{4}))?(\d{4}|\d{2})?\b", name.lower())
    if not m or m.group(1) not in MONTHS:
        return None
    year = m.group(2)
    year = int(year) if year and len(year) == 4 else (2000 + int(year) if year and 20 <= int(year) <= 40 else None)
    return year, MONTHS.index(m.group(1)) + 1


def month_sheets(wb):
    """Month sheets, newest first: [{"sheet", "year", "month"}]."""
    found = []
    for name in wb.sheets:
        ym = sheet_month(name)
        if ym:
            found.append({"sheet": name, "year": ym[0], "month": ym[1]})
    found.sort(key=lambda s: (s["year"] or 0, s["month"]), reverse=True)
    return found


def is_np_workbook(path):
    """True if this looks like a New Patient Analysis workbook."""
    try:
        wb = Workbook(path)
    except (NPAnalysisError, KeyError):
        return False
    sheets = month_sheets(wb)
    if len(sheets) < 1:
        return False
    rows = wb.rows(sheets[0]["sheet"])
    return any(_header_map(r) for r in list(rows.values())[:50])


def _col_number(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def _amount(text):
    try:
        return float(text)
    except (TypeError, ValueError):
        return 0.0


def _header_map(cells):
    """Column letters for the columns we use, if this row is a header row."""
    by_label = {v.rstrip(":").strip().lower(): k for k, v in cells.items()}
    if "insurance type" in by_label:
        notes = [k for label, k in by_label.items() if label.startswith(("notes", "follow up"))]
        # Prime lenders: "Alphaeon Ran Yes / No", "Care Credit Ran Yes / No", ... ; subprime: "Subprime Co. Ran"
        prime = [k for label, k in by_label.items() if re.search(r"yes\s*/\s*no", label) and not label.startswith(("has ", "no "))]
        subprime = [k for label, k in by_label.items() if label.startswith("subprime") and label.endswith(" ran")]
        # Each prime lender's "Ran" column is followed by its "$Amt Approved" and "$Amt Coll" columns.
        order = sorted(cells, key=_col_number)
        lenders = []
        for ran in sorted(prime, key=_col_number):
            after = order[order.index(ran) + 1:]
            find = lambda start: next((k for k in after
                                       if cells[k].lower().lstrip("$").strip().startswith(start)), None)
            lenders.append({"name": re.sub(r"\s*(?:ran\s*)?yes\s*/\s*no.*$", "", cells[ran], flags=re.I).strip(), "ran": ran,
                            "approved": find("amt approved"), "collected": find("amt coll")})
        # Subprime: one set of columns -- "Subprime Co. Approved", "Subprime $Amt Approved", "Subprime $Amt Coll".
        sub_cols = {key: next((k for label, k in by_label.items() if label.startswith(start)), None)
                    for key, start in (("approved_by", "subprime co. approved"),
                                       ("approved", "subprime $amt approved"), ("collected", "subprime $amt coll"))}
        first = lambda *starts: next((k for label, k in by_label.items() if label.startswith(starts)), None)
        return {"pat_id": by_label.get("pat id"), "insurance": by_label["insurance type"], "notes": notes,
                "prime_ran": prime, "subprime_ran": subprime, "prime_lenders": lenders, "subprime": sub_cols,
                "tx_diagnosed": first("$ treatment diagnosed", "$ tx diagnosed", "treatment plan diagnosis"),
                "no_ssn_col": first("no ssn", "no itin"),   # some offices track it in its own column
                "has_appt": first("has an appointment", "pt has nv appt", "next visit sched"),
                "collection": by_label.get("collection"),   # Denticon's Collection column (payments are negative)
                "other_payment": first("other form of payment amount")}
    return None


def insurance_types(text):
    """Normalized insurance types in a cell, e.g. "Cash, Medi-Cal" -> {"cash", "medical"}."""
    return {_norm(part) for part in re.split(r"[,;/&]|\band\b", text or "") if part.strip()}


def _first_match(pattern, text_cells, header_labels):
    """{"column", "text"} for the first cell matching `pattern`: the column's header and just the
    matched words (e.g. "no itin or social"), never the rest of the note."""
    for col, value in text_cells:
        m = pattern.search(value)
        if m:
            return {"column": (header_labels.get(col) or col).strip(), "text": m.group(0).strip().lower()}
    return None


# Staff sometimes add the paste's dates after the office name: "Office: Dentist Of Montclair 09/02/26-09/04/26".
_TRAILING_DATES = re.compile(r"[\s,(]*\d{1,2}/\d{1,2}(?:/\d{2,4})?(?:\s*(?:-|–|to|thru)\s*\d{1,2}/\d{1,2}(?:/\d{2,4})?)?[\s)]*$", re.I)


def clean_office(name):
    """Office name without trailing dates or date ranges."""
    name = (name or "").strip()
    while True:
        trimmed = _TRAILING_DATES.sub("", name).strip()
        if trimmed == name or not trimmed:
            return name
        name = trimmed


def _excel_date(serial):
    return date(1899, 12, 30) + timedelta(days=int(float(serial)))


def dated_months(wb):
    """Month sheets with a year, newest first. A tab with no year ("October 🎃") is taken as the latest
    such month up to today, unless another tab already has that month and year (old year-less tabs)."""
    all_months = month_sheets(wb)
    months = [m for m in all_months if m["year"]]
    have = {(m["year"], m["month"]) for m in months}
    today = date.today()
    for m in all_months:
        if not m["year"]:
            year = today.year if m["month"] <= today.month else today.year - 1
            if (year, m["month"]) not in have:
                months.append({**m, "year": year})
                have.add((year, m["month"]))
    months.sort(key=lambda s: (s["year"], s["month"]), reverse=True)
    if not months:
        raise NPAnalysisError("No month sheets found (sheet names like 'September 2026')")
    return months


def parse_latest(path):
    """Counts from the latest month's sheet that has patients (e.g. 'September 2026').
    An empty newer tab (next month set up early) is skipped."""
    wb = Workbook(path)
    return months_with_data(wb, 1)[0]


def months_with_data(wb, how_many):
    """Parse month tabs newest first, skipping empty ones, until `how_many` have been read."""
    found, empty, fallback = [], [], None
    for month in dated_months(wb):
        try:
            try:
                parsed = parse_month(wb, month, found[0]["offices"][0]["office"] if found else fallback)
            except NoOfficeError:
                if fallback is not None or not (fallback := find_office(wb)):
                    raise
                parsed = parse_month(wb, month, fallback)
        except EmptyMonthError:
            empty.append(month["sheet"])
            continue
        parsed["empty_tabs_skipped"] = list(empty)
        found.append(parsed)
        if len(found) == how_many:
            break
    if not found:
        raise NPAnalysisError("None of the month tabs have any patients yet")
    return found


def parse_month(wb, chosen, default_office=None):
    """Counts from one month's sheet. `wb` is a downloaded Workbook or a Google Sheet
    (anything with .sheets and .rows(sheet)); `chosen` is an entry from month_sheets().
    `default_office` is used if this tab has no "Office: ..." line (another tab of the
    same workbook had one)."""
    rows = wb.rows(chosen["sheet"])
    cols, office, patient_rows, dates = None, None, 0, []
    counts = {m: 0 for m in INSURANCE_METRICS}
    counts["no_ssn_itin_patients"] = 0
    counts["financing_applicants"] = 0
    counts["existing_finance_patients"] = 0
    counts["prime_applicants"] = 0
    counts["prime_approved"] = 0
    counts["prime_apps_ran"] = 0   # every prime application: a patient who ran 4 lenders adds 4
    counts["prime_apps_approved"] = 0   # every approved prime application (of those ran)
    counts["prime_approved_amount"] = 0.0   # $ approved by prime lenders (sum of "$Amt Approved")
    counts["prime_collected_amount"] = 0.0  # $ collected through prime lenders (sum of "$Amt Coll")
    counts["prime_funded"] = 0   # patients with money collected through a prime lender
    counts["prime_approved_funded"] = 0   # approved prime patients (see prime_approved) who were also funded
    counts["subprime_applicants"] = 0   # a lender named under "Subprime Co. Ran"
    counts["subprime_approved"] = 0     # applied, and a lender named under "Subprime Co. Approved" or an amount
    counts["subprime_apps_ran"] = 0       # every subprime lender run (a patient run with 3 lenders adds 3)
    counts["subprime_apps_approved"] = 0  # every subprime lender named as approving
    counts["subprime_approved_amount"] = 0.0   # sum of "Subprime $Amt Approved"
    counts["subprime_collected_amount"] = 0.0  # sum of "Subprime $Amt Coll"
    counts["subprime_funded"] = 0   # patients with money collected through a subprime lender
    # Treatment: diagnosed $, collections, next visits and patients who started (paid).
    counts["np_rows"] = 0                 # new patients in the workbook (the denominator for per-NP numbers)
    counts["tx_diagnosed_amount"] = 0.0   # sum of "$ Treatment Diagnosed"
    counts["np_collection_wb"] = 0.0      # sum of Denticon "Collection" for these patients, as a positive amount
    counts["no_next_visit"] = 0           # "Has an Appointment" = No
    counts["started_tx"] = 0              # paid anything: Denticon collection, financing collected, or other payment
    # Prime + subprime combined; a patient who used both counts once.
    counts["fin_approved"] = 0            # approved by a prime or subprime lender
    counts["fin_funded"] = 0              # money collected through a prime or subprime lender
    counts["fin_approved_funded"] = 0     # approved (either) and funded (either)
    # Data check: patients with more collected than approved (prime per lender, or subprime) -- often a typo.
    counts["collected_over_approved"] = 0
    details = {m: [] for m in counts}   # which sheet rows were counted, and why (never names)
    for r in sorted(rows):
        cells = rows[r]
        header = _header_map(cells)
        if header:
            # Some sheets label the team's columns in one header row and leave "Pat ID" to Denticon's
            # own header rows further down; keep a Pat ID column already found.
            if not header["pat_id"] and cols and cols.get("pat_id"):
                header["pat_id"] = cols["pat_id"]
            cols, header_labels = header, cells
            continue
        # A pasted Denticon block's own header row ("Chart #", "Pat ID", "Collection", ...). Pastes are
        # sometimes shifted a column or two, so take Pat ID and Collection from this block's header.
        block = {v.strip().rstrip(":").lower(): k for k, v in cells.items()}
        if "pat id" in block:
            if cols:
                cols = {**cols, "pat_id": block["pat id"], "collection": block.get("collection", cols["collection"])}
            continue
        a = cells.get("A", "")
        m = OFFICE_LINE_RE.match(a)
        if m and not office:
            office = clean_office(m.group(1))
        if NUMBER_RE.match(a) and 40000 < float(a) < 60000:   # a date marker (Excel date number)
            dates.append(_excel_date(a))
        if cols and NUMBER_RE.match(cells.get(cols["pat_id"], "")):
            patient_rows += 1
            # The Insurance Type dropdown allows several choices ("Cash, Medi-Cal"); count the
            # patient if any of them matches.
            ins = insurance_types(cells.get(cols["insurance"]))
            for metric, keys in INSURANCE_KEYS.items():
                if ins & keys:
                    counts[metric] += 1
                    details[metric].append({"row": r, "insurance": cells.get(cols["insurance"], "")})
            # Treatment
            counts["np_rows"] += 1
            dx = _amount(cells.get(cols["tx_diagnosed"])) if cols["tx_diagnosed"] else 0
            if dx:
                counts["tx_diagnosed_amount"] += dx
                details["tx_diagnosed_amount"].append({"row": r, "total": f"${dx:,.2f}"})
            paid_denticon = -_amount(cells.get(cols["collection"])) if cols["collection"] else 0
            if paid_denticon:
                counts["np_collection_wb"] += paid_denticon
                details["np_collection_wb"].append({"row": r, "total": f"${paid_denticon:,.2f}"})
            if cols["has_appt"] and cells.get(cols["has_appt"], "").strip().lower() == "no":
                counts["no_next_visit"] += 1
                details["no_next_visit"].append({"row": r, "appt": cells[cols["has_appt"]].strip()})
            other_paid = _amount(cells.get(cols["other_payment"])) if cols["other_payment"] else 0
            fin_paid = sum(abs(_amount(cells.get(L["collected"]))) for L in cols["prime_lenders"] if L["collected"])
            if cols["subprime"]["collected"]:
                fin_paid += abs(_amount(cells.get(cols["subprime"]["collected"])))
            how = ([f"Denticon ${paid_denticon:,.2f}"] if paid_denticon > 0 else []) \
                + ([f"Financing ${fin_paid:,.2f}"] if fin_paid else []) + ([f"Other ${other_paid:,.2f}"] if other_paid > 0 else [])
            if how:
                counts["started_tx"] += 1
                details["started_tx"].append({"row": r, "paid": how})
            # Notes are sometimes typed into other columns (e.g. Chart #), so search every text cell
            # on the row -- except column A, which holds the referral type ("EXISTING PATIENT").
            text_cells = [(k, v) for k, v in sorted(cells.items(), key=lambda kv: _col_number(kv[0]))
                          if k != "A" and not NUMBER_RE.match(v)]
            hit = _first_match(NO_SSN_ITIN_RE, text_cells, header_labels)
            flag = cells.get(cols["no_ssn_col"], "").strip() if cols["no_ssn_col"] else ""
            if not hit and (flag.lower() in {"yes", "y", "x", "true", "✓", "✔"} or NO_SSN_ITIN_RE.search(flag)):
                hit = {"column": header_labels.get(cols["no_ssn_col"], "No SSN/ITIN").strip(), "text": flag.lower()}
            if hit:
                counts["no_ssn_itin_patients"] += 1
                details["no_ssn_itin_patients"].append({"row": r, **hit})
            # Applied for financing: "Yes" with any prime lender, or a subprime lender named.
            prime = [header_labels[c].split(" Ran")[0].strip() for c in cols["prime_ran"]
                     if cells.get(c, "").lower().startswith("y")]
            subprime = [part.strip() for c in cols["subprime_ran"] if cells.get(c, "").lower() not in NOT_RAN
                        for part in cells[c].split(",") if part.strip()]
            if prime or subprime:
                counts["financing_applicants"] += 1
                details["financing_applicants"].append({"row": r, "prime": prime, "subprime": subprime})
            # Prime: applied if "Ran" is Yes (as for financing applicants -- someone paying with an
            # existing account didn't apply). Approved if they applied and a prime lender shows an
            # approved amount, or a collected amount (approved even if the amount wasn't filled in).
            applied, approved, apps_approved = [], [], []
            for L in cols["prime_lenders"]:
                amt = max(_amount(cells.get(L["approved"])), 0) if L["approved"] else 0
                coll = abs(_amount(cells.get(L["collected"]))) if L["collected"] else 0
                ran = cells.get(L["ran"], "").lower().startswith("y")
                if ran:
                    applied.append(L["name"])
                if amt or coll:
                    label = f"{L['name']} ${amt:,.0f}" if amt else f"{L['name']} (collected ${coll:,.0f})"
                    approved.append(label)
                    if ran:
                        apps_approved.append(label)
            if applied:
                counts["prime_applicants"] += 1
                details["prime_applicants"].append({"row": r, "applied": applied})
                counts["prime_apps_ran"] += len(applied)
                details["prime_apps_ran"].append({"row": r, "applied": applied, "apps": len(applied)})
            if applied and approved:
                counts["prime_approved"] += 1
                details["prime_approved"].append({"row": r, "approved": approved})
                funded_by = [f"{L['name']} ${abs(_amount(cells.get(L['collected']))):,.2f}"
                             for L in cols["prime_lenders"] if L["collected"] and _amount(cells.get(L["collected"]))]
                if funded_by:
                    counts["prime_approved_funded"] += 1
                    details["prime_approved_funded"].append({"row": r, "lenders": funded_by})
            # Dollar totals: every prime amount on the row, including money collected through an
            # existing account (no new application).
            appr_amts = [(L["name"], _amount(cells.get(L["approved"]))) for L in cols["prime_lenders"] if L["approved"]]
            coll_amts = [(L["name"], abs(_amount(cells.get(L["collected"])))) for L in cols["prime_lenders"] if L["collected"]]
            appr_amts = [(name, v) for name, v in appr_amts if v]
            coll_amts = [(name, v) for name, v in coll_amts if v]
            if appr_amts:
                counts["prime_approved_amount"] += sum(v for _, v in appr_amts)
                details["prime_approved_amount"].append(
                    {"row": r, "lenders": [f"{name} ${v:,.2f}" for name, v in appr_amts], "total": f"${sum(v for _, v in appr_amts):,.2f}"})
            if coll_amts:
                counts["prime_funded"] += 1
                details["prime_funded"].append({"row": r, "lenders": [f"{name} ${v:,.2f}" for name, v in coll_amts]})
                counts["prime_collected_amount"] += sum(v for _, v in coll_amts)
                details["prime_collected_amount"].append(
                    {"row": r, "lenders": [f"{name} ${v:,.2f}" for name, v in coll_amts], "total": f"${sum(v for _, v in coll_amts):,.2f}"})
            # Data check: more collected than approved, for a lender with an approved amount.
            over = [{"kind": "Prime", "lender": L["name"], "approved": f"${appr:,.2f}", "collected": f"${coll:,.2f}"}
                    for L in cols["prime_lenders"] if L["approved"] and L["collected"]
                    for appr, coll in [(_amount(cells.get(L["approved"])), abs(_amount(cells.get(L["collected"]))))]
                    if appr > 0 and coll > appr]
            # Subprime dollar totals: every amount on the row.
            sc = cols["subprime"]
            # The sheet has one subprime amount per row, so show who approved it (or who was run).
            named = [p.strip() for p in cells.get(sc["approved_by"], "").split(",")
                     if p.strip() and p.strip().lower() not in NOT_RAN] if sc["approved_by"] else []
            ran_with = ", ".join(named or subprime) or "(lender not named)"
            s_appr_amt = _amount(cells.get(sc["approved"])) if sc["approved"] else 0
            s_coll_amt = abs(_amount(cells.get(sc["collected"]))) if sc["collected"] else 0
            if s_appr_amt:
                counts["subprime_approved_amount"] += s_appr_amt
                details["subprime_approved_amount"].append({"row": r, "lenders": ran_with, "total": f"${s_appr_amt:,.2f}"})
            if s_appr_amt > 0 and s_coll_amt > s_appr_amt:
                over.append({"kind": "Subprime", "lender": ran_with, "approved": f"${s_appr_amt:,.2f}",
                             "collected": f"${s_coll_amt:,.2f}"})
            if over:
                counts["collected_over_approved"] += 1
                details["collected_over_approved"].append({"row": r, "issues": over, "text": [
                    f"{o['kind']} {o['lender']}: collected {o['collected']}, approved {o['approved']}" for o in over]})
            if s_coll_amt:
                counts["subprime_funded"] += 1
                details["subprime_funded"].append({"row": r, "lenders": ran_with, "total": f"${s_coll_amt:,.2f}"})
                counts["subprime_collected_amount"] += s_coll_amt
                details["subprime_collected_amount"].append({"row": r, "lenders": ran_with, "total": f"${s_coll_amt:,.2f}"})
            # Subprime: applied if a lender is named under "Subprime Co. Ran"; approved if they applied and
            # a lender is named under "Subprime Co. Approved" or an amount was approved or collected.
            if subprime:
                counts["subprime_applicants"] += 1
                details["subprime_applicants"].append({"row": r, "subprime": subprime})
                counts["subprime_apps_ran"] += len(subprime)
                details["subprime_apps_ran"].append({"row": r, "subprime": subprime, "apps": len(subprime)})
                sc = cols["subprime"]
                approved_by = [p.strip() for p in cells.get(sc["approved_by"], "").split(",")
                               if p.strip() and p.strip().lower() not in NOT_RAN] if sc["approved_by"] else []
                s_amt = _amount(cells.get(sc["approved"])) if sc["approved"] else 0
                s_coll = abs(_amount(cells.get(sc["collected"]))) if sc["collected"] else 0
                if approved_by or s_amt or s_coll:
                    # An amount with no lender named still counts as one approved app.
                    n_apps = len(approved_by) or 1
                    counts["subprime_apps_approved"] += n_apps
                    details["subprime_apps_approved"].append(
                        {"row": r, "approved": approved_by or ["(lender not named)"], "apps": n_apps})
                    counts["subprime_approved"] += 1
                    details["subprime_approved"].append({
                        "row": r, "approved": approved_by or ["(lender not named)"],
                        "amount": f"${s_amt:,.2f}" if s_amt else ("—" if not s_coll else f"collected ${s_coll:,.2f}")})
            if apps_approved:
                counts["prime_apps_approved"] += len(apps_approved)
                details["prime_apps_approved"].append({"row": r, "approved": apps_approved, "apps": len(apps_approved)})
            hit = _first_match(EXISTING_FIN_RE, text_cells, header_labels)
            if hit:
                counts["existing_finance_patients"] += 1
                details["existing_finance_patients"].append({"row": r, "lender": lender_named(hit["text"]), **hit})

    # Combined prime + subprime patient counts, from the per-patient lists above.
    by_row = lambda key: {d["row"]: d for d in details[key]}
    p_appr, s_appr = by_row("prime_approved"), by_row("subprime_approved")
    p_fund, s_fund = by_row("prime_funded"), by_row("subprime_funded")
    source = lambda r, p, s: "Prime + Subprime" if r in p and r in s else ("Prime" if r in p else "Subprime")
    approved_rows, funded_rows = sorted(p_appr.keys() | s_appr.keys()), sorted(p_fund.keys() | s_fund.keys())
    counts["fin_approved"] = len(approved_rows)
    details["fin_approved"] = [{"row": r, "source": source(r, p_appr, s_appr)} for r in approved_rows]
    counts["fin_funded"] = len(funded_rows)
    details["fin_funded"] = [{"row": r, "source": source(r, p_fund, s_fund)} for r in funded_rows]
    both = [r for r in approved_rows if r in funded_rows]
    counts["fin_approved_funded"] = len(both)
    details["fin_approved_funded"] = [{"row": r, "source": source(r, p_fund, s_fund)} for r in both]

    if cols is not None and patient_rows == 0:
        raise EmptyMonthError(f"Sheet {chosen['sheet']!r} has no patients yet")
    if cols is None and len(rows) <= 3:          # a blank or title-only tab
        raise EmptyMonthError(f"Sheet {chosen['sheet']!r} is empty")
    if cols is None:
        raise NPAnalysisError(f"Sheet {chosen['sheet']!r} has no 'Insurance Type' column, so Medi-Cal and HMO "
                              "counts aren't available for that month (it was added to the sheets in November 2025)")
    office = office or default_office
    if not office:
        raise NoOfficeError(f"Couldn't find the office name ('Office: …') on sheet {chosen['sheet']!r} or any other month tab")

    period = {}
    if chosen["year"]:
        first = date(chosen["year"], chosen["month"], 1)
        in_month = [d for d in dates if (d.year, d.month) == (first.year, first.month)]
        period["from"] = first.isoformat()
        if in_month:
            period["thru"] = max(in_month).isoformat()
    return {
        "report_type": "np_analysis",
        "offices": [{"office": office, **{k: round(v, 2) if isinstance(v, float) else v for k, v in counts.items()},
                     "details": details}],
        "skipped": [],
        "grand_total": None,
        "period": period,
        "patient_rows": patient_rows,
        "sheet": chosen["sheet"],
    }
