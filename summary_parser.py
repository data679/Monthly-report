"""Parse Denticon / PlanetDDS "Executive Summary" reports (PDF or XLSX).

Both formats are reduced to the same shape -- a list of (label, numbers) rows --
so one extractor handles either file type. A report can hold many offices; each
office's block starts at its "New Patients :" row, with the office name on the
row just above it.
"""

import re
import subprocess
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

import lender_journal
import np_analysis

NUM_RE = re.compile(r"^\(?-?\$?[\d,]*\.?\d+\)?$")
DATE_RE = r"(\d{1,2}/\d{1,2}/\d{4})"
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


class ParseError(Exception):
    pass


def to_number(token):
    t = token.strip().replace(",", "").replace("$", "")
    negative = t.startswith("(") and t.endswith(")")
    t = t.strip("()")
    value = float(t)
    return -value if negative else value


def clean_label(text):
    return re.sub(r"\s+", " ", text).strip().rstrip(":").strip()


# ---------------------------------------------------------------- Word (.doc / .docx)

def _rows_from_tokens(tokens):
    """(label, numbers) rows from a report's text pieces in reading order: a text piece starts a row and
    the numbers after it belong to it. "Patient Type :" + "Ortho" become one label, as in the PDF."""
    rows = []
    for t in tokens:
        words = t.split()
        if rows and words and all(NUM_RE.match(w) for w in words):
            rows[-1][1].extend(to_number(w) for w in words)
        elif rows and rows[-1][0].endswith(":") and not rows[-1][1]:
            rows[-1][0] += " " + t
        else:
            rows.append([t, []])
    return [(clean_label(label), nums) for label, nums in rows]


def doc_rows(path):
    """Rows and period from a Denticon report saved as Word."""
    tokens = doc_tokens(path)
    joined = " ".join(tokens)
    period = {}
    m = re.search(r"From\s*:\s*" + DATE_RE, joined)
    if m:
        period["from"] = m.group(1)
    m = re.search(r"Thru\s*:\s*" + DATE_RE, joined)
    if m:
        period["thru"] = m.group(1)
    return _rows_from_tokens(tokens), period


def doc_tokens(path):
    """Rows and period from a Denticon report saved as Word. A .doc (Word 97) keeps its text as UTF-16
    with \r between paragraphs and \x07 between table cells; a .docx is zipped XML. Nothing is executed."""
    raw = Path(path).read_bytes()
    if raw[:4] == b"PK\x03\x04":
        try:
            with zipfile.ZipFile(path) as z:
                xml = z.read("word/document.xml")
        except (zipfile.BadZipFile, KeyError):
            raise ParseError("Could not read this Word file")
        w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        tokens = ["".join(t.text or "" for t in p.iter(w + "t")).strip() for p in ET.fromstring(xml).iter(w + "p")]
    elif raw[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        text = "".join(m.decode("utf-16-le") for m in
                       re.findall(rb"(?:[\x07\x09\x0b\x0d\x20-\x7e\xa0-\xff]\x00){6,}", raw))
        tokens = [t.strip() for t in re.split(r"[\r\x07\x0b]+", text)]
    else:
        raise ParseError("This doesn't look like a Word document (.doc or .docx)")
    tokens = [t for t in tokens if t]
    if not tokens:
        raise ParseError("No text found in this Word file")
    return tokens


# ---------------------------------------------------------------- PDF

def pdf_rows(path):
    try:
        text = subprocess.run(
            ["pdftotext", "-layout", str(path), "-"],
            capture_output=True, text=True, check=True,
        ).stdout
    except FileNotFoundError:
        raise ParseError("pdftotext is not installed (sudo apt-get install poppler-utils)")
    except subprocess.CalledProcessError as e:
        raise ParseError(f"Could not read PDF: {e.stderr.strip() or e}")

    rows = []
    for line in text.replace("\f", "\n").splitlines():
        if not line.strip():
            continue
        # Columns in the layout text are separated by 2+ spaces.
        parts = [p for p in re.split(r"\s{2,}", line.strip()) if p]
        label_parts, nums = [], []
        for i, p in enumerate(parts):
            words = p.split()
            # Numbers are sometimes only one space apart ("961,916.29 -1,113,007.64").
            if i > 0 and all(NUM_RE.match(w) for w in words):
                nums.extend(to_number(w) for w in words)
            else:
                label_parts.append(p)
        rows.append((clean_label(" ".join(label_parts)), nums))

    period = {}
    m = re.search(r"From\s*:\s*" + DATE_RE, text)
    if m:
        period["from"] = m.group(1)
    m = re.search(r"Thru\s*:\s*" + DATE_RE, text)
    if m:
        period["thru"] = m.group(1)
    return rows, period


# ---------------------------------------------------------------- XLSX

def _col_index(ref):
    letters = re.match(r"[A-Z]+", ref).group(0)
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def xlsx_rows(path):
    try:
        z = zipfile.ZipFile(path)
    except zipfile.BadZipFile:
        raise ParseError("Not a valid .xlsx file")

    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
        for si in root.findall("m:si", NS):
            shared.append("".join(t.text or "" for t in si.iter(f"{{{NS['m']}}}t")))

    # Sheets in workbook order.
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    targets = {r.get("Id"): r.get("Target") for r in rels}
    rid_attr = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    sheet_paths = []
    for s in wb.find("m:sheets", NS):
        target = targets[s.get(rid_attr)].lstrip("/")
        sheet_paths.append(target if target.startswith("xl/") else "xl/" + target)

    rows = []
    for sp in sheet_paths:
        root = ET.fromstring(z.read(sp))
        for row in root.iter(f"{{{NS['m']}}}row"):
            cells = []
            for c in row.findall("m:c", NS):
                kind = c.get("t")
                if kind == "s":
                    v = c.find("m:v", NS)
                    val = shared[int(v.text)] if v is not None else ""
                elif kind == "inlineStr":
                    val = "".join(t.text or "" for t in c.iter(f"{{{NS['m']}}}t"))
                else:
                    v = c.find("m:v", NS)
                    text = v.text if v is not None and v.text else ""
                    try:
                        val = float(text) if text and kind not in ("str", "e", "b") else text
                    except ValueError:        # e.g. a formula showing "#DIV/0!"
                        val = text
                cells.append((_col_index(c.get("r")), val))
            cells.sort()
            label = " ".join(v for _, v in cells if isinstance(v, str) and v.strip())
            nums = [v for _, v in cells if isinstance(v, float)]
            if label or nums:
                rows.append((clean_label(label), nums))
        rows.append(("", []))  # sheet boundary
    return rows, {}


# ---------------------------------------------------------------- extraction

def _office_blocks(rows):
    """Yield (office_name, block_rows) for each office in the report."""
    starts = [i for i, (label, _) in enumerate(rows) if label == "New Patients"]
    for n, start in enumerate(starts):
        name = ""
        j = start - 1
        while j >= 0 and not rows[j][0] and not rows[j][1]:
            j -= 1
        if j >= 0 and not rows[j][1]:
            name = rows[j][0]
        end = starts[n + 1] - 1 if n + 1 < len(starts) else len(rows)
        yield name, rows[start:end]


def _extract(block):
    np_first_visit = None
    existing = 0.0
    production_total = None
    insurance_collection = None
    total_collection = None
    section = None

    for label, nums in block:
        first = nums[0] if nums else None
        upper = label.upper()
        if label == "New Patients by First Visit Date":
            np_first_visit = first
        elif label == "Production":
            section = "production"
        elif label == "Collection":
            section = "collection"
            total_collection = first
        elif label == "Referrals":
            section = "referrals"
        elif label in ("Adjustments", "Dentist Production", "Hygienist Production"):
            section = None
        elif section == "production" and label == "Total" and production_total is None:
            production_total = first
            section = None
        elif section == "collection" and label == "Insurance" and insurance_collection is None:
            insurance_collection = first
        elif section == "referrals" and upper == "EXISTING PATIENT":
            existing = first or 0.0

    return {
        "np_first_visit": np_first_visit,
        "existing_patient_referrals": existing,
        "insurance_collection": insurance_collection,
        "total_collection": total_collection,
        "total_production": production_total,
    }


# ---------------------------------------------------------------- Referral Production Listing

OFFICE_TOTAL_RE = re.compile(r"^Total for Office\s*,\s*(.+?)\s*\[\s*\d+")
GRAND_TOTAL_RE = re.compile(r"^Grand Total\b")
# Columns on a total line: # Visits, #Proc, Production, Collection, Adjs
COLLECTION_INDEX = 3


def parse_referral(rows):
    """NP collection per office from a Referral Production Listing.

    Only the "Total for Office" lines are read; patient rows are never kept.
    """
    offices, grand_total = [], None
    i = 0
    while i < len(rows):
        label, nums = rows[i]
        # In the PDF a long office name can push the numbers onto the next line.
        if (label.startswith("Total for Office") or GRAND_TOTAL_RE.match(label)) and not nums \
                and i + 1 < len(rows):
            label = f"{label} {rows[i + 1][0]}".strip()
            nums = rows[i + 1][1]
            i += 1
        m = OFFICE_TOTAL_RE.match(label)
        if m or GRAND_TOTAL_RE.match(label):
            if len(nums) != 5:
                raise ParseError(f"Unexpected numbers on total line {label!r}")
            collection = abs(nums[COLLECTION_INDEX])
            if m:
                offices.append({"office": m.group(1).strip(), "np_collection": round(collection, 2)})
            else:
                grand_total = round(collection, 2)
        i += 1

    if not offices:
        raise ParseError("No 'Total for Office' lines found in this Referral Production Listing")
    # For the Maps tab: each new patient's ZIP, Pat ID and collection (never the name), when the report is grouped
    # by Zip Code. Used only if an office's patients add up to its own total.
    patients = _referral_patients(rows)
    grouped = any(label.startswith("Zip Code") for label, _ in rows)
    for o in offices:
        mine = patients.get(o["office"], [])
        if mine and abs(sum(p["collection"] for p in mine) - o["np_collection"]) <= 0.05:
            o["details"] = {"_map": mine}
    total = round(sum(o["np_collection"] for o in offices), 2)
    if grand_total is not None and abs(total - grand_total) > 0.01:
        raise ParseError(f"Office collections add up to {total:,.2f} but the report's Grand Total is "
                         f"{grand_total:,.2f}; not logging a mismatched report")
    return {"report_type": "referral", "offices": offices, "skipped": [], "grand_total": grand_total,
            "map_note": None if grouped else "This report isn't grouped by Zip Code, so it won't add new patients to the "
                                             "Maps tab. Run it with Group by Zip Code to include them."}


def _referral_patients(rows):
    """{office: [{zip, referral_type, pat_id, visits, production, collection}]} from a Referral Production Listing
    grouped by Zip Code. A patient line's numbers are [chart #,] Pat ID, # visits, #proc, production, collection,
    adjustments; its text (the patient's name) is ignored."""
    out, office, referral, zip_code = {}, None, None, None
    for label, nums in rows:
        if (m := re.match(r"Office\s*:\s*(.+)$", label)):
            office, referral, zip_code = m.group(1).strip(), None, None
        elif (m := re.match(r"Referral Type\s*:\s*(.+)$", label)):
            referral, zip_code = m.group(1).strip(), None
        elif (m := re.match(r"Zip Code\s*:\s*(\d{5})", label)):
            zip_code = m.group(1)
        elif office and zip_code and not label.startswith("Total") and \
                (len(nums) in (6, 7) or (len(nums) == 5 and re.match(r"\d{6,9}\b", label))):
            # In a PDF the Pat ID is at the start of the line's text: "90228259 <name>" [visits, proc, ...]
            n = nums[-6:] if len(nums) != 5 else [float(re.match(r"\d{6,9}", label).group()), *nums]
            if float(n[0]).is_integer() and 100000 <= n[0] < 10**9:
                out.setdefault(office, []).append({
                    "zip": zip_code, "referral_type": referral or "", "pat_id": str(int(n[0])),
                    "visits": int(n[1]), "production": round(n[3], 2), "collection": round(abs(n[4]), 2)})
    return out


def parse_patient_types(path):
    """Pat ID -> "General" / "Ortho" from Denticon's Patient List - Address (one office per file), for the Maps
    tab's filter. Only the office, Pat IDs and types are returned; names and addresses are never kept."""
    if not str(path).lower().endswith((".xlsx", ".xlsm")):
        raise ParseError("Upload the Patient List - Address as an Excel file (.xlsx)")
    wb = np_analysis.Workbook(path)
    rows = wb.rows(next(iter(wb.sheets)))
    keys = sorted(rows)
    if not any("patient list" in v.lower() for r in keys[:6] for v in rows[r].values()):
        raise ParseError("This doesn't look like Denticon's Patient List - Address report")
    office = next((v.strip() for r in keys[:4] for v in rows[r].values()
                   if v.strip() and "patient list" not in v.lower()), None)
    head_r = next((r for r in keys[:20] if any(v.strip().lower() == "pat id" for v in rows[r].values())), None)
    if not office or head_r is None:
        raise ParseError("Couldn't find the office name or the header row (Pat ID) in this patient list")
    head = {re.sub(r"\s+", " ", v).strip().lower(): c for c, v in rows[head_r].items()}
    pid_c = head.get("pat id")
    type_c = next((c for h, c in head.items() if "general" in h or "ortho" in h), None)
    if not type_c:
        raise ParseError("This patient list has no 'General / Ortho' column")
    types = {}
    for r in keys:
        if r <= head_r:
            continue
        pid = rows[r].get(pid_c, "").split(".")[0]
        t = rows[r].get(type_c, "").strip().lower()
        if pid.isdigit():
            types[pid] = "Ortho" if t.startswith("ortho") else "General" if t.startswith("gen") else ""
    if not types:
        raise ParseError("No patients found in this patient list")
    return {"office": office, "types": types}


# ---------------------------------------------------------------- Daily Journal (Ortho)

SUMMARY_FOR_RE = re.compile(r"\bSummary for\s+(.+)$")


def parse_ortho_journal(rows):
    """Ortho collection per office from a Daily Journal - Detail run for Patient Type: Ortho.

    Reads only the "Summary for <office>" blocks: the Total column of "Total Payments".
    Transaction rows (patients) are never kept.
    """
    if not any(label.startswith("Patient Type") and "Ortho" in label for label, _ in rows):
        raise ParseError("This Daily Journal isn't filtered to Patient Type: Ortho, so its payments "
                         "aren't ortho collection. Re-run it with Patient Type = Ortho.")
    totals, order, current, grand_total = {}, [], None, None
    i = 0
    while i < len(rows):
        label, nums = rows[i]
        m = SUMMARY_FOR_RE.search(label)
        if label.startswith("Grand Total Summary"):
            current = "__grand__"
        elif m:
            current = m.group(1).strip()
            if current not in totals:
                totals[current] = 0.0
                order.append(current)
        elif label == "Total Payments" and current:
            if not nums and i + 1 < len(rows) and not rows[i + 1][0]:
                nums = rows[i + 1][1]
                i += 1
            if not nums:
                raise ParseError(f"No amount on the Total Payments line for {current}")
            amount = round(abs(nums[-1]), 2)   # last column is Total (Est. Pat. + Est. Ins.)
            if current == "__grand__":
                grand_total = amount
            else:
                totals[current] = amount
            current = None   # only the first Total Payments after each summary heading
        i += 1

    if not order:
        raise ParseError("No 'Summary for <office>' sections found in this Daily Journal")
    offices = [{"office": o, "ortho_collection": totals[o]} for o in order]
    total = round(sum(totals.values()), 2)
    if grand_total is not None and abs(total - grand_total) > 0.01:
        raise ParseError(f"Office ortho payments add up to {total:,.2f} but the report's Grand Total is "
                         f"{grand_total:,.2f}; not logging a mismatched report")
    return {"report_type": "ortho_journal", "offices": offices, "skipped": [], "grand_total": grand_total}


# ---------------------------------------------------------------- Treatment Plan Status (ortho referrals)

OFFICE_COUNT_RE = re.compile(r"^Total for office\s*,\s*(.+?)\s*\[\s*(\d+)\s*\]?", re.I)
PROVIDER_COUNT_RE = re.compile(r"^Total for provider\s*,.*\[\s*(\d+)\s*\]", re.I)
ORTHO_REFERRAL_CODE = "ZREFORTH"


def parse_tx_plan_status(rows):
    """# of ortho referrals per office: the count in "Total for office, <name> [ n ]".

    Only office and provider total lines are read; patient lines are never kept.
    """
    proc_filter = next((label for label, _ in rows if label.startswith("Proc Codes")), None)
    if proc_filter is not None and ORTHO_REFERRAL_CODE not in proc_filter:
        raise ParseError(f"This Treatment Plan Status report isn't filtered to proc code {ORTHO_REFERRAL_CODE}, "
                         f"so its counts aren't ortho referrals ({proc_filter}).")
    offices, provider_sum = [], 0
    i = 0
    while i < len(rows):
        label = rows[i][0]
        # A long office name can push the closing "[ n ]" onto the next line.
        if label.lower().startswith("total for office") and "]" not in label and i + 1 < len(rows):
            label = f"{label} {rows[i + 1][0]}"
            i += 1
        m = OFFICE_COUNT_RE.match(label)
        p = PROVIDER_COUNT_RE.match(label)
        if m:
            office, count = m.group(1).strip(), int(m.group(2))
            if provider_sum and provider_sum != count:
                raise ParseError(f"{office}: office total says {count} referrals but its providers add up to "
                                 f"{provider_sum}; not logging a report that doesn't add up")
            offices.append({"office": office, "ortho_referrals": count})
            provider_sum = 0
        elif p:
            provider_sum += int(p.group(1))
        i += 1
    if not offices:
        raise ParseError("No 'Total for office' lines found in this Treatment Plan Status report")
    return {"report_type": "ortho_referrals", "offices": offices, "skipped": [], "grand_total": None,
            "filter_confirmed": proc_filter is not None}


def parse_report(path, filename=None):
    """Return {"report_type", "period", "offices", "skipped"} for a supported report."""
    name = (filename or str(path)).lower()
    if name.endswith((".xlsx", ".xlsm")) and np_analysis.is_np_workbook(path):
        try:
            return np_analysis.parse_latest(path)
        except np_analysis.NPAnalysisError as e:
            raise ParseError(str(e))
    if name.endswith(".pdf"):
        rows, period = pdf_rows(path)
    elif name.endswith((".xlsx", ".xlsm")):
        rows, period = xlsx_rows(path)
    elif name.endswith((".doc", ".docx")):
        rows, period = doc_rows(path)
    else:
        raise ParseError("Upload a .pdf, .xlsx or Word (.doc, .docx) report")

    heading = [label for label, _ in rows[:20]]
    if "Referral Production Listing" in heading:
        result = parse_referral(rows)
        result["period"] = period
        return result
    if any(h.startswith("Treatment Plan Status") for h in heading):
        result = parse_tx_plan_status(rows)
        result["period"] = period
        return result
    if any(h.startswith("Daily Journal") for h in heading):
        ortho = any(label.startswith("Patient Type") and "Ortho" in label for label, _ in rows)
        if not ortho:
            # All patients: the lenders' payments and fees per patient, to check the NP audit sheets.
            try:
                if name.endswith(".pdf"):
                    lines = subprocess.run(["pdftotext", "-layout", str(path), "-"], capture_output=True,
                                           text=True, check=True).stdout.replace("\f", "\n").splitlines()
                    result = lender_journal.from_pdf_lines(lines)
                elif name.endswith((".doc", ".docx")):
                    result = lender_journal.from_tokens(doc_tokens(path))
                else:
                    raise ParseError("Save the Daily Journal as PDF or Word to check lender payments")
            except lender_journal.JournalError as e:
                raise ParseError(f"{e}. (For ortho collection, run it with Patient Type = Ortho.)")
            result["period"] = period
            return result
        result = parse_ortho_journal(rows)
        result["period"] = period
        return result

    offices, skipped = [], []
    for office, block in _office_blocks(rows):
        data = _extract(block)
        if data["np_first_visit"] is None and not data["total_production"]:
            skipped.append(office or "(unnamed)")
            continue
        npfv = data["np_first_visit"] or 0
        offices.append({
            "office": office or "(unnamed)",
            "total_collection": round(abs(data["total_collection"] or 0), 2),
            "insurance_collection": round(abs(data["insurance_collection"] or 0), 2),
            "np_first_visit": int(npfv),
            "existing_patient_referrals": int(data["existing_patient_referrals"]),
            "new_patients": int(npfv - data["existing_patient_referrals"]),
            "total_production": round(data["total_production"] or 0, 2),
        })

    if not offices:
        raise ParseError("This doesn't look like an Executive Summary or Referral Production Listing")
    return {"report_type": "executive_summary", "period": period, "offices": offices, "skipped": skipped}


if __name__ == "__main__":
    import json
    import sys
    for p in sys.argv[1:]:
        print(json.dumps(parse_report(p), indent=2))
