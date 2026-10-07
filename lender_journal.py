"""Lender payments per patient from a Denticon Daily Journal - Detail run for all patients
(Patient Type: Both, Transactions: Payments, Adjustments), saved as Word (.doc/.docx) or PDF.

Each financing lender shows up twice per funded patient: a payment line ("PMT PAT - Alphaeon") and a
merchant-fee adjustment ("ADJUST ALPHAEON FEE"). The NP audit sheets record the amount before the fee,
so payment + fee is what gets compared. Every office's detail lines are checked against Denticon's own
"Summary for <office>" totals before anything is returned. Patient names are never kept; Pat IDs are
returned for matching and the app replaces them with keyed hashes before saving or sending them anywhere.
"""

import re

# (pattern in the description, lender, prime?) -- order matters: "Access Alph" before "Alphaeon".
LENDERS = [
    (re.compile(r"access\s*alph", re.I), "Access Alph", False),
    (re.compile(r"alph", re.I), "Alphaeon", True),
    (re.compile(r"care\s*credit", re.I), "Care Credit", True),
    (re.compile(r"patient\s*fi", re.I), "PatientFi", True),
    (re.compile(r"proceed", re.I), "Proceed", True),
    (re.compile(r"\bhfd\b", re.I), "HFD", False),
    (re.compile(r"covered\s*care", re.I), "Covered Care", False),
    (re.compile(r"cherry", re.I), "Cherry", False),
    (re.compile(r"sunbit", re.I), "Sunbit", False),
]
PRIME = {name for _, name, prime in LENDERS if prime}
PAYMENT_RE = re.compile(r"^\s*(pmt\s*pat|pymt|pt\s*pymt)", re.I)
FEE_RE = re.compile(r"^\s*adj(ust)?\b.*\b(fee|merch)", re.I)
MONEY_RE = re.compile(r"^-?\(?[\d,]*\d\.\d\d\)?$")
PAT_ID_RE = re.compile(r"^\d{7,9}$")
DATE_RE = r"(\d{1,2}/\d{1,2}/\d{4})"


class JournalError(Exception):
    pass


def classify(desc):
    """("payment" | "fee", lender) for a financing line, else None."""
    kind = "payment" if PAYMENT_RE.search(desc) else "fee" if FEE_RE.search(desc) else None
    if not kind:
        return None
    squeezed = re.sub(r"\s+", "", desc)   # PDFs break words across lines: "CARECREDI" / "T"
    for pat, name, _ in LENDERS:
        if pat.search(desc) or pat.search(squeezed):
            return kind, name
    return None


def _money(t):
    t = t.replace(",", "")
    neg = t.startswith("(") or t.startswith("-")
    return -float(t.strip("()-")) if neg else float(t.strip("()"))


def from_tokens(tokens):
    """Parse a Word export: text pieces in reading order (see summary_parser.doc_rows)."""
    joined = " ".join(tokens)
    m = re.search(r"Patient Type\s*:\s*(\w+)", joined)
    patient_type = m.group(1) if m else None
    transactions = (re.search(r"Transactions\s*:\s*([A-Za-z, ]+?)(?:\s{2,}|Providers|Patient Type|$)", joined) or [None, ""])[1]
    period = {}
    if (m := re.search(r"From\s*:\s*" + DATE_RE, joined)):
        period["from"] = m.group(1)
    if (m := re.search(r"Thru\s*:\s*" + DATE_RE, joined)):
        period["thru"] = m.group(1)

    detail, summary = {}, {}      # detail: office -> [(pat, lender, kind, amount)]; summary: (office, lender, kind) -> amount
    office, pat, in_summary, summary_office = None, None, False, None
    for i, t in enumerate(tokens):
        if (m := re.match(r"Office\s*:\s*(.+)$", t)):
            office, pat, in_summary = m.group(1).strip(), None, False
            continue
        if (m := re.match(r"Summary for\s+(.+)$", t)):
            in_summary, summary_office, pat = True, m.group(1).strip(), None
            continue
        if t.startswith("Grand Total"):
            in_summary, summary_office, pat = True, None, None
            continue
        if not in_summary and PAT_ID_RE.match(t):
            pat = t
            continue
        c = classify(t)
        if not c:
            continue
        nums = [_money(x) for x in tokens[i + 1:i + 4] if MONEY_RE.match(x)]
        if in_summary:
            if summary_office and nums:
                summary[(summary_office, c[1], c[0])] = summary.get((summary_office, c[1], c[0]), 0) + abs(nums[-1])
        elif pat and office:
            amount = next((abs(n) for n in nums if n), 0.0)
            detail.setdefault(office, []).append((pat, c[1], c[0], amount))
    return _result(detail, summary, period, patient_type, transactions)


def from_pdf_lines(lines):
    """Parse a PDF export (pdftotext -layout lines). A detail line starts with the Pat ID; its description
    can wrap onto the next line ("PMT PAT-" / "Alphaeon")."""
    text = "\n".join(lines)
    m = re.search(r"Patient Type\s*:\s*(\w+)", text)
    patient_type = m.group(1) if m else None
    transactions = (re.search(r"Transactions\s*:\s*([A-Za-z, ]+)", text) or [None, ""])[1]
    period = {}
    if (m := re.search(r"From\s*:\s*" + DATE_RE, text)):
        period["from"] = m.group(1)
    if (m := re.search(r"Thru\s*:\s*" + DATE_RE, text)):
        period["thru"] = m.group(1)
    detail, summary = {}, {}
    office, in_summary, summary_office = None, False, None
    for i, line in enumerate(lines):
        s = line.strip()
        if (m := re.match(r"Office\s*:\s*(.+?)(\s{2,}|$)", s)):
            office, in_summary = m.group(1).strip(), False
        if (m := re.search(r"Summary for\s+(.+?)(\s{2,}|$)", s)):
            in_summary, summary_office = True, m.group(1).strip()
            continue
        if s.startswith("Grand Total"):
            in_summary, summary_office = True, None
        nums = [_money(x) for x in re.findall(r"-?\(?[\d,]*\d\.\d\d\)?", s)]
        nxt = lines[i + 1].strip() if i + 1 < len(lines) else ""
        wrap = nxt if nxt and not re.search(r"\d\.\d\d", nxt) and not PAT_ID_RE.match(nxt.split()[0] if nxt.split() else "") else ""
        if in_summary:
            label = re.split(r"\s{2,}", s)[0] if s else ""
            c = classify(label)
            if c and summary_office and nums:
                summary[(summary_office, c[1], c[0])] = summary.get((summary_office, c[1], c[0]), 0) + abs(nums[-1])
            continue
        first = s.split()[0] if s.split() else ""
        if office and PAT_ID_RE.match(first) and nums:
            # The description sits in its own column and wraps onto the next lines ("PMT PAT -" / "Alphaeon",
            # "ADJUST" / "ALPHAEON" / "FEE"). Read only that column, so the wrapped patient name is never read.
            amt_col = re.search(r"\s-?\(?[\d,]*\d\.\d\d", line).start()
            start = re.search(r"\b(pmt\s*pat|pymt|pt\s*pymt|adj(ust)?)\b", line[:amt_col], re.I)
            if not start:
                continue
            col, desc = start.start(), line[start.start():amt_col].strip()
            for nxt in lines[i + 1:i + 4]:
                if not nxt.strip() or re.search(r"\d\.\d\d", nxt) or PAT_ID_RE.match(nxt.split()[0]):
                    break
                desc += " " + nxt[max(0, col - 1):amt_col].strip()
            c = classify(re.sub(r"\s+", " ", desc))
            if c:
                detail.setdefault(office, []).append((first, c[1], c[0], next((abs(n) for n in nums if n), 0.0)))
    return _result(detail, summary, period, patient_type, transactions)


def _result(detail, summary, period, patient_type, transactions):
    if patient_type and patient_type.lower() == "ortho":
        raise JournalError("This Daily Journal is filtered to Patient Type: Ortho (that's for ortho collection)")
    if "adjust" not in (transactions or "").lower():
        raise JournalError("Run the Daily Journal with Transactions = Payments, Adjustments, so the lenders' "
                           "merchant fees are included")
    offices = []
    for office in sorted(set(detail) | {o for o, _, _ in summary}):
        lines = detail.get(office, [])
        by_lender = {}
        for _, lender, kind, amount in lines:
            by_lender.setdefault(lender, {"payment": 0.0, "fee": 0.0})[kind] += amount
        # Detail must add up to Denticon's own summary for this office, lender and line type.
        for (o, lender, kind), total in summary.items():
            if o != office:
                continue
            got = by_lender.get(lender, {}).get(kind, 0.0)
            if abs(got - total) > 0.01:
                raise JournalError(f"{office}: {lender} {kind}s per patient add up to {got:,.2f} but the report's "
                                   f"summary says {total:,.2f}; not using a report that doesn't add up")
        patients = {}
        for pat, lender, kind, amount in lines:
            p = patients.setdefault((pat, lender), {"pat_id": pat, "lender": lender, "payment": 0.0, "fee": 0.0})
            p[kind] += amount
        r2 = lambda v: round(v, 2)
        offices.append({
            "office": office,
            "prime_paid": r2(sum(v["payment"] + v["fee"] for k, v in by_lender.items() if k in PRIME)),
            "subprime_paid": r2(sum(v["payment"] + v["fee"] for k, v in by_lender.items() if k not in PRIME)),
            "fees": r2(sum(v["fee"] for v in by_lender.values())),
            "patients_funded": len({p["pat_id"] for p in patients.values()}),
            "by_lender": {k: {"payment": r2(v["payment"]), "fee": r2(v["fee"])} for k, v in sorted(by_lender.items())},
            "_patients": [{**p, "payment": r2(p["payment"]), "fee": r2(p["fee"])} for p in patients.values()],
        })
    if not any(o["by_lender"] for o in offices):
        raise JournalError("No financing lender payments found in this Daily Journal")
    return {"report_type": "lender_journal", "period": period, "offices": offices, "skipped": []}
