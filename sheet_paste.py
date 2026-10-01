"""Read rows pasted from a spreadsheet (tab-separated) and match office names.

The sheet's office names ("TORRANCE", "el segundo") differ from the names in
Denticon reports ("Dentist Of Torrance"), so each pasted office gets a list of
likely matches; the user confirms any that aren't clear-cut and the choice is
remembered as an alias.
"""

import re
from datetime import datetime

# Columns we pull from a pasted sheet: metric name -> header pattern.
# Add an entry here to start collecting another column.
METRICS = {
    "refunds": re.compile(r"^refunds?$", re.I),
}
OFFICE_HEADER = re.compile(r"^office$", re.I)
UPDATED_RE = re.compile(r"updated\s*:?\s*(\d{1,2}/\d{1,2}/\d{2,4})", re.I)


class PasteError(Exception):
    pass


def parse_money(text):
    t = (text or "").strip().replace("$", "").replace(",", "")
    if not t or t in ("-", "—"):
        return None
    negative = t.startswith("(") and t.endswith(")")
    t = t.strip("()")
    try:
        value = float(t)
    except ValueError:
        return None
    return -value if negative else value


def parse_date(text):
    for fmt in ("%m/%d/%Y", "%m/%d/%y"):
        try:
            return datetime.strptime(text, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return None


def normalize(name):
    return " ".join(re.findall(r"[a-z0-9]+", name.lower()))


def parse_paste(text):
    """Return {"as_of": "YYYY-MM-DD" | None, "metrics": [...], "rows": [{sheet_office, values}]}."""
    lines = [line.split("\t") for line in text.replace("\r\n", "\n").split("\n")]

    m = UPDATED_RE.search(text)
    as_of = parse_date(m.group(1)) if m else None

    header_i, office_col, metric_cols = None, None, {}
    for i, cells in enumerate(lines):
        cells = [c.strip() for c in cells]
        oc = next((j for j, c in enumerate(cells) if OFFICE_HEADER.match(c)), None)
        found = {name: j for name, pat in METRICS.items()
                 for j, c in enumerate(cells) if pat.match(c)}
        if oc is not None and found:
            header_i, office_col, metric_cols = i, oc, found
            break
    if header_i is None:
        wanted = ", ".join(METRICS)
        raise PasteError(
            f"Couldn't find the header row. Copy the rows including the header that has "
            f"'Office' and {wanted!r} columns (paste straight from the sheet so columns stay tab-separated)."
        )

    rows = []
    for cells in lines[header_i + 1:]:
        cells = [c.strip() for c in cells]
        office = cells[office_col] if office_col < len(cells) else ""
        if not office or normalize(office) in ("total", "totals", "grand total"):
            continue
        values = {name: parse_money(cells[j]) if j < len(cells) else None
                  for name, j in metric_cols.items()}
        if all(v is None for v in values.values()):
            continue
        rows.append({"sheet_office": office, "values": values})
    if not rows:
        raise PasteError("Found the header row but no office rows under it")
    return {"as_of": as_of, "metrics": list(metric_cols), "rows": rows}


def candidates(sheet_office, known_offices):
    """Known offices whose name contains every word of the sheet name."""
    words = normalize(sheet_office).split()
    out = []
    for office in known_offices:
        office_words = normalize(office).split()
        if all(w in office_words for w in words):
            out.append(office)
    return out


def title_case(name):
    return " ".join(w.capitalize() if not w.isdigit() else w for w in name.split())
