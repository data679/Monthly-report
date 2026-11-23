"""Procedures by patient ZIP code from a Denticon data dump, into the app's procedure_zips table.

    python3 tools/dump_procedure_zips.py [--dump data/denticon-dump/x] [--from 2025-11] [--to 2026-11]

The dump's LEDGER, PATIENT and Office tables must be unzipped in --dump. Every charge (not voided) is counted in the
month of its transaction date, at the office it was charted at, under the patient's current ZIP and the procedure's
category (the same CDT code ranges as Denticon's Production Analysis by Category). Only totals per office, month,
ZIP and category are saved, plus a "*" row per office, month and ZIP with all categories together; patient IDs are
used for counting and never stored. Months loaded replace what was there for them.
"""
import argparse
import collections
import csv
import glob
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import app  # noqa: E402

RANGES = ((100, 999, "DIAGNOSTIC"), (1000, 1999, "PREVENTIVE"), (2000, 2999, "RESTORATIVE"), (3000, 3999, "ENDODONTICS"),
          (4000, 4999, "PERIODONTICS"), (5000, 5899, "PROSTHODONTICS (REMOVABLE)"), (5900, 5999, "MAXILLOFACIAL PROSTHETICS"),
          (6000, 6199, "IMPLANT SERVICES"), (6200, 6999, "PROSTHODONTICS (FIXED)"), (7000, 7999, "ORAL AND MAXILOFACIAL SURGERY"),
          (8000, 8999, "ORTHODONTICS"), (9000, 9999, "ADJUNCT GENERAL SERVICES"))


def category(code):
    c = (code or "").strip().upper()
    if c[:1] == "D" and c[1:5].isdigit():
        n = int(c[1:5])
        for lo, hi, name in RANGES:
            if lo <= n <= hi:
                return name
    return "OTHER"


def rows(pattern):
    csv.field_size_limit(10 ** 9)
    for f in sorted(glob.glob(pattern)):
        with open(f, newline="", encoding="latin-1") as fh:
            for r in csv.DictReader(fh):
                if None not in r.values():
                    yield r


def month_of(s):   # "MM/DD/YYYY hh:mm:ss" -> "YYYY-MM"
    return f"{s[6:10]}-{s[:2]}" if len(s) >= 10 and s[2] == "/" else ""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dump", default=str(ROOT / "data" / "denticon-dump" / "x"))
    ap.add_argument("--from", dest="frm", default="2025-11")
    ap.add_argument("--to", default="9999-12")
    a = ap.parse_args()
    d = Path(a.dump)
    office = {r["OID"]: r["OFFICENAME"] for r in rows(str(d / "Office.txt"))}
    zips = {}
    for r in rows(str(d / "PATIENT" / "*.txt")):
        z = (r["ZIP"] or "").strip()[:5]
        if z.isdigit() and len(z) == 5:
            zips[r["PATID"]] = z
    prod = collections.defaultdict(float)
    count = collections.Counter()
    pats = collections.defaultdict(set)
    as_of = ""
    for r in rows(str(d / "LEDGER" / "*.txt")):
        c = r["CREATEDON"]
        if len(c) >= 10 and c[2] == "/":
            as_of = max(as_of, f"{c[6:10]}-{c[:2]}-{c[3:5]}")
        if r["LTYPE"] != "C" or r["ISVOID"] == "True":
            continue
        m = month_of(r["TRANDATE"])
        z = zips.get(r["PATID"])
        if not (a.frm <= m <= a.to) or not z or r["OID"] not in office:
            continue
        try:
            amount = float(r["AMOUNT"] or 0)
        except ValueError:
            continue
        for cat in (category(r["ADACODE"] or r["CODE"]), "*"):
            k = (office[r["OID"]], m, z, cat)
            prod[k] += amount
            count[k] += 1
            pats[k].add(r["PATID"])
    for k in [k for k in prod if k[1] > as_of[:7]]:   # charges dated after the dump (booked ahead), not done yet
        del prod[k]
    months = sorted({k[1] for k in prod})
    conn = app.db()
    with conn:
        conn.executemany("DELETE FROM procedure_zips WHERE month = ?", [(m,) for m in months])
        conn.executemany("INSERT INTO procedure_zips VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                         [(*k, round(v, 2), count[k], len(pats[k]), as_of) for k, v in prod.items()])
    conn.close()
    print(f"procedure_zips: {len(prod)} rows, months {months[0] if months else '-'} to {months[-1] if months else '-'}, dump as of {as_of}")


if __name__ == "__main__":
    main()
