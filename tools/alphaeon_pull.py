"""Pull Alphaeon Credit's applications and sales reports for a date range and save the funnel, read-only.

    python3 tools/alphaeon_pull.py                         # this month so far, locations from the settings file
    python3 tools/alphaeon_pull.py --from 2026-10-01 --to 2026-10-31
    python3 tools/alphaeon_pull.py --locations 3935,1272   # other location ids
    python3 tools/alphaeon_pull.py --check                 # only get (or reuse) a token
    python3 tools/alphaeon_pull.py --new-token             # throw away the cached token first
    python3 tools/alphaeon_pull.py --demo                  # same flow against a local stand-in with sample data

The Client ID and Secret come from environment variables or data/alphaeon.env (git-ignored, owner-only), never from
this file, and are never printed; neither is the token. The token is cached in data/alphaeon_token.json until shortly
before it expires. Only report endpoints are called (POST reports/applications and reports/sales).

Saves to data/alphaeon.db (one row per application and per sale, without names, emails, phone numbers or account
numbers) and writes the funnel to data/alphaeon/Alphaeon funnel FROM to TO.csv, which opens as a spreadsheet.
Funnel: total applications, net (approved + declined), approved, declined, account exists, funded sales
(transaction_type sale with status settled). Exit codes: 0 ok, 2 Alphaeon refused the request, 3 settings missing.
"""

import argparse
import csv
import datetime as dt
import json
import os
import sqlite3
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
ENV_FILE = DATA / "alphaeon.env"
TOKEN_FILE = DATA / "alphaeon_token.json"
DB_FILE = DATA / "alphaeon.db"
OUT_DIR = DATA / "alphaeon"
DEFAULTS = {
    "ALPHAEON_TOKEN_URL": "https://login.alphaeontest.com/oauth/token",
    "ALPHAEON_AUDIENCE": "https://api.alphaeon.com/",
    "ALPHAEON_API": "https://api.sandbox.alphaeontest.com/merchant_management/v1",
    "ALPHAEON_LOCATIONS": "",
}


class Refused(Exception):
    """Alphaeon answered with an error; the message is safe to show on screen."""


def settings():
    s = dict(DEFAULTS)
    if ENV_FILE.exists():
        for line in ENV_FILE.read_text().splitlines():
            if "=" in line and not line.lstrip().startswith("#"):
                k, v = line.split("=", 1)
                s[k.strip()] = v.strip()
    s.update({k: v for k, v in os.environ.items() if k.startswith("ALPHAEON_")})
    missing = [k for k in ("ALPHAEON_CLIENT_ID", "ALPHAEON_CLIENT_SECRET") if not s.get(k)]
    if missing:
        print(f"Missing {', '.join(missing)}: set them as environment variables or in {ENV_FILE.relative_to(ROOT)}.")
        sys.exit(3)
    return s


def post(url, body, token=None):
    headers = {"content-type": "application/json", "accept": "application/json"}
    if token:
        headers["authorization"] = "Bearer " + token
    req = urllib.request.Request(url, json.dumps(body).encode(), headers, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=120) as r:
            return r.status, json.load(r)
    except urllib.error.HTTPError as e:
        raw = e.read().decode("utf-8", "replace")
        try:
            detail = json.loads(raw)
        except ValueError:
            detail = {"error": raw[:300]}
        return e.code, detail
    except (urllib.error.URLError, TimeoutError) as e:
        raise Refused(f"could not reach {url.split('/')[2]}: {getattr(e, 'reason', e)}")


def error_text(detail, s):
    if isinstance(detail, dict):
        msg = detail.get("error_description") or detail.get("message") or detail.get("errors") or detail.get("error")
    else:
        msg = detail
    msg = msg if isinstance(msg, str) else json.dumps(msg)
    return msg.replace(s["ALPHAEON_CLIENT_ID"], "<client id>").replace(s["ALPHAEON_CLIENT_SECRET"], "<secret>")


def token(s, fresh=False):
    """A bearer token: the cached one while it has more than 5 minutes left, otherwise a new one."""
    if not fresh and TOKEN_FILE.exists():
        try:
            t = json.loads(TOKEN_FILE.read_text())
            if t.get("client") == s["ALPHAEON_CLIENT_ID"][:6] and t["expires_at"] - time.time() > 300:
                return t["access_token"], f"reused, {(t['expires_at'] - time.time()) / 3600:.1f} h left"
        except (ValueError, KeyError):
            pass
    status, d = post(s["ALPHAEON_TOKEN_URL"], {
        "client_id": s["ALPHAEON_CLIENT_ID"], "client_secret": s["ALPHAEON_CLIENT_SECRET"],
        "audience": s["ALPHAEON_AUDIENCE"], "grant_type": "client_credentials"})
    if status != 200 or "access_token" not in d:
        raise Refused(f"token request refused ({status}): {error_text(d, s)}")
    expires_at = time.time() + int(d.get("expires_in", 86400))
    TOKEN_FILE.touch(mode=0o600)
    TOKEN_FILE.write_text(json.dumps({"access_token": d["access_token"], "expires_at": expires_at,
                                      "client": s["ALPHAEON_CLIENT_ID"][:6]}))
    return d["access_token"], f"new, valid {int(d.get('expires_in', 86400)) / 3600:.0f} h"


def report(s, tok, name, body):
    status, d = post(f"{s['ALPHAEON_API']}/reports/{name}", body, tok)
    if status not in (200, 201):
        raise Refused(f"reports/{name} refused ({status}): {error_text(d, s)}")
    return d


def save(frm, to, locations, apps, sales):
    c = sqlite3.connect(DB_FILE)
    c.executescript("""
        CREATE TABLE IF NOT EXISTS alphaeon_pulls (id INTEGER PRIMARY KEY, run_at TEXT, start_date TEXT,
            end_date TEXT, location_ids TEXT, applications INTEGER, sales INTEGER);
        CREATE TABLE IF NOT EXISTS alphaeon_applications (uuid TEXT PRIMARY KEY, location_id INTEGER,
            location_name TEXT, created_at TEXT, application_status TEXT, prequal_status TEXT, product_type TEXT,
            program_product_type TEXT, lender_name TEXT, application_source TEXT, requested_amount REAL,
            credit_limit REAL, pull_id INTEGER);
        CREATE TABLE IF NOT EXISTS alphaeon_sales (id INTEGER PRIMARY KEY, location_id INTEGER, location_name TEXT,
            created_at TEXT, funding_date TEXT, transaction_type TEXT, status TEXT, amount REAL, plan_name TEXT,
            program_product_type TEXT, merchant_discount_rate REAL, pull_id INTEGER);
    """)
    pull = c.execute("INSERT INTO alphaeon_pulls (run_at, start_date, end_date, location_ids, applications, sales) "
                     "VALUES (?,?,?,?,?,?)", (dt.datetime.now().isoformat(timespec="seconds"), frm, to,
                                              ",".join(map(str, locations)), len(apps), len(sales))).lastrowid
    c.executemany("INSERT OR REPLACE INTO alphaeon_applications VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", [
        (a.get("uuid") or a.get("application_identifier"), a.get("location_id"), a.get("location_name"),
         a.get("created_at"), a.get("application_status"), a.get("prequal_status"), a.get("product_type"),
         a.get("program_product_type"), a.get("lender_name"), a.get("application_source"),
         a.get("requested_amount"), a.get("credit_limit"), pull) for a in apps])
    c.executemany("INSERT OR REPLACE INTO alphaeon_sales VALUES (?,?,?,?,?,?,?,?,?,?,?,?)", [
        (t.get("id"), t.get("location_id"), t.get("location_name"), t.get("created_at"), t.get("funding_date"),
         t.get("transaction_type"), t.get("status"), t.get("amount"), t.get("plan_name"),
         t.get("program_product_type"), t.get("merchant_discount_rate"), pull) for t in sales])
    c.commit()
    c.close()
    return pull


def funnel(apps, sales):
    """Funnel rows per location plus a total row."""
    rows = {}

    def row(loc_id, name):
        return rows.setdefault(loc_id, {"location_id": loc_id, "location": name or "", "total": 0, "approved": 0,
                                        "declined": 0, "account_exists": 0, "pending": 0,
                                        "funded_sales": 0, "funded_amount": 0.0})
    for a in apps:
        r = row(a.get("location_id"), a.get("location_name"))
        r["total"] += 1
        st = (a.get("application_status") or "pending").strip().lower().replace(" ", "_")
        if st in ("approved", "declined", "account_exists"):
            r[st] += 1
        else:
            r["pending"] += 1
    for t in sales:
        if (t.get("transaction_type") or "").lower() == "sale" and (t.get("status") or "").lower() == "settled":
            r = row(t.get("location_id"), t.get("location_name"))
            r["funded_sales"] += 1
            r["funded_amount"] += float(t.get("amount") or 0)
    out = sorted(rows.values(), key=lambda r: str(r["location"]))
    total = {k: (sum(r[k] for r in out) if isinstance(v, (int, float)) else "") for k, v in
             (out[0] if out else row(None, "")).items()}
    total.update(location_id="", location="All locations")
    for r in out + [total]:
        r["net"] = r["approved"] + r["declined"]
        r["approval_rate"] = r["approved"] / r["net"] if r["net"] else None
    return out, total


COLUMNS = [("location", "Location"), ("location_id", "ID"), ("total", "Applications"), ("net", "Net (appr+decl)"),
           ("approved", "Approved"), ("declined", "Declined"), ("account_exists", "Account exists"),
           ("pending", "Pending/other"), ("approval_rate", "Approval rate"), ("funded_sales", "Funded sales"),
           ("funded_amount", "Funded $")]


def fmt(k, v):
    if v is None or v == "":
        return "" if k != "approval_rate" else "-"
    if k == "approval_rate":
        return f"{v:.0%}"
    if k == "funded_amount":
        return f"${v:,.2f}"
    if k == "location_id":
        return str(v)
    return f"{v:,}" if isinstance(v, int) else str(v)


def write_csv(frm, to, rows, total):
    OUT_DIR.mkdir(exist_ok=True)
    path = OUT_DIR / f"Alphaeon funnel {frm} to {to}.csv"
    with open(path, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow([h for _, h in COLUMNS])
        for r in rows + [total]:
            w.writerow([round(r[k], 4) if k == "approval_rate" and r[k] is not None
                        else round(r[k], 2) if k == "funded_amount" else r[k] for k, _ in COLUMNS])
    return path


def print_table(rows, total):
    cells = [[h for _, h in COLUMNS]] + [[fmt(k, r[k]) for k, _ in COLUMNS] for r in rows + [total]]
    widths = [min(max(len(c[i]) for c in cells), 34) for i in range(len(COLUMNS))]
    for n, c in enumerate(cells):
        print("  ".join((v[:34].ljust(w) if i == 0 else v.rjust(w)) for i, (v, w) in enumerate(zip(c, widths))))
        if n == 0 or n == len(cells) - 2:
            print("  ".join("-" * w for w in widths))


def demo():
    """Point everything at the local stand-in, with its own files, so sample data never mixes with real pulls."""
    global TOKEN_FILE, DB_FILE, OUT_DIR
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import alphaeon_demo
    url = alphaeon_demo.start()
    os.environ.update(ALPHAEON_CLIENT_ID="demo-client", ALPHAEON_CLIENT_SECRET="demo-secret",
                      ALPHAEON_TOKEN_URL=url + "/oauth/token", ALPHAEON_API=url + "/v1",
                      ALPHAEON_LOCATIONS="3935")
    TOKEN_FILE, DB_FILE, OUT_DIR = DATA / "alphaeon_demo_token.json", DATA / "alphaeon_demo.db", DATA / "alphaeon_demo"
    print("*** DEMO MODE: local stand-in for the Alphaeon sandbox, sample data only, not real Alphaeon data ***")


def main():
    today = dt.date.today()
    p = argparse.ArgumentParser(description="Pull Alphaeon applications and sales reports (read-only).")
    p.add_argument("--from", dest="frm", default=today.replace(day=1).isoformat(), help="YYYY-MM-DD")
    p.add_argument("--to", default=today.isoformat(), help="YYYY-MM-DD")
    p.add_argument("--locations", help="comma-separated location ids (default: ALPHAEON_LOCATIONS)")
    p.add_argument("--check", action="store_true", help="only get or reuse a token")
    p.add_argument("--new-token", action="store_true", help="ignore the cached token")
    p.add_argument("--demo", action="store_true", help="use a local stand-in for Alphaeon with sample data")
    a = p.parse_args()
    if a.demo:
        demo()
    s = settings()
    try:
        tok, how = token(s, a.new_token)
        print(f"Token: {how}.")
        if a.check:
            return 0
        locations = [int(x) for x in (a.locations or s["ALPHAEON_LOCATIONS"]).split(",") if x.strip()]
        if not locations:
            print("No location ids: pass --locations or set ALPHAEON_LOCATIONS.")
            return 3
        body = {"location_ids": locations, "start_date": a.frm, "end_date": a.to}
        print(f"Pulling {a.frm} to {a.to}, location(s) {', '.join(map(str, locations))} ...")
        try:
            apps_r = report(s, tok, "applications", body)
        except Refused as e:
            if "(401)" not in str(e):
                raise
            tok, how = token(s, fresh=True)   # the cached token was revoked or expired early
            print(f"Token: {how}.")
            apps_r = report(s, tok, "applications", body)
        sales_r = report(s, tok, "sales", body)
    except Refused as e:
        print(f"Alphaeon refused: {e}")
        print("Nothing was saved.")
        return 2
    apps, sales = apps_r.get("applications") or [], sales_r.get("transactions") or []
    pull = save(a.frm, a.to, locations, apps, sales)
    rows, total = funnel(apps, sales)
    print(f"{len(apps):,} applications and {len(sales):,} sales transactions saved (pull #{pull}, "
          f"{DB_FILE.relative_to(ROOT)}).\n")
    print_table(rows, total)
    summ = apps_r.get("summary") or {}
    if summ.get("total_applications") is not None and summ["total_applications"] != total["total"]:
        print(f"\nNote: Alphaeon's summary says {summ['total_applications']:,} applications; the rows add up to "
              f"{total['total']:,}.")
    print(f"\nSpreadsheet: {write_csv(a.frm, a.to, rows, total).relative_to(ROOT)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
