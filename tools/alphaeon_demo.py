"""A stand-in for the Alphaeon sandbox, for `alphaeon_pull.py --demo` while the real sandbox refuses our token.

It runs on 127.0.0.1 inside the same process and answers the token endpoint and POST reports/applications and
reports/sales with the response shapes from Alphaeon's Swagger docs. The data is made up (seeded by the date range, so
the same range gives the same numbers). Location 3935 (Gardena in the sandbox) is allowed; any other location gets the
422 the real API returns for a location the user has no access to.
"""

import datetime as dt
import json
import random
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer

ALLOWED = {3935: "Gardena"}
TOKEN = "demo-token"
STATUSES = ["approved"] * 46 + ["declined"] * 31 + ["account_exists"] * 9 + ["pending"] * 6 + [None] * 8


def _days(frm, to):
    a, b = dt.date.fromisoformat(frm), dt.date.fromisoformat(to)
    return [a + dt.timedelta(n) for n in range(max((b - a).days + 1, 1))]


def _applications(loc, frm, to):
    rng = random.Random(f"apps {loc} {frm} {to}")
    out = []
    for day in _days(frm, to):
        for n in range(rng.choice([0, 1, 1, 2, 2, 3, 4]) if day.weekday() < 6 else 0):
            st = rng.choice(STATUSES)
            out.append({"location_id": loc, "location_name": ALLOWED[loc], "uuid": f"demo-{day:%Y%m%d}-{n}",
                        "created_at": f"{day}T{rng.randint(15, 23):02d}:{rng.randint(0, 59):02d}:00Z",
                        "product_type": "line_of_credit", "application_status": st,
                        "requested_amount": rng.choice([1500, 2500, 4000, 6500, 9000]),
                        "credit_limit": rng.choice([2000, 3500, 5000, 8000]) if st in ("approved", "account_exists")
                        else None, "lender_name": "Alphaeon", "application_source": rng.choice(["in_office", "link"]),
                        "customer_name": "Sample Patient", "email": "sample@example.com",
                        "primary_phone": "555-0100"})
    return out


def _sales(loc, frm, to):
    rng = random.Random(f"sales {loc} {frm} {to}")
    out, n = [], 0
    for day in _days(frm, to):
        for _ in range(rng.choice([0, 0, 1, 1, 2]) if day.weekday() < 6 else 0):
            n += 1
            kind = rng.choice(["sale"] * 9 + ["refund"])
            out.append({"id": n, "location_id": loc, "location_name": ALLOWED[loc],
                        "created_at": f"{day}T18:00:00Z", "funding_date": str(day + dt.timedelta(2)),
                        "transaction_type": kind, "amount": round(rng.uniform(400, 7000), 2),
                        "status": rng.choice(["settled"] * 8 + ["pending_signature", "voided"]),
                        "plan_name": rng.choice(["12 Months Deferred Interest", "60 Months Low APR, Equal Payment"]),
                        "program_product_type": "ac_card", "merchant_discount_rate": rng.choice([6.5, 9.9, 12.5]),
                        "name": {"first_name": "Sample", "middle_initial": "", "last_name": "Patient"}})
    return out


class _Handler(BaseHTTPRequestHandler):
    def log_message(self, *args):
        pass

    def _send(self, code, data):
        body = json.dumps(data).encode()
        self.send_response(code)
        self.send_header("content-type", "application/json")
        self.send_header("content-length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self):
        body = json.loads(self.rfile.read(int(self.headers.get("content-length") or 0)) or b"{}")
        if self.path.endswith("/oauth/token"):
            return self._send(200, {"access_token": TOKEN, "expires_in": 86400, "token_type": "Bearer"})
        if self.headers.get("authorization") != f"Bearer {TOKEN}":
            return self._send(401, {"error": "Unauthorized"})
        locs = body.get("location_ids") or []
        bad = [x for x in locs if x not in ALLOWED]
        if bad:
            return self._send(422, {"error": f"Not authorized to view location {bad[0]}"})
        frm, to = body["start_date"], body["end_date"]
        if self.path.endswith("/reports/applications"):
            apps = [a for loc in locs for a in _applications(loc, frm, to)]
            return self._send(201, {"summary": {"total_applications": len(apps)}, "applications": apps})
        if self.path.endswith("/reports/sales"):
            return self._send(201, {"summary": {}, "transactions": [t for loc in locs for t in _sales(loc, frm, to)]})
        self._send(404, {"error": "not found"})


def start():
    """Start the stand-in on a free local port; returns its base URL."""
    server = HTTPServer(("127.0.0.1", 0), _Handler)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    return f"http://127.0.0.1:{server.server_port}"
