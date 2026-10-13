"""Invoca (call tracking) API client, standard library only.

Transactions API (developers.invoca.net, version 2020-10-01):
  https://<site>.invoca.net/api/2020-10-01/<networks|advertisers>/transactions/<id>.json
Every request sends the API token in the `Authorization` header (Invoca's recommended way; it also
accepts ?oauth_token=, which we never use so the token stays out of URLs and logs). The site name,
account type, id and token live in data/invoca_api.json (git-ignored, readable only by this user);
they are entered on the app's Audit tab on the computer running the app and never printed.
Calls are read with a date range and paged with start_after_transaction_id, as the docs describe.
The caller's number is fingerprinted by the app in memory and never stored.
"""

import json
import os
import re
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

API_VERSION = "2020-10-01"
# The fields we ask for: enough to count and time calls per office/campaign. The caller's number is asked for
# only so the app can turn it into a keyed fingerprint in memory (to match new-patient appointments); it is
# never stored or shown.
COLUMNS = ["transaction_id", "calling_phone_number", "transaction_type", "complete_call_id", "start_time_local", "start_time_utc", "duration",
           "connect_duration", "advertiser_id", "advertiser_name", "advertiser_campaign_id", "advertiser_campaign_name",
           "virtual_line_id", "promo_line_description", "city", "region", "mobile", "repeat_calling_phone_number",
           "call_result_description_detail", "custom_data"]


class InvocaError(Exception):
    pass


class Client:
    def __init__(self, data_dir):
        self.file = Path(data_dir) / "invoca_api.json"

    def exists(self):
        return self.file.exists()

    def config(self):
        """Site, account type and id (never the token)."""
        if not self.exists():
            return None
        c = json.loads(self.file.read_text())
        return {"site": c.get("site"), "kind": c.get("kind"), "id": c.get("id")}

    def save(self, site, kind, id_, token):
        site = re.sub(r"^https?://", "", (site or "").strip().lower()).split("/")[0].replace(".invoca.net", "")
        if not re.fullmatch(r"[a-z0-9][a-z0-9-]{0,62}", site):
            raise InvocaError("The site name is the first part of your Invoca address: for https://acme.invoca.net it is 'acme'.")
        if kind not in ("networks", "advertisers"):
            raise InvocaError("Account type must be Network or Advertiser.")
        id_ = str(id_ or "").strip()
        if not id_.isdigit():
            raise InvocaError("The ID is the number after /networks/ (or /advertisers/) in your Invoca address, e.g. 99.")
        token = (token or "").strip()
        if not re.fullmatch(r"[A-Za-z0-9_\-]{8,256}", token):
            raise InvocaError("That doesn't look like an API token (letters and numbers, no spaces).")
        self.file.parent.mkdir(parents=True, exist_ok=True)
        fd = os.open(self.file, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
        with os.fdopen(fd, "w") as f:
            json.dump({"site": site, "kind": kind, "id": id_, "token": token}, f)

    def remove(self):
        self.file.unlink(missing_ok=True)

    def _cfg(self):
        if not self.exists():
            raise InvocaError("Invoca isn't connected yet. Add the site, ID and API token on the Audit tab.")
        return json.loads(self.file.read_text())

    def transactions(self, frm, to, limit=1000, start_after=None, columns=None, retries=3):
        """One page of transactions (list of dicts) for the date range (YYYY-MM-DD, inclusive, the user's timezone)."""
        c = self._cfg()
        params = {"from": frm, "to": to, "limit": min(int(limit), 4000), "include_columns": ",".join(columns or COLUMNS)}
        if start_after:
            params["start_after_transaction_id"] = start_after
        url = f"https://{c['site']}.invoca.net/api/{API_VERSION}/{c['kind']}/transactions/{c['id']}.json?" + urllib.parse.urlencode(params)
        req = urllib.request.Request(url, headers={"Authorization": c["token"], "Accept": "application/json"})
        try:
            with urllib.request.urlopen(req, timeout=120) as resp:
                data = json.loads(resp.read() or b"[]")
        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", "replace")[:300]
            if e.code == 429 and retries:
                time.sleep(15)
                return self.transactions(frm, to, limit, start_after, columns, retries - 1)
            if e.code in (401, 403):
                raise InvocaError("Invoca didn't accept the API token (HTTP %d). Check the token, and that it was made on the %s account with ID %s."
                                  % (e.code, "network" if c["kind"] == "networks" else "advertiser", c["id"]))
            if e.code == 404:
                raise InvocaError(f"Invoca found nothing at {c['site']}.invoca.net for {c['kind'][:-1]} ID {c['id']} (HTTP 404). Check the site name, account type and ID.")
            raise InvocaError(f"Invoca returned HTTP {e.code}: {body or e.reason}")
        except urllib.error.URLError as e:
            raise InvocaError(f"Couldn't reach {c['site']}.invoca.net: {e.reason}")
        if isinstance(data, dict):
            data = data.get("transactions") or data.get("data") or []
        return data

    def all_transactions(self, frm, to, columns=None, max_pages=200):
        """Every transaction in the range, following start_after_transaction_id until a page comes back empty."""
        out, last = [], None
        for _ in range(max_pages):
            page = self.transactions(frm, to, 4000, last, columns)
            if not page:
                break
            out.extend(page)
            last = page[-1].get("transaction_id")
            if len(page) < 4000 or not last:
                break
        return out
