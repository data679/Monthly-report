"""Denticon (Planet DDS) API client, standard library only.

Every request sends the subscription key in the `PDDS-Subscription-Key` header to
https://api.planetdds.com/denticon/... (see developer.planetdds.com). The key is kept in
data/denticon_api.json (git-ignored, readable only by this user) and is entered in the app's
Audit tab on the computer running the app; it is never printed or sent anywhere else.
"""

import json
import os
import re
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

API_BASE = os.environ.get("DENTICON_API", "https://api.planetdds.com/denticon")


class DenticonError(Exception):
    pass


class Client:
    def __init__(self, data_dir):
        self.key_file = Path(data_dir) / "denticon_api.json"

    def exists(self):
        return self.key_file.exists()

    def save_key(self, key):
        key = (key or "").strip()
        if not re.fullmatch(r"[A-Za-z0-9_\-]{16,128}", key):
            raise DenticonError("That doesn't look like an API subscription key (letters and numbers, no spaces).")
        self.key_file.parent.mkdir(parents=True, exist_ok=True)
        fd = os.open(self.key_file, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
        with os.fdopen(fd, "w") as f:
            json.dump({"subscription_key": key}, f)

    def remove_key(self):
        self.key_file.unlink(missing_ok=True)

    def _key(self):
        if not self.exists():
            raise DenticonError("No Denticon API key yet. Add it on the Audit tab.")
        return json.loads(self.key_file.read_text())["subscription_key"]

    def get(self, path, retries=4, **params):
        """GET one page: returns the parsed JSON. Waits and retries when Denticon says to slow down (429)."""
        query = urllib.parse.urlencode({k: v for k, v in params.items() if v is not None})
        req = urllib.request.Request(f"{API_BASE}{path}" + (f"?{query}" if query else ""),
                                     headers={"PDDS-Subscription-Key": self._key(), "Cache-Control": "no-cache",
                                              "Accept": "application/json"})
        try:
            with urllib.request.urlopen(req, timeout=120) as resp:
                return json.loads(resp.read() or b"{}")
        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", "replace")
            if e.code == 429 and retries:
                m = re.search(r"in (\d+) seconds", body)
                time.sleep(min(int(m.group(1)) + 1 if m else 30, 300))
                return self.get(path, retries - 1, **params)
            try:
                detail = json.loads(body)
                detail = detail.get("detail") or detail.get("title") or detail.get("message") or body
            except ValueError:
                detail = body or e.reason
            if e.code == 401:
                raise DenticonError("Denticon didn't accept the API key (401). Check the key on the Audit tab.")
            if e.code == 403:
                raise DenticonError(f"The API key isn't allowed to use this ({path}). {str(detail)[:200]}")
            raise DenticonError(f"Denticon returned {e.code} for {path}: {str(detail)[:200]}")
        except urllib.error.URLError as e:
            raise DenticonError(f"Couldn't reach Denticon: {e.reason}")

    def get_all(self, path, page_size=1000, max_pages=500, **params):
        """Every record from a paginated endpoint."""
        out, page = [], 1
        while page <= max_pages:
            data = self.get(path, PageNumber=page, PageSize=page_size, **params)
            out.extend(data.get("data") or [])
            if page >= (data.get("totalPages") or 1):
                break
            page += 1
        return out

    def offices(self):
        return self.get_all("/practices/v0/offices", page_size=1000)
