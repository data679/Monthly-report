"""Read Google Sheets with a service account (read-only), using only the standard library.

The service account's key is kept in data/google/ (git-ignored). Each office shares its
sheet with the service account's email as Viewer; the app then reads the month tabs the
same way it reads an uploaded workbook (see np_analysis.Workbook).

Sign-in uses a JWT signed with the key via the `openssl` command line tool.
"""

import base64
import json
import os
import re
import subprocess
import tempfile
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

SCOPE = "https://www.googleapis.com/auth/spreadsheets.readonly"
# Overridable so the sync can be tested against a local stand-in for Google.
TOKEN_URL = os.environ.get("GOOGLE_TOKEN_URL", "https://oauth2.googleapis.com/token")
API_BASE = os.environ.get("GOOGLE_SHEETS_API", "https://sheets.googleapis.com/v4")


class GoogleError(Exception):
    pass


def _b64url(data):
    return base64.urlsafe_b64encode(data).rstrip(b"=").decode()


def spreadsheet_id(url_or_id):
    """The spreadsheet ID from a Google Sheets link (or the ID itself)."""
    m = re.search(r"/spreadsheets/d/([A-Za-z0-9_-]{20,})", url_or_id or "")
    if m:
        return m.group(1)
    if re.fullmatch(r"[A-Za-z0-9_-]{20,}", (url_or_id or "").strip()):
        return url_or_id.strip()
    raise GoogleError("That doesn't look like a Google Sheets link (…/spreadsheets/d/…)")


class ServiceAccount:
    def __init__(self, key_dir):
        self.key_dir = Path(key_dir)
        self.key_file = self.key_dir / "service_account.json"
        self._token, self._expires = None, 0

    # ------------------------------------------------------------------ key file
    def exists(self):
        return self.key_file.exists()

    def email(self):
        return json.loads(self.key_file.read_text())["client_email"] if self.exists() else None

    def save_key(self, raw):
        try:
            info = json.loads(raw)
        except ValueError:
            raise GoogleError("The key file isn't valid JSON. Download the service account key as JSON.")
        if info.get("type") != "service_account" or not info.get("private_key") or not info.get("client_email"):
            raise GoogleError("That isn't a service account key (it needs type, client_email and private_key).")
        self.key_dir.mkdir(parents=True, exist_ok=True)
        os.chmod(self.key_dir, 0o700)
        self.key_file.write_text(json.dumps(info))
        os.chmod(self.key_file, 0o600)
        self._token = None
        return info["client_email"]

    # ------------------------------------------------------------------ sign-in
    def token(self):
        if self._token and time.time() < self._expires - 60:
            return self._token
        if not self.exists():
            raise GoogleError("No Google service account key yet. Add it on the Google Sheets tab.")
        info = json.loads(self.key_file.read_text())
        now = int(time.time())
        header = _b64url(json.dumps({"alg": "RS256", "typ": "JWT"}).encode())
        claims = _b64url(json.dumps({"iss": info["client_email"], "scope": SCOPE, "aud": TOKEN_URL,
                                     "iat": now, "exp": now + 3600}).encode())
        signing_input = f"{header}.{claims}".encode()
        # openssl needs the key in a file; keep it private and remove it straight away.
        fd, pem = tempfile.mkstemp(dir=self.key_dir, suffix=".pem")
        try:
            with os.fdopen(fd, "w") as f:
                f.write(info["private_key"])
            sig = subprocess.run(["openssl", "dgst", "-sha256", "-sign", pem], input=signing_input,
                                 capture_output=True, check=True).stdout
        except FileNotFoundError:
            raise GoogleError("openssl isn't installed; it's needed to sign in to Google.")
        except subprocess.CalledProcessError as e:
            raise GoogleError(f"Couldn't sign with the key: {e.stderr.decode().strip()}")
        finally:
            os.unlink(pem)
        body = urllib.parse.urlencode({"grant_type": "urn:ietf:params:oauth:grant-type:jwt-bearer",
                                       "assertion": f"{header}.{claims}.{_b64url(sig)}"}).encode()
        data = _request(urllib.request.Request(TOKEN_URL, data=body))
        self._token, self._expires = data["access_token"], time.time() + int(data.get("expires_in", 3600))
        return self._token


def _request(req):
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            return json.loads(resp.read())
    except urllib.error.HTTPError as e:
        try:
            message = json.loads(e.read()).get("error", {})
            message = message.get("message") if isinstance(message, dict) else message
        except ValueError:
            message = e.reason
        if e.code == 400 and "not supported for this document" in str(message):
            raise GoogleError("That link is an Excel file stored in Google Drive, not a Google Sheet. Open it and use "
                              "File → Save as Google Sheets, then share and add the new sheet's link.")
        if e.code == 403:
            raise GoogleError(f"No access. Share the sheet with the service account as Viewer. ({message})")
        if e.code == 404:
            raise GoogleError("Sheet not found. Check the link, and that it's shared with the service account.")
        raise GoogleError(f"Google returned {e.code}: {message}")
    except urllib.error.URLError as e:
        raise GoogleError(f"Couldn't reach Google: {e.reason}")


def _col_letters(i):
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


class SheetsWorkbook:
    """A Google Sheet, read like an uploaded workbook: .title, .sheets, .rows(sheet)."""

    def __init__(self, account, sheet_id):
        self.account, self.id = account, sheet_id
        meta = self._get(f"/spreadsheets/{sheet_id}", fields="properties.title,sheets.properties.title")
        self.title = meta["properties"]["title"]
        self.sheets = {s["properties"]["title"]: s["properties"]["title"] for s in meta.get("sheets", [])}

    def _get(self, path, **params):
        url = API_BASE + urllib.parse.quote(path, safe="/!:',") + "?" + urllib.parse.urlencode(params)
        return _request(urllib.request.Request(url, headers={"Authorization": f"Bearer {self.account.token()}"}))

    def rows(self, sheet):
        """{row number: {column letters: text}}, the same shape as np_analysis.Workbook.rows."""
        title = "'" + sheet.replace("'", "''") + "'"
        data = self._get(f"/spreadsheets/{self.id}/values/{title}",
                         valueRenderOption="UNFORMATTED_VALUE", dateTimeRenderOption="SERIAL_NUMBER")
        out = {}
        for r, values in enumerate(data.get("values", []), start=1):
            cells = {}
            for c, v in enumerate(values):
                if isinstance(v, bool):
                    v = "TRUE" if v else "FALSE"
                text = repr(float(v)) if isinstance(v, float) else str(v)
                if text.strip():
                    cells[_col_letters(c)] = text.strip()
            if cells:
                out[r] = cells
        return out
