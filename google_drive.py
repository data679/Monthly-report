"""Read Google Drive for the chat assistant, read-only, using only the standard library.

It signs in with the same service account as google_sheets.py, so it sees only the files and folders someone has shared
with that account's email (as Viewer); nothing else in anyone's Drive. Google Docs, Sheets and Slides are read; other
file types are listed but not opened. Needs the Google Drive API turned on in the service account's Cloud project.
"""

import json
import os
import time
import urllib.error
import urllib.parse
import urllib.request

import assistant
from google_sheets import GoogleError

SCOPE = "https://www.googleapis.com/auth/drive.readonly"
API = os.environ.get("GOOGLE_DRIVE_API", "https://www.googleapis.com/drive/v3")
DOC, SHEET, SLIDES, FOLDER = ("application/vnd.google-apps." + k for k in ("document", "spreadsheet", "presentation", "folder"))
READABLE = {DOC: "Google Doc", SHEET: "Google Sheet", SLIDES: "Google Slides"}
KINDS = {**READABLE, FOLDER: "folder"}
FIELDS = "files(id,name,mimeType,modifiedTime,webViewLink,owners(displayName),parents)"
READ_CHUNK = 40_000     # characters per drive_read call; longer files are read in parts (offset)
_cache = {}             # file id -> (modifiedTime, text): a file is only exported again after it changes
_CACHE_MAX = 40


def _get(account, path, raw=False, **params):
    url = f"{API}/{path}?" + urllib.parse.urlencode({**params, "supportsAllDrives": "true"})
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {account.token(SCOPE)}"})
    for attempt in range(3):
        try:
            with urllib.request.urlopen(req, timeout=60) as r:
                body = r.read()
                return body if raw else json.loads(body)
        except urllib.error.HTTPError as e:
            if e.code in (429, 500, 503) and attempt < 2:
                time.sleep(5 * (attempt + 1))
                continue
            try:
                err = json.loads(e.read()).get("error", {})
            except ValueError:
                err = {}
            message = err.get("message", e.reason) if isinstance(err, dict) else str(err)
            reasons = {x.get("reason") for x in err.get("errors", [])} if isinstance(err, dict) else set()
            if "accessNotConfigured" in reasons or "has not been used in project" in message or "is disabled" in message:
                raise GoogleError("The Google Drive API is turned off for the app's Google Cloud project. Turn it on in the "
                                  "Google Cloud console (APIs & Services → Library → Google Drive API → Enable), then try again.")
            if e.code == 404:
                raise GoogleError("That Drive file wasn't found, or it isn't shared with the app's Google account.")
            if e.code == 403 and "exportSizeLimitExceeded" in reasons:
                raise GoogleError("That file is too large for Google to export (10 MB limit).")
            if e.code == 403:
                raise GoogleError(f"No access to that Drive file. Share it with the app's Google account as Viewer. ({message})")
            raise GoogleError(f"Google Drive returned {e.code}: {message}")
        except urllib.error.URLError as e:
            raise GoogleError(f"Couldn't reach Google Drive: {e.reason}")


def _brief(f):
    return {"id": f["id"], "name": f["name"], "type": KINDS.get(f["mimeType"], f["mimeType"].split("/")[-1].replace("vnd.google-apps.", "")),
            "modified": (f.get("modifiedTime") or "")[:10], "link": f.get("webViewLink"),
            "owner": ((f.get("owners") or [{}])[0]).get("displayName"), "readable": f["mimeType"] in READABLE}


def _q(text):
    return text.replace("\\", "\\\\").replace("'", "\\'")


def search(account, query="", folder_id=None, limit=15):
    """Files shared with the app whose name or text matches the query (newest first); or a folder's contents;
    or, with neither, the most recently changed files."""
    terms = ["trashed = false"]
    if folder_id:
        terms.append(f"'{_q(folder_id)}' in parents")
    if query:
        terms.append(f"(name contains '{_q(query)}' or fullText contains '{_q(query)}')")
    params = {"q": " and ".join(terms), "fields": FIELDS, "pageSize": max(1, min(int(limit or 15), 50)),
              "includeItemsFromAllDrives": "true", "corpora": "allDrives"}
    if not query:   # Drive can't sort a fullText search
        params["orderBy"] = "modifiedTime desc"
    try:
        files = _get(account, "files", **params).get("files", [])
    except GoogleError as e:
        if "allDrives" not in str(e) and "corpora" not in str(e):
            raise
        params.pop("corpora"); params.pop("includeItemsFromAllDrives")
        files = _get(account, "files", **params).get("files", [])
    return [_brief(f) for f in files]


def read(account, file_id, offset=0):
    """A Google Doc, Sheet or Slides file as text, READ_CHUNK characters from offset."""
    meta = _get(account, f"files/{urllib.parse.quote(file_id)}", fields="id,name,mimeType,modifiedTime,webViewLink,owners(displayName)")
    kind = meta["mimeType"]
    if kind == FOLDER:
        return {**_brief(meta), "note": "This is a folder; list it with drive_search(folder_id=...).", "files": search(account, folder_id=file_id, limit=50)}
    if kind not in READABLE:
        raise GoogleError(f"'{meta['name']}' isn't a Google Doc, Sheet or Slides file, so it can't be opened here.")
    cached = _cache.get(file_id)
    if cached and cached[0] == meta.get("modifiedTime"):
        text = cached[1]
    else:
        if kind == SHEET:   # every tab, as tab-separated rows
            raw = _get(account, f"files/{urllib.parse.quote(file_id)}/export", raw=True,
                       mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            text = assistant._xlsx_text(raw)[0]
        else:
            text = _get(account, f"files/{urllib.parse.quote(file_id)}/export", raw=True, mimeType="text/plain").decode("utf-8-sig", "replace")
        text = text.replace("\r\n", "\n").strip()
        if len(_cache) >= _CACHE_MAX:
            _cache.pop(next(iter(_cache)))
        _cache[file_id] = (meta.get("modifiedTime"), text)
    offset = max(0, int(offset or 0))
    part = text[offset:offset + READ_CHUNK]
    out = {**_brief(meta), "characters": len(text), "offset": offset, "text": part}
    if offset + READ_CHUNK < len(text):
        out["more"] = f"The file continues; call drive_read again with offset={offset + READ_CHUNK} to read on."
    return out
