"""Check the Report builder's share copy (HANDOFF 9e) against this script's own query of map_patients.

For many layouts it rebuilds every cell straight from the database and tests what the app would publish: every shown
number covers 11 or more patients, each column's Other is safe (11+ patients, marks can't be counted against it), a
row's totals add up its shown cells, the Total row is complete, the row order can be rebuilt from the published numbers,
and (with "files") the finished Excel file is two tabs of plain values that match the table.

    python3 tools/check_share_copy.py           the layout sweep (about a minute)
    python3 tools/check_share_copy.py files     also build and open real files (needs LibreOffice; a few minutes)

Read-only; prints counts and PASS / FAIL only, never a patient, a label with its numbers or a dollar figure."""
import contextlib, csv, io, pathlib, re, sqlite3, subprocess, sys, tempfile, warnings, zipfile
warnings.simplefilter("ignore")
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
import app, report_xlsx
LIMIT, MARK, OTHER = app.REPORT_SMALL, app.REPORT_MARK, app.REPORT_SHARE_OTHER
ro = sqlite3.connect(f"file:{app.DB_PATH}?mode=ro", uri=True); ro.row_factory = sqlite3.Row
c0 = app.db(); short = app.office_short_names(c0); c0.close()
where = app.city_names().get("where", {})
alloff = sorted({r[0] for r in ro.execute("select distinct office from map_patients")})
four = alloff[:4]   # any four offices will do
last = max(r[0] for r in ro.execute("select distinct substr(period_from,1,7) from map_patients"))
first = f"{last[:4]}-01" if last[5:] != "01" else last   # the latest year to date


def truth(spec):
    """{(label, col): (distinct patients, patient rows, production, collection)} by my own query."""
    key = "zip" if spec["source"] == "zip" else "UPPER(TRIM(referral_type))"
    cells = {}
    for r in ro.execute(f"select office, substr(period_from,1,7) m, {key} name, pat_id, production, collection from map_patients "
                        "where substr(period_from,1,7) between ? and ?", (spec["from"], spec["to"])):
        if spec.get("offices") and r["office"] not in spec["offices"]:
            continue
        name = r["name"] or "(not set)"
        if spec.get("group") == "city":
            name = where.get(name) or "(outside the map area)"
        name = report_xlsx.safe_label(name, strict=True)
        col = short.get(r["office"], r["office"]) if spec.get("columns") == "office" else r["m"] if spec.get("columns") == "month" else ""
        c = cells.setdefault((name, col), [set(), 0, 0.0, 0.0])
        c[0].add(r["pat_id"]); c[1] += 1; c[2] += r["production"] or 0; c[3] += r["collection"]
    return cells


def check(label, spec, files=False):
    spec = {"source": "zip", "group": "zip", "from": first, "to": last, "columns": "office",
            "measures": ["collection", "new_patients"], "offices": four, "share": True, **spec}
    if spec["source"] == "referral":
        spec["group"] = "referral"
    try:
        rep = app.report_data(spec)
    except ValueError as e:
        print(f"{label:34} refused: {e}"); return None
    t, cells, bad = report_xlsx.table(rep), truth(spec), []
    dim = rep["columns"]
    cols = rep["offices"] if dim == "office" else rep["months"] if dim == "month" else []
    n_lab = sum(1 for k in t["kinds"] if k == "text"); main = rep["measures"][0]; mkeys = [m["key"] for m in rep["measures"]]
    colkey = lambda c: c
    if not t["rows"]:
        print(f"{label:34} nothing to list (rows 0)"); return rep
    named = [r for r in t["rows"] if r[0] != OTHER]; other = [r for r in t["rows"] if r[0] == OTHER]
    if t["rows"][-1][0] != OTHER and other: bad.append("Other is not last")
    # 1. every shown cell covers >= LIMIT distinct patients; every mark is a real cell; zeros are true zeros
    pooled = {c: set() for c in (cols or [""])}
    for r in named:
        if cols:
            for j, c in enumerate(cols):
                v, tr = r[n_lab + j], cells.get((r[0], c))
                if v == MARK:
                    if not tr: bad.append("mark on an empty cell")
                    pooled[c].add(r[0])
                elif tr is None:
                    if v not in (0, None): bad.append("number on an empty cell")
                elif len(tr[0]) < LIMIT: bad.append(f"shown cell covers {len(tr[0])} patients")
            if all(r[n_lab + j] == MARK or not cells.get((r[0], c)) for j, c in enumerate(cols)): bad.append("row listed with nothing shown")
        else:
            tr = cells.get((r[0], ""))
            if not tr or len(tr[0]) < LIMIT: bad.append("shown row under the limit")
    listed = {r[0] for r in named}
    # 2. Other per column = everything not shown; covers >= LIMIT distinct patients
    kept = cols or [""]
    for c in kept:
        hidden = [(n, v) for (n, cc), v in cells.items() if cc == c and not (n in listed and n not in pooled[c])]
        ids = set().union(*(v[0] for _, v in hidden)) if hidden else set()
        if hidden and len(ids) < LIMIT: bad.append(f"Other in a column covers {len(ids)}")
        m = len(pooled[c])   # the marks a reader can count in this column
        if m and sum(v[1] for _, v in hidden) < m + LIMIT - 2: bad.append("the marks can be counted against Other")
        shown_cells = [v for (n, cc), v in cells.items() if cc == c and n in listed and n not in pooled[c]]
        for f in (2, 3):
            if m and not any(v[f] for _, v in hidden) and any(v[f] for v in shown_cells): bad.append("Other has no dollars beside marks")
    # 3. totals: each row's totals = its shown cells; Total row = column sums; complete against the truth
    def num(v): return 0 if v in (None, MARK) else v
    for r in t["rows"]:
        if cols and not main.get("num") and abs(sum(num(v) for v in r[n_lab:n_lab + len(cols)]) - num(r[n_lab + len(cols)])) > 0.5:
            bad.append("row total is not the sum of its shown cells")
    for j in range(n_lab, len(t["head"])):
        m = main if j < n_lab + len(cols) else rep["measures"][j - n_lab - len(cols)]
        if not m.get("num") and abs(sum(num(r[j]) for r in t["rows"]) - num(t["total"][j])) > 0.5: bad.append(f"Total row col {j} is not the column sum")
    if "new_patients" in mkeys:
        j = n_lab + len(cols) + mkeys.index("new_patients")
        true_n = sum(v[1] for (n, cc), v in cells.items() if cc in kept)
        if t["total"][j] != true_n: bad.append(f"grand total of new patients is not complete ({t['total'][j]} vs {true_n})")
    left = sorted({cc for _, cc in cells} - set(kept))
    # 4. the order can be rebuilt from the published numbers alone
    tot_j = n_lab + len(cols)
    if rep["measures"][0].get("num") is None:
        want = sorted(named, key=lambda r: (-(num(r[tot_j])), r[0]))
        if [r[0] for r in want] != [r[0] for r in named]: bad.append("row order differs from the published totals")
    # 5. counts shown are 0 or >= LIMIT
    for r in t["rows"] + [t["total"]]:
        for v, k in zip(r, t["kinds"]):
            if k == "count" and isinstance(v, (int, float)) and 0 < v < LIMIT: bad.append("a count under the limit is shown")
    out = f"{label:34} rows {len(named):3} marks {rep['info']['marks']:3} extra {rep['info']['extra']} left out {len(left)} other {rep['info']['other_pct']:2}%"
    if files:
        prof = app.DATA_DIR / "libreoffice-profile"
        x = report_xlsx.to_xlsx(rep, prof); z = zipfile.ZipFile(io.BytesIO(x))
        wb = z.read("xl/workbook.xml").decode()
        if re.findall(r'<sheet [^>]*?name="([^"]*)"', wb) != ["Summary", "Notes"]: bad.append("tabs")
        if any(re.search(rb"<f[ >/]", z.read(n)) for n in z.namelist() if n.startswith("xl/worksheets/")): bad.append("formula in file")
        if b"LibreOffice/" in z.read("docProps/app.xml") or b"Linux" in z.read("docProps/app.xml"): bad.append("app.xml still names the build")
        if len(z.namelist()) != 11: bad.append(f"{len(z.namelist())} parts")
        if {i.date_time for i in z.infolist()} != {(1980, 1, 1, 0, 0, 0)}: bad.append("zip dates")
        with tempfile.TemporaryDirectory() as tmp:
            f = pathlib.Path(tmp) / "s.xlsx"; f.write_bytes(x)
            subprocess.run(["soffice", f"-env:UserInstallation={prof.resolve().as_uri()}", "--headless", "--convert-to",
                            "csv:Text - txt - csv (StarCalc):44,34,76,1,,0,false,true,false,false,false,-1", str(f), "--outdir", tmp], capture_output=True, timeout=180)
            got = list(csv.reader(open(pathlib.Path(tmp) / "s-Summary.csv", encoding="utf-8")))
            notes = open(pathlib.Path(tmp) / "s-Notes.csv", encoding="utf-8").read()
        body = got[4:]
        if got[0][0] != rep["title"] or not got[1][0].startswith("Share copy."): bad.append("title/subtitle")
        want_rows = t["rows"] + [t["total"]]
        if len(body) != len(want_rows): bad.append(f"file has {len(body)} rows, table {len(want_rows)}")
        for g, w in zip(body, want_rows):
            for a, b, k in zip(g, w, t["kinds"]):
                b2 = "" if b is None else b if isinstance(b, str) else str(report_xlsx._stored(b, k))
                if str(a) != str(b2): bad.append(f"cell differs: file {a!r} table {b2!r}"); break
        if any(w in notes for w in ("Denticon", "Logger", "Data tab")): bad.append("notes name internal things")
        csvb = list(csv.reader(io.StringIO(report_xlsx.to_csv(rep).decode("utf-8-sig"))))
        if len(csvb) != len(want_rows) + 1 or csvb[0] != t["head"]: bad.append("share csv shape")
        if any("." in v for row in csvb[1:] for v, k in zip(row, t["kinds"]) if k != "text"): bad.append("csv keeps decimals")
        out += f" | xlsx {len(x)//1024} KB ok"
    print(out, "| PASS" if not bad else "| FAIL: " + "; ".join(sorted(set(bad))[:6]))
    return rep


if __name__ == "__main__":
    full = sys.argv[1:] == ["files"]
    months = sorted({r[0] for r in ro.execute("select distinct substr(period_from,1,7) from map_patients")})
    q3 = months[-3] if len(months) >= 3 else months[0]
    check("R1 ZIP x office, 4 offices", {}, full)
    check("R2 ZIP x month, last 3 months", {"from": q3, "columns": "month"}, full)
    check("R3 ZIP totals only", {"columns": ""}, full)
    check("R4 referral x office, all, 3 months", {"source": "referral", "offices": [], "from": q3}, full)
    check("R5 referral x month, all", {"source": "referral", "offices": [], "columns": "month"}, full)
    check("R6 city x office", {"group": "city"}, full)
    check("ratio first (per new patient)", {"measures": ["collection_per_np", "new_patients", "collection", "production"]}, full)
    check("new patients first", {"measures": ["new_patients", "collection"]}, full)
    if full: sys.exit()
    n = fails = 0
    for src, grp in (("zip", "zip"), ("zip", "city"), ("referral", "referral")):
        for cols in ("office", "month", ""):
            for offs in ([o] for o in alloff[:15:3]):
                for m in months[::3]:
                    buf = io.StringIO()
                    with contextlib.redirect_stdout(buf): check("x", {"source": src, "group": grp, "columns": cols, "offices": offs, "from": m, "to": m})
                    n += 1; fails += "FAIL" in buf.getvalue()
                    if "FAIL" in buf.getvalue(): print(src, grp, cols, len(offs), m, buf.getvalue().strip()[-200:])
            for offs in (four, alloff, alloff[:2], alloff[5:9]):
                for a, b in ((first, last), (q3, last), (last, last), (months[0], months[min(2, len(months) - 1)]), (months[0], last)):
                    for meas in (["collection", "new_patients"], ["new_patients"], ["production", "collection_per_np"]):
                        buf = io.StringIO()
                        with contextlib.redirect_stdout(buf): check("x", {"source": src, "group": grp, "columns": cols, "offices": offs, "from": a, "to": b, "measures": meas})
                        n += 1; fails += "FAIL" in buf.getvalue()
                        if "FAIL" in buf.getvalue(): print(src, grp, cols, len(offs), a, b, meas, buf.getvalue().strip()[-200:])
    print(f"sweep: {n} share copies checked, {fails} failed")
    sys.exit(1 if fails else 0)
