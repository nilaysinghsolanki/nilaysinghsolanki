#!/usr/bin/env python3
"""netsearch - offline people / customer / account search over your spreadsheets.

Everything runs on this PC: spreadsheets are indexed once into a local SQLite
full-text index (FTS5), and every search after that is a millisecond lookup.
No internet, no API, no credits.

    python netsearch.py index  PATH [PATH ...]   # files or folders (.xlsx .xlsm .csv)
    python netsearch.py search titan capital     # one-off search
    python netsearch.py                          # interactive prompt
    python netsearch.py serve                    # browser UI on http://127.0.0.1:8765
    python netsearch.py stats                    # what is indexed

Query syntax (terms are ANDed, every word is a prefix match):
    rahul                  anyone/anything matching "rahul..."
    "venture capital"      exact phrase
    name:abhi org:meddo    restrict to name / org / role / email / phone
    batch:2010..2015       any other column by (part of) its header, numeric range
    sector:vc help:fund    substring match on that column
    9899862689             phones match on digits, whatever their formatting
Typos are forgiven: if exact matching finds too little, a fuzzy pass on
names and organisations fills in (marked with ~).
"""
from __future__ import annotations

import argparse
import csv
import difflib
import fnmatch
import json
import os
import re
import sqlite3
import sys
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
DEFAULT_DB = HERE / "netsearch.db"
SCHEMAS_FILE = HERE / "schemas.json"
SUPPORTED = {".xlsx", ".xlsm", ".csv"}

# Canonical, searchable fields. Order matters: the first match wins, so
# "Account Name" lands in org, not name.
FIELD_KEYWORDS = [
    ("email", ["email", "e-mail", "mail id", "mail"]),
    ("phone", ["phone", "mobile", "contact no", "contact number", "whatsapp", "cell"]),
    ("org", ["company", "organisation", "organization", "org", "account",
             "customer", "firm", "employer", "startup", "client", "partner"]),
    ("role", ["designation", "title", "role", "position", "job"]),
    ("name", ["name", "person", "alumni", "member", "contact"]),
]
CANONICAL = ["name", "org", "role", "email", "phone"]
FIELD_ALIASES = {"company": "org", "account": "org", "customer": "org",
                 "organisation": "org", "organization": "org",
                 "title": "role", "designation": "role",
                 "mail": "email", "mobile": "phone"}
BM25_WEIGHTS = (10.0, 6.0, 3.0, 2.0, 2.0, 1.0)  # name, org, role, email, phone, other

EMAIL_RE = re.compile(r"[^@\s]+@[^@\s]+\.[A-Za-z]{2,}")
URL_RE = re.compile(r"https?://|www\.|linkedin\.com", re.I)
STAMP_RE = re.compile(r"^\d{1,2}/\d{1,2}/\d{2,4}\b")


# ---------------------------------------------------------------- reading

def clean(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).replace("_x000D_", "").replace("\r", "").strip()
    return re.sub(r"[ \t]+", " ", s)


def read_sheets(path: Path):
    """Yield (sheet_name, rows) where rows is a list of lists of strings."""
    if path.suffix.lower() == ".csv":
        with open(path, newline="", encoding="utf-8-sig", errors="replace") as f:
            yield "csv", [[clean(c) for c in r] for r in csv.reader(f)]
        return
    try:
        import openpyxl
    except ImportError:
        sys.exit("openpyxl is needed for Excel files:  pip install openpyxl")
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        for ws in wb.worksheets:
            yield ws.title, [[clean(c) for c in r] for r in ws.iter_rows(values_only=True)]
    finally:
        wb.close()


def load_schemas() -> dict:
    if SCHEMAS_FILE.exists():
        with open(SCHEMAS_FILE, encoding="utf-8") as f:
            return {k: v for k, v in json.load(f).items() if not k.startswith("_")}
    return {}


def schema_for(path: Path, schemas: dict):
    for pattern, cols in schemas.items():
        if fnmatch.fnmatch(path.name.lower(), pattern.lower()):
            return cols
    return None


def canonical_of(header: str):
    h = header.lower()
    for field, words in FIELD_KEYWORDS:
        if any(w in h for w in words):
            return field
    return None


def looks_like_header(row) -> bool:
    cells = [c for c in row if c]
    if len(cells) < 2:
        return False
    if any("@" in c or URL_RE.search(c) or c.replace(".", "").isdigit() for c in cells):
        return False
    if any(len(c) > 60 for c in cells):
        return False
    return any(canonical_of(c) for c in cells)


def infer_headers(rows, width):
    """Name columns of a headerless sheet from their contents."""
    sample = rows[:200]
    headers, have_name = [], False
    for i in range(width):
        vals = [r[i] for r in sample if i < len(r) and r[i]]
        n = len(vals) or 1
        share = lambda pred: sum(1 for v in vals if pred(v)) / n  # noqa: E731
        if not vals:
            headers.append(f"col{i + 1}")
        elif share(lambda v: EMAIL_RE.search(v)) > 0.5:
            headers.append("email" if "email" not in headers else f"email{i + 1}")
        elif share(lambda v: URL_RE.search(v)) > 0.5:
            headers.append("link")
        elif share(lambda v: 10 <= len(re.sub(r"\D", "", v)) <= 13 and len(v) < 20) > 0.5:
            headers.append("phone")
        elif share(lambda v: v.isdigit() and 1950 <= int(v) <= 2040) > 0.7:
            headers.append("year")
        elif share(lambda v: STAMP_RE.match(v)) > 0.7:
            headers.append("date")
        elif not have_name and share(lambda v: 1 <= len(v.split()) <= 4 and len(v) < 40
                                     and not any(ch.isdigit() for ch in v)) > 0.7:
            headers.append("name")
            have_name = True
        else:
            headers.append(f"col{i + 1}")
    return headers


def records_from_sheet(rows, schema=None):
    """Yield (row_number, {header: value}) for each non-empty row."""
    width = max((len(r) for r in rows), default=0)
    start = 0
    if schema:
        headers = list(schema) + [f"col{i + 1}" for i in range(len(schema), width)]
    elif rows and looks_like_header(rows[0]):
        headers = [h or f"col{i + 1}" for i, h in enumerate(rows[0])]
        start = 1
    else:
        headers = infer_headers(rows, width)
    seen = {}
    for i, h in enumerate(headers):  # make duplicate headers unique
        if h in seen:
            headers[i] = f"{h} ({i + 1})"
        seen[h] = True
    for n, r in enumerate(rows[start:], start=start + 1):
        rec = {headers[i]: v for i, v in enumerate(r) if v}
        if sum(1 for v in rec.values() if len(v) > 1) >= 1:
            yield n, rec


def split_fields(rec: dict) -> dict:
    out = {f: [] for f in CANONICAL + ["other"]}
    for h, v in rec.items():
        field = canonical_of(h) or "other"
        out[field].append(v)
    # Phones are also indexed as bare digits so any formatting matches.
    out["phone"] += [re.sub(r"\D", "", p) for p in out["phone"]]
    out["other"] += [h for h in rec]  # headers are searchable too (cheap context)
    return {k: " ".join(v) for k, v in out.items()}


# ---------------------------------------------------------------- index

def connect(db: Path) -> sqlite3.Connection:
    con = sqlite3.connect(db)
    con.executescript("""
        CREATE TABLE IF NOT EXISTS sources(path TEXT PRIMARY KEY, mtime REAL, rows INTEGER);
        CREATE TABLE IF NOT EXISTS records(
            id INTEGER PRIMARY KEY, path TEXT, sheet TEXT, rownum INTEGER, data TEXT);
        CREATE INDEX IF NOT EXISTS records_path ON records(path);
        CREATE VIRTUAL TABLE IF NOT EXISTS fts USING fts5(
            name, org, role, email, phone, other,
            tokenize='unicode61 remove_diacritics 2', prefix='2 3');
    """)
    return con


def iter_files(paths):
    for p in map(Path, paths):
        if p.is_dir():
            for f in sorted(p.rglob("*")):
                if f.suffix.lower() in SUPPORTED and not f.name.startswith("~$"):
                    yield f.resolve()
        elif p.suffix.lower() in SUPPORTED:
            yield p.resolve()
        else:
            print(f"skip (not .xlsx/.xlsm/.csv): {p}", file=sys.stderr)


def drop_source(con, path: str):
    ids = [r[0] for r in con.execute("SELECT id FROM records WHERE path=?", (path,))]
    con.executemany("DELETE FROM fts WHERE rowid=?", [(i,) for i in ids])
    con.execute("DELETE FROM records WHERE path=?", (path,))
    con.execute("DELETE FROM sources WHERE path=?", (path,))


def index(con, paths, force=False):
    schemas = load_schemas()
    total, t0 = 0, time.time()
    for f in iter_files(paths):
        key, mtime = str(f), f.stat().st_mtime
        old = con.execute("SELECT mtime FROM sources WHERE path=?", (key,)).fetchone()
        if old and old[0] == mtime and not force:
            print(f"unchanged  {f.name}")
            continue
        drop_source(con, key)
        count = 0
        try:
            for sheet, rows in read_sheets(f):
                for rownum, rec in records_from_sheet(rows, schema_for(f, schemas)):
                    cur = con.execute(
                        "INSERT INTO records(path, sheet, rownum, data) VALUES (?,?,?,?)",
                        (key, sheet, rownum, json.dumps(rec, ensure_ascii=False)))
                    flds = split_fields(rec)
                    con.execute(
                        "INSERT INTO fts(rowid, name, org, role, email, phone, other) "
                        "VALUES (?,?,?,?,?,?,?)",
                        (cur.lastrowid, *(flds[k] for k in CANONICAL + ["other"])))
                    count += 1
        except Exception as e:  # a broken file must not stop the rest
            print(f"ERROR      {f.name}: {e}", file=sys.stderr)
            continue
        con.execute("INSERT INTO sources VALUES (?,?,?)", (key, mtime, count))
        con.commit()
        total += count
        print(f"indexed    {f.name}: {count} rows")
    con.commit()
    print(f"done: {total} rows in {time.time() - t0:.1f}s")


def prune(con):
    gone = [p for (p,) in con.execute("SELECT path FROM sources") if not Path(p).exists()]
    for p in gone:
        drop_source(con, p)
        print(f"removed    {p}")
    con.commit()


# ---------------------------------------------------------------- search

PHONE_RE = re.compile(r"\+?\(?\d[\d()\-. ]{6,}\d")
TOKEN_RE = re.compile(r'(\w+):("[^"]*"|\S+)|("[^"]*")|(\S+)')


def parse_query(q: str):
    """-> (fts_terms [(field|None, text, is_phrase)], column_filters [(key, value)])."""
    terms, filters = [], []
    # "(989) 986-2689" / "+91 98765 43210" -> one digits-only token
    q = PHONE_RE.sub(lambda m: " " + re.sub(r"\D", "", m[0]) + " "
                     if len(re.sub(r"\D", "", m[0])) >= 7 else m[0], q)
    for m in TOKEN_RE.finditer(q):
        key, kval, phrase, word = m.groups()
        if key:
            key = FIELD_ALIASES.get(key.lower(), key.lower())
            val = kval.strip('"')
            if key in CANONICAL:
                terms.append((key, val, kval.startswith('"')))
            else:
                filters.append((key, val))
        elif phrase:
            terms.append((None, phrase.strip('"'), True))
        else:
            terms.append((None, word, False))
    return terms, filters


def fts_expr(terms) -> str:
    parts = []
    for field, text, phrase in terms:
        text = re.sub(r"\D", "", text) if field == "phone" or (
            text.replace("+", "").replace("-", "").replace(" ", "").isdigit()
            and len(re.sub(r"\D", "", text)) >= 6) else text
        words = re.findall(r"\w+", text)
        if not words:
            continue
        if phrase:
            expr = '"' + " ".join(words) + '"'
        else:
            expr = " ".join(f'"{w}"*' for w in words)
        parts.append(f"{field} : ({expr})" if field else f"({expr})")
    return " AND ".join(parts)


def passes(rec: dict, filters) -> bool:
    for key, val in filters:
        cols = [v for h, v in rec.items() if key in h.lower()]
        if not cols:
            return False
        rng = re.fullmatch(r"(\d*)\.\.(\d*)", val)
        if rng:
            lo = int(rng[1]) if rng[1] else -10**12
            hi = int(rng[2]) if rng[2] else 10**12
            if not any(c.isdigit() and lo <= int(c) <= hi for c in cols):
                return False
        elif not any(val.lower() in c.lower() for c in cols):
            return False
    return True


def fuzzy(con, text: str, exclude: set, filters, limit: int):
    """Typo-tolerant pass over names and orgs (rapidfuzz if present, else difflib)."""
    q = text.lower().strip()
    if len(q) < 3:
        return []
    rows = []
    for rid, data in con.execute("SELECT id, data FROM records"):
        if rid in exclude:
            continue
        rec = json.loads(data)
        keys = [v for h, v in rec.items() if canonical_of(h) in ("name", "org")]
        if keys:
            rows.append((rid, rec, keys))
    try:
        from rapidfuzz import fuzz
        score_fn = lambda a, b: fuzz.WRatio(a, b)  # noqa: E731
    except ImportError:
        def score_fn(a, b):
            best = difflib.SequenceMatcher(None, a, b).ratio()
            for w in b.split():  # also compare against each word ("rahl" ~ "rahul gupta")
                best = max(best, difflib.SequenceMatcher(None, a, w).ratio())
            return best * 100
    scored = []
    for rid, rec, keys in rows:
        s = max(score_fn(q, k.lower()) for k in keys)
        if s >= 75 and passes(rec, filters):
            scored.append((s, rid, rec))
    scored.sort(key=lambda x: -x[0])
    return scored[:limit]


def search(con, q: str, limit: int = 20):
    terms, filters = parse_query(q)
    expr = fts_expr(terms)
    results, seen = [], set()
    if expr:
        sql = (f"SELECT r.id, r.path, r.sheet, r.rownum, r.data FROM fts "
               f"JOIN records r ON r.id = fts.rowid WHERE fts MATCH ? "
               f"ORDER BY bm25(fts, {', '.join(map(str, BM25_WEIGHTS))}) LIMIT ?")
        try:
            rows = con.execute(sql, (expr, limit * 20 if filters else limit)).fetchall()
        except sqlite3.OperationalError:
            rows = []
    else:  # filters only
        rows = con.execute("SELECT id, path, sheet, rownum, data FROM records").fetchall()
    for rid, path, sheet, rownum, data in rows:
        rec = json.loads(data)
        if passes(rec, filters):
            results.append({"id": rid, "src": f"{Path(path).name}:{sheet}:{rownum}",
                            "rec": rec, "fuzzy": False})
            seen.add(rid)
            if len(results) >= limit:
                break
    free = " ".join(t for f, t, _ in terms if f in (None, "name", "org"))
    if len(results) < min(limit, 5) and free:
        for s, rid, rec in fuzzy(con, free, seen, filters, limit - len(results)):
            path, sheet, rownum = con.execute(
                "SELECT path, sheet, rownum FROM records WHERE id=?", (rid,)).fetchone()
            results.append({"id": rid, "src": f"{Path(path).name}:{sheet}:{rownum}",
                            "rec": rec, "fuzzy": True})
    return results


# ---------------------------------------------------------------- output

def pick(rec, field):
    return next((v for h, v in rec.items() if canonical_of(h) == field), "")


def render(results, full=False) -> str:
    if not results:
        return "  no matches"
    lines = []
    for i, r in enumerate(results, 1):
        rec = r["rec"]
        name, org, role = pick(rec, "name"), pick(rec, "org"), pick(rec, "role")
        head = name or org or next(iter(rec.values()), "?")
        sub = " @ ".join(x for x in (role, org if name else "") if x)
        lines.append(f"{'~' if r['fuzzy'] else ' '}{i:>2}. {head}" + (f"  -  {sub}" if sub else ""))
        shown = {name, org, role}
        for h, v in rec.items():
            if v in shown:
                continue
            if not full and len(v) > 140:
                v = v[:137] + "..."
            lines.append(f"      {h}: {v}")
        lines.append(f"      [{r['src']}]")
    return "\n".join(lines)


def export(results, path: str):
    headers = []
    for r in results:
        headers += [h for h in r["rec"] if h not in headers]
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=headers + ["source"])
        w.writeheader()
        for r in results:
            w.writerow({**r["rec"], "source": r["src"]})
    print(f"wrote {len(results)} rows to {path}")


def repl(con, limit):
    n = con.execute("SELECT count(*) FROM records").fetchone()[0]
    print(f"netsearch - {n} records indexed. Type a query, ':full' to toggle long text, "
          f"':csv FILE' to export the last results, empty line to quit.")
    full, last = False, []
    while True:
        try:
            q = input("\nsearch> ").strip()
        except (EOFError, KeyboardInterrupt):
            print()
            return
        if not q:
            return
        if q == ":full":
            full = not full
            print(f"full text {'on' if full else 'off'}")
            continue
        if q.startswith(":csv"):
            export(last, q[4:].strip() or "results.csv")
            continue
        t0 = time.perf_counter()
        last = search(con, q, limit)
        print(render(last, full))
        print(f"  ({len(last)} shown, {1000 * (time.perf_counter() - t0):.1f} ms)")


# ---------------------------------------------------------------- web UI

PAGE = """<!doctype html><meta charset=utf-8>
<meta name=viewport content="width=device-width,initial-scale=1">
<title>netsearch</title>
<style>
:root{--bg:#fafafa;--fg:#1a1a1a;--mute:#666;--card:#fff;--line:#e3e3e3;--acc:#2457d6}
@media (prefers-color-scheme:dark){:root{--bg:#141414;--fg:#eee;--mute:#999;--card:#1e1e1e;--line:#333;--acc:#7aa2ff}}
body{margin:0;background:var(--bg);color:var(--fg);font:15px/1.45 system-ui,sans-serif}
main{max-width:860px;margin:auto;padding:24px 16px}
input{width:100%;box-sizing:border-box;padding:12px 14px;font-size:17px;border:1px solid var(--line);
 border-radius:8px;background:var(--card);color:var(--fg)}
.hint{color:var(--mute);font-size:13px;margin:6px 2px 18px}
.card{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:12px 14px;margin:10px 0}
.h{font-weight:600}.s{color:var(--mute)}.f{font-size:13px;margin-top:4px}.f b{font-weight:500;color:var(--mute)}
.src{font-size:11px;color:var(--mute);margin-top:6px}a{color:var(--acc)}
</style>
<main><input id=q autofocus placeholder="Search people, companies, customers...">
<div class=hint>name:abhi &nbsp; org:meddo &nbsp; batch:2010..2015 &nbsp; sector:vc &nbsp; "exact phrase" &nbsp;
 - typos ok (~ = fuzzy match)</div><div id=out></div></main>
<script>
const q=document.getElementById('q'),out=document.getElementById('out');let t;
const esc=s=>s.replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const link=v=>/^https?:\\/\\//.test(v)?`<a href="${esc(v)}" target=_blank rel=noopener>${esc(v)}</a>`
 :/^[^@\\s]+@[^@\\s]+$/.test(v)?`<a href="mailto:${esc(v)}">${esc(v)}</a>`:esc(v);
q.oninput=()=>{clearTimeout(t);t=setTimeout(go,120)};
async function go(){if(!q.value.trim()){out.innerHTML='';return}
 const r=await (await fetch('/api?q='+encodeURIComponent(q.value))).json();
 out.innerHTML=r.ms.toFixed(1)+' ms'+(r.results.length?'':' - no matches')+r.results.map(x=>`<div class=card>
 <div class=h>${x.fuzzy?'~ ':''}${esc(x.title)} <span class=s>${esc(x.sub)}</span></div>
 ${x.fields.map(([h,v])=>`<div class=f><b>${esc(h)}:</b> ${link(v)}</div>`).join('')}
 <div class=src>${esc(x.src)}</div></div>`).join('')}
</script>"""


def serve(db: Path, port: int, limit: int):
    from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
    from urllib.parse import parse_qs, urlparse

    class H(BaseHTTPRequestHandler):
        def log_message(self, *a):
            pass

        def do_GET(self):
            u = urlparse(self.path)
            if u.path == "/api":
                q = parse_qs(u.query).get("q", [""])[0]
                con = connect(db)
                t0 = time.perf_counter()
                res = search(con, q, limit)
                ms = 1000 * (time.perf_counter() - t0)
                con.close()
                items = []
                for r in res:
                    rec = r["rec"]
                    name, org, role = pick(rec, "name"), pick(rec, "org"), pick(rec, "role")
                    items.append({
                        "title": name or org or next(iter(rec.values()), "?"),
                        "sub": " @ ".join(x for x in (role, org if name else "") if x),
                        "fields": [[h, v] for h, v in rec.items() if v not in (name, org, role)],
                        "src": r["src"], "fuzzy": r["fuzzy"]})
                body = json.dumps({"results": items, "ms": ms}).encode()
                ctype = "application/json"
            else:
                body, ctype = PAGE.encode(), "text/html; charset=utf-8"
            self.send_response(200)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

    srv = ThreadingHTTPServer(("127.0.0.1", port), H)  # local only, never exposed
    print(f"netsearch UI on http://127.0.0.1:{port}  (Ctrl+C to stop)")
    try:
        import webbrowser
        webbrowser.open(f"http://127.0.0.1:{port}")
    except Exception:
        pass
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass


# ---------------------------------------------------------------- CLI

def main(argv=None):
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # Windows consoles
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--db", default=os.environ.get("NETSEARCH_DB", DEFAULT_DB),
                    help="index file (default: netsearch.db next to this script)")
    ap.add_argument("-n", "--limit", type=int, default=20)
    sub = ap.add_subparsers(dest="cmd")
    p = sub.add_parser("index", help="add/refresh files or folders")
    p.add_argument("paths", nargs="+")
    p.add_argument("--force", action="store_true", help="re-read even unchanged files")
    sub.add_parser("prune", help="forget files that no longer exist")
    p = sub.add_parser("search", help="run one query")
    p.add_argument("query", nargs="+")
    p.add_argument("--full", action="store_true", help="do not truncate long text")
    p.add_argument("--csv", help="also export results to this CSV file")
    p.add_argument("--json", action="store_true", help="print raw JSON")
    p = sub.add_parser("serve", help="local browser UI")
    p.add_argument("--port", type=int, default=8765)
    sub.add_parser("stats", help="list indexed files")
    a = ap.parse_args(argv)

    db = Path(a.db)
    con = connect(db)
    if a.cmd == "index":
        index(con, a.paths, a.force)
    elif a.cmd == "prune":
        prune(con)
    elif a.cmd == "search":
        res = search(con, " ".join(a.query), a.limit)
        if a.json:
            print(json.dumps(res, ensure_ascii=False, indent=1))
        else:
            print(render(res, a.full))
        if a.csv:
            export(res, a.csv)
    elif a.cmd == "serve":
        con.close()
        serve(db, a.port, a.limit)
    elif a.cmd == "stats":
        for path, mtime, rows in con.execute("SELECT * FROM sources ORDER BY path"):
            print(f"{rows:>7}  {time.strftime('%Y-%m-%d %H:%M', time.localtime(mtime))}  {path}")
        print(f"{con.execute('SELECT count(*) FROM records').fetchone()[0]:>7}  total")
    else:
        repl(con, a.limit)


if __name__ == "__main__":
    main()
