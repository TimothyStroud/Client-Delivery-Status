"""Generate a static HTML COBC / TRR / ABII File Report dashboard.

Self-contained single .html file: embedded JSON data + vanilla JS filtering.
No external dependencies, no server, no auth. Styled to match the
MissingPayDates / CAQH Traffic / MSPI reports. Drop it on a shared drive and
users open it in a browser.

Data source: the **RAMP web app** (http://ramp), not SQL.  Two REST endpoints
behind the RAMP UI pages, both via Windows auth (`curl --negotiate -u :`):

  * FileLogger  -> GET /api/Ramp/FileLog/List/{startDate}/{endDate}
      Backs http://ramp/Ramp/FileLogger.  Dates are JS `toDateString()`
      format ("Tue Jul 21 2026").  Returns every file RAMP logged in the
      range, with ClientName / FeedName / JobName already resolved, the
      timestamp in LogDate (DateCreated is not populated by this projection),
      and one row per feed/job step (so a file appears a few times -> deduped).
      Pulled in monthly chunks from START_DATE to today.

  * UnconfiguredFiles -> GET /api/Ramp/ConfiguredFiles
      Backs http://ramp/Ramp/UnconfiguredFiles.  Items with
      IsConfigured == false are files RAMP received but has no config for.
      Joined to the FileLogger rows by FileLogId so the dashboard can flag a
      file as UNCONFIGURED (and any unconfigured COBC/TRR/ABII outside the
      pulled window is injected so it still shows).

We keep only COBC (MARXCOB), TRR and ABII files (matched on FileName), dedupe
to one row per file (FileName + log-date, encryption-wrapper twins merged),
and take the client straight from RAMP's ClientName.

Output paths (overwritten each run):
- \\\\trgfile1\\Shared\\DIG\\Data Business Delivery Team\\Delivery Schedule\\Daily Status Reports\\COBCReport.html
- C:\\Users\\tls2\\.claude\\projects\\H--\\COBCReport.html
- C:\\Users\\tls2\\OneDrive - Machinify\\Documents\\Reports\\COBCReport.html

Run:
    python C:\\Users\\tls2\\.claude\\projects\\H--\\cobc_report.py
"""

import json
import os
import re
import shutil
import subprocess
import sys
import time
import urllib.parse
from datetime import date, datetime, timedelta

RAMP_BASE = "http://ramp"
# History start (matches the prior report's "since 2024-12-31").
START_DATE = date(2025, 1, 1)

# --- Rolling-window cache -------------------------------------------------
# Re-pulling the full START_DATE->today history every run (~2.3M FileLogger
# rows over 19 months) is expensive, and the report now runs hourly on
# weekdays.  Instead we cache the built rows and, on each subsequent run,
# only re-fetch a recent rolling window (by LogDate) and merge it over the
# cached history.  Pass --full to force a from-scratch rebuild.
CACHE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                          "cobc_report_cache.json")
# Days back from today to refresh live each run.  Generous so late-arriving
# files and any recently-unconfigured file land inside the refreshed window.
WINDOW_DAYS = 45

OUTPUT_PATHS = [
    r"\\trgfile1\Shared\DIG\Data Business Delivery Team\Delivery Schedule\Daily Status Reports\COBCReport.html",
    r"C:\Users\tls2\.claude\projects\H--\COBCReport.html",
    r"C:\Users\tls2\OneDrive - Machinify\Documents\Reports\COBCReport.html",
]

# ---- File-name parsing --------------------------------------------------

# A contract dot/underscore-segment: optional routing 'R', a letter, then 3-5
# digits (anchored so date/time segments like D260622 / T0858590 don't match).
CONTRACT_SEG_RE = re.compile(r"^R?[A-Z]\d{3,5}$")
# Drop a single leading routing 'R' only when a real contract (letter+digit)
# follows it: RH1607 -> H1607, RR6694 -> R6694, but R6694 stays R6694.
LEAD_R_RE = re.compile(r"^R(?=[A-Z]\d)")

# Data-date tokens, tried in priority order against the filename.
DATE_D_RE    = re.compile(r"^D(\d{2})(\d{2})(\d{2})$")         # D260508      -> 2026-05-08
DATE_DTRR_RE = re.compile(r"DTRRD_(\d{2})(\d{2})(\d{2})")      # DTRRD_260721 -> 2026-07-21 (YYMMDD)
DATE_TRR_RE  = re.compile(r"TRR(\d{2})(\d{2})(\d{4})")         # TRR07232026  -> 2026-07-23 (MMDDYYYY)
DATE_YMD_RE  = re.compile(r"(20\d{2})(\d{2})(\d{2})")          # 20260722     -> 2026-07-22 (YYYYMMDD)
DATE_MDY_RE  = re.compile(r"^(\d{2})(\d{2})(20\d{2})$")        # 01282026     -> 2026-01-28 (MMDDYYYY)
DATE_YM_RE   = re.compile(r"^(20\d{2})(\d{2})$")               # 202501       -> 2025-01-01 (YYYYMM)

# Leaf-name wrappers that are pure encryption layers over an otherwise
# identical file; strip them so encrypted+decrypted twins dedupe to one file.
ENC_EXT_RE = re.compile(r"\.(pgp|gpg)$", re.IGNORECASE)


def leaf_name(path: str) -> str:
    return re.split(r"[\\/]", path or "")[-1]


def dedup_key(fname: str) -> str:
    """Filename with an encryption wrapper (.pgp/.gpg) stripped, upper-cased."""
    return ENC_EXT_RE.sub("", leaf_name(fname)).upper()


def row_type(fname: str, job: str = None, feed: str = None) -> str:
    """Classify a file the way RAMP does. Filename tokens are strongest, but
    many files carry the type only in the JobName/FeedName (RAMP's own
    classification), e.g. ABII Transplant reports named
    '<contract>_Transplant_Report_YYYYMM.xlsx' under 'Wellpoint RX ABII 0100
    Stage', or Centene Fidelis COBC named 'COBD_YYYYMMDD_..._H5599.txt'."""
    u = (fname or "").upper()
    if "MARX" in u:
        return "COBC"
    if "ABII" in u:
        return "ABII"
    if "TRR" in u:
        return "TRR"
    jf = ((job or "") + " " + (feed or "")).upper()
    if "MARXCOB" in jf or "COBC" in jf:
        return "COBC"
    if "ABII" in jf:
        return "ABII"
    if "DTRR" in jf or "TRR" in jf:
        return "TRR"
    return "Other"


def parse_contract(fname: str) -> str:
    """Contract token in the leaf file name, leading routing 'R' dropped."""
    for seg in re.split(r"[._ ]", leaf_name(fname).upper()):
        if CONTRACT_SEG_RE.match(seg):
            return LEAD_R_RE.sub("", seg)
    return ""


def parse_extracted(fname: str) -> str:
    """Best-effort data (extracted) date from the leaf name; '' if none."""
    up = leaf_name(fname).upper()
    segs = re.split(r"[._ ]", up)
    for seg in segs:                       # D260508 style
        m = DATE_D_RE.match(seg)
        if m:
            return f"20{m.group(1)}-{m.group(2)}-{m.group(3)}"
    m = DATE_DTRR_RE.search(up)            # DTRRD_260721 (YYMMDD)
    if m:
        return f"20{m.group(1)}-{m.group(2)}-{m.group(3)}"
    m = DATE_TRR_RE.search(up)             # TRR07232026 (MMDDYYYY)
    if m:
        return f"{m.group(3)}-{m.group(1)}-{m.group(2)}"
    for seg in segs:                       # 01282026 (MMDDYYYY)
        m = DATE_MDY_RE.match(seg)
        if m:
            return f"{m.group(3)}-{m.group(1)}-{m.group(2)}"
    m = DATE_YMD_RE.search(up)             # 20260722 (YYYYMMDD)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    for seg in segs:                       # 202501 (YYYYMM) -> month, day=01
        m = DATE_YM_RE.match(seg)
        if m:
            return f"{m.group(1)}-{m.group(2)}-01"
    return ""


# ---- RAMP API -----------------------------------------------------------

def ramp_get_json(url: str, tries: int = 5):
    """GET a RAMP endpoint via NTLM/negotiate, retrying transient non-JSON
    bodies (RAMP intermittently returns an empty/HTML body)."""
    for _ in range(tries):
        out = subprocess.run(
            ["curl", "-s", "--negotiate", "-u", ":", url],
            capture_output=True, text=True,
        ).stdout
        try:
            return json.loads(out)
        except json.JSONDecodeError:
            time.sleep(5)
    raise RuntimeError(f"RAMP returned no valid JSON after {tries} tries: {url}")


def _unwrap(payload):
    """RAMP wraps rows as {'Data': [[...rows...]]}; unwrap to the row list."""
    data = payload.get("Data")
    if isinstance(data, list) and data and isinstance(data[0], list):
        return data[0]
    return data or []


def to_date_string(d: date) -> str:
    """JS Date.toDateString() format, e.g. 'Tue Jul 21 2026'."""
    return d.strftime("%a %b %d %Y")


def month_chunks(start: date, end: date):
    """Consecutive (start, end) day ranges, one per calendar month, with a
    1-day overlap at boundaries (dedup removes the overlap)."""
    cur = start
    while cur <= end:
        nxt = date(cur.year + 1, 1, 1) if cur.month == 12 else date(cur.year, cur.month + 1, 1)
        yield cur, min(nxt, end)
        cur = nxt


def fetch_filelog(start: date, end: date):
    url = (f"{RAMP_BASE}/api/Ramp/FileLog/List/"
           f"{urllib.parse.quote(to_date_string(start))}/"
           f"{urllib.parse.quote(to_date_string(end))}")
    return _unwrap(ramp_get_json(url))


def fetch_unconfigured():
    """Items RAMP received with no matching config (IsConfigured == false)."""
    items = _unwrap(ramp_get_json(f"{RAMP_BASE}/api/Ramp/ConfiguredFiles"))
    return [i for i in items if not i.get("IsConfigured")]


# ---- Data assembly ------------------------------------------------------

def _mkrow(fname, client, feed, log_ts, size, unconf, ftype):
    return {
        "client": (client or "Unknown").strip() or "Unknown",
        "ftype": ftype,
        "contract": parse_contract(fname),
        "filename": leaf_name(fname),
        "extracted": parse_extracted(fname),
        "logged": (log_ts or "")[:19].replace("T", " "),
        "feed": (feed or "").strip(),
        "size": int(size) if isinstance(size, (int, float)) else 0,
        "unconf": 1 if unconf else 0,
    }


def _ingest(by_key, fname, client, feed, log_ts, size, fid, unconf, job=None):
    """Type + dedupe one file into by_key (dedup_key|log-day -> row)."""
    ftype = row_type(fname, job, feed)
    if ftype == "Other":
        return
    day = (log_ts or "")[:10]
    key = dedup_key(fname) + "|" + day
    row = by_key.get(key)
    if row is None:
        row = _mkrow(fname, client, feed, log_ts, size, unconf, ftype)
        row["_ids"] = set()
        by_key[key] = row
    else:
        # Keep the earliest log timestamp; OR-in the unconfigured flag.
        lg = (log_ts or "")[:19].replace("T", " ")
        if lg and (not row["logged"] or lg < row["logged"]):
            row["logged"] = lg
        if unconf:
            row["unconf"] = 1
        if not row["size"] and isinstance(size, (int, float)):
            row["size"] = int(size)
    if fid:
        row["_ids"].add(fid)


def _ingest_filelog_range(by_key, start, end, unconf_ids):
    """Pull FileLogger in monthly chunks over [start, end] into by_key."""
    total_raw = 0
    for cs, ce in month_chunks(start, end):
        rows = fetch_filelog(cs, ce)
        total_raw += len(rows)
        for r in rows:
            fid = r.get("FileLogId")
            _ingest(by_key, r.get("FileName"), r.get("ClientName"), r.get("FeedName"),
                    r.get("LogDate"), r.get("FileSize"), fid, fid in unconf_ids,
                    r.get("JobName"))
        print(f"[info]   {cs} -> {ce}: {len(rows):>7} rows  (running files: {len(by_key)})")
    return total_raw


def _inject_unconfigured(by_key, unconfigured):
    """Inject unconfigured COBC/TRR/ABII (may be outside the pulled window)."""
    for i in unconfigured:
        fl = i.get("FileLog") or {}
        fname = i.get("File") or fl.get("FileName")
        log_ts = fl.get("LogDate") or fl.get("DateCreated") or i.get("CreateDate")
        _ingest(by_key, fname, fl.get("ClientName"), fl.get("FeedName"),
                log_ts, fl.get("FileSize"), i.get("FileLogId"), True,
                fl.get("JobName"))


def _finalize(by_key):
    rows = list(by_key.values())
    for r in rows:
        r.pop("_ids", None)
    return rows


def build_rows_full():
    """Full rebuild: pull the entire START_DATE->today history from RAMP."""
    today = datetime.now().date()
    unconfigured = fetch_unconfigured()
    unconf_ids = {i.get("FileLogId") for i in unconfigured if i.get("FileLogId")}
    print(f"[info] {len(unconfigured)} unconfigured file(s) in RAMP right now")

    by_key = {}
    total_raw = _ingest_filelog_range(by_key, START_DATE, today, unconf_ids)
    _inject_unconfigured(by_key, unconfigured)
    rows = _finalize(by_key)
    print(f"[info] {total_raw} raw FileLogger rows -> {len(rows)} files (deduped) [FULL]")
    return rows


# ---- Rolling-window cache -----------------------------------------------

def _load_cache():
    """Return the cached payload, or None if missing/invalid/stale schema."""
    try:
        with open(CACHE_PATH, "r", encoding="utf-8") as f:
            c = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return None
    if not isinstance(c.get("rows"), list):
        return None
    if c.get("start_date") != START_DATE.isoformat():
        print("[info] cache START_DATE differs from current -> full rebuild")
        return None
    return c


def _save_cache(rows):
    payload = {
        "built": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "start_date": START_DATE.isoformat(),
        "window_days": WINDOW_DAYS,
        "rows": rows,
    }
    tmp = CACHE_PATH + ".tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(payload, f, separators=(",", ":"))
        os.replace(tmp, CACHE_PATH)
        print(f"[info] cache saved ({len(rows)} rows) -> {CACHE_PATH}")
    except OSError as e:
        print(f"[warn] couldn't save cache: {e}")


def build_rows_rolling():
    """Merge a fresh rolling window over the cached history.

    Rows older than the window come straight from the cache; rows whose log
    date falls inside the window are dropped and re-fetched live, so recent
    additions/updates/removals are reflected.  Unconfigured files are pulled
    and injected in full every run (cheap, and keeps the flag current for
    anything inside the refreshed window).  Falls back to a full rebuild when
    there's no usable cache.
    """
    cache = _load_cache()
    if cache is None:
        rows = build_rows_full()
        _save_cache(rows)
        return rows

    today = datetime.now().date()
    window_start = today - timedelta(days=WINDOW_DAYS)
    ws = window_start.isoformat()

    unconfigured = fetch_unconfigured()
    unconf_ids = {i.get("FileLogId") for i in unconfigured if i.get("FileLogId")}
    print(f"[info] {len(unconfigured)} unconfigured file(s) in RAMP right now")

    # Seed with cached rows OLDER than the window; in-window rows get refetched.
    by_key = {}
    kept = 0
    for r in cache["rows"]:
        day = (r.get("logged") or "")[:10]
        if day and day >= ws:
            continue
        r = dict(r)
        r["_ids"] = set()
        by_key[dedup_key(r.get("filename", "")) + "|" + day] = r
        kept += 1
    print(f"[info] cache built {cache.get('built')}: kept {kept} rows older than {ws}")

    total_raw = _ingest_filelog_range(by_key, window_start, today, unconf_ids)
    _inject_unconfigured(by_key, unconfigured)
    rows = _finalize(by_key)
    print(f"[info] rolling {ws}->{today}: {total_raw} raw window rows -> "
          f"{len(rows)} files total (deduped) [ROLLING]")
    _save_cache(rows)
    return rows


HTML_TEMPLATE = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>COBC / TRR / ABII File Report</title>
<style>
  :root {
    --bg: #f4f6f9;
    --card: #ffffff;
    --border: #d8dee6;
    --text: #1f2a37;
    --muted: #5b6776;
    --accent: #2c5f8a;
    --accent-dark: #1f3d5c;
    --navy: #0a1f4d;
  }
  * { box-sizing: border-box; }
  body {
    margin: 0;
    font-family: "Segoe UI", -apple-system, BlinkMacSystemFont, sans-serif;
    background: var(--bg);
    color: var(--text);
    font-size: 14px;
  }
  header {
    background: var(--accent-dark);
    color: #39ff14;
    padding: 16px 24px;
  }
  header h1 { margin: 0; font-size: 18px; font-weight: 600; }
  header .meta { font-size: 12px; opacity: 0.85; margin-top: 4px; }
  main { padding: 16px 24px 32px; }
  .tabs { display: flex; gap: 4px; margin-bottom: 16px; }
  .tab {
    background: var(--card);
    border: 1px solid var(--border);
    border-bottom: none;
    border-radius: 6px 6px 0 0;
    padding: 8px 18px;
    cursor: pointer;
    font-size: 13px;
    font-weight: 600;
    color: var(--muted);
  }
  .tab.active { background: var(--accent); color: #39ff14; border-color: var(--accent); }
  .view { display: none; }
  .view.active { display: block; }
  .kpis {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
    gap: 12px;
    margin-bottom: 16px;
  }
  .kpi {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 6px;
    padding: 12px 16px;
  }
  .kpi .label { font-size: 11px; color: var(--muted); text-transform: uppercase; letter-spacing: 0.5px; }
  .kpi .value { font-size: 22px; font-weight: 600; margin-top: 2px; color: var(--accent-dark); }
  .kpi.warn .value { color: #b23b2e; }
  .filters {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 6px;
    padding: 12px 16px;
    margin-bottom: 16px;
    display: flex;
    flex-wrap: wrap;
    gap: 12px;
    align-items: end;
  }
  .field { display: flex; flex-direction: column; gap: 4px; }
  .field label { font-size: 11px; color: var(--muted); text-transform: uppercase; letter-spacing: 0.5px; font-weight: 700; }
  .field select, .field input {
    border: 1px solid var(--border);
    border-radius: 4px;
    padding: 6px 8px;
    font-size: 13px;
    background: #fff;
    min-width: 120px;
  }
  .field input[type=date] { min-width: 140px; }
  .field.check { flex-direction: row; align-items: center; gap: 6px; }
  .field.check input { min-width: 0; }
  .field.check label { text-transform: none; font-weight: 600; color: var(--text); letter-spacing: 0; }
  button {
    background: var(--accent);
    color: #fff;
    border: 0;
    border-radius: 4px;
    padding: 6px 14px;
    cursor: pointer;
    font-size: 13px;
  }
  button.secondary { background: #fff; color: var(--accent); border: 1px solid var(--accent); }
  table {
    width: auto;
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 6px;
    border-collapse: separate;
    border-spacing: 0;
    overflow: hidden;
  }
  thead { background: var(--accent); color: #39ff14; position: sticky; top: 0; z-index: 1; }
  th {
    text-align: left;
    padding: 6px 8px;
    font-weight: 600;
    font-size: 12px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    cursor: pointer;
    user-select: none;
    white-space: nowrap;
  }
  th .arrow { opacity: 0.5; font-size: 10px; margin-left: 3px; }
  th.sort-asc .arrow, th.sort-desc .arrow { opacity: 1; }
  td {
    padding: 4px 8px;
    border-top: 1px solid var(--border);
    font-size: 13px;
    white-space: nowrap;
  }
  td.num { text-align: right; font-variant-numeric: tabular-nums; }
  td.zero { color: #b6bec8; }
  tr:hover td { background: #f8fafc; }
  tr.is-unconf td { background: #fff6f4; }
  tr.is-unconf:hover td { background: #ffeee9; }
  .pager { display: flex; align-items: center; gap: 12px; margin-top: 12px; justify-content: flex-end; font-size: 13px; color: var(--muted); }
  .empty { padding: 24px; text-align: center; color: var(--muted); }

  /* File-type pills */
  .ft { display: inline-block; padding: 1px 8px; border-radius: 10px; font-size: 11px; font-weight: 700;
        letter-spacing: .4px; }
  .ft-COBC { background: #e3f0fb; color: #1f5b8a; }
  .ft-TRR  { background: #eae6fb; color: #5b3a9e; }
  .ft-ABII { background: #e4f6ea; color: #1c7a44; }
  /* Unconfigured badge */
  .uc { display: inline-block; margin-left: 8px; padding: 1px 7px; border-radius: 10px;
        font-size: 10px; font-weight: 800; letter-spacing: .5px;
        background: #fbe4df; color: #b23b2e; border: 1px solid #f0b6ab; vertical-align: middle; }

  /* ---- Monthly matrix ---- */
  .matrix-wrap {
    overflow: auto; max-height: calc(100vh - 230px);
    border: 1px solid var(--border); border-radius: 10px; background: var(--card);
    box-shadow: 0 1px 3px rgba(16,32,64,.06), 0 10px 28px rgba(16,32,64,.05);
  }
  table.matrix { border-collapse: separate; border-spacing: 0; width: auto; border: 0; }
  table.matrix th, table.matrix td { border-bottom: 1px solid #eef1f5; white-space: nowrap; }

  /* Month header row */
  table.matrix thead th {
    position: sticky; top: 0; z-index: 3;
    background: linear-gradient(#ffffff, #f2f6fb); color: var(--accent-dark);
    padding: 10px 10px; font-size: 11px; text-align: center; font-weight: 700;
    text-transform: uppercase; letter-spacing: .6px;
    border-bottom: 2px solid var(--accent);
  }
  table.matrix thead th.rowhdr { z-index: 4; text-align: left; }

  /* Frozen Client & Contract columns */
  table.matrix th.rowhdr, table.matrix td.rowhdr {
    padding: 7px 14px; font-size: 12.5px; position: sticky;
    overflow: hidden; text-overflow: ellipsis;
  }
  table.matrix td.rowhdr { background: #fff; color: var(--text); z-index: 2; }
  table.matrix th.c-client, table.matrix td.c-client {
    left: 0; width: 190px; min-width: 190px; max-width: 190px;
    border-right: 1px solid #eef1f5;
  }
  table.matrix th.c-contract, table.matrix td.c-contract {
    left: 190px; width: 120px; min-width: 120px; max-width: 120px;
    font-variant-numeric: tabular-nums; color: var(--muted);
    border-right: 2px solid var(--border);
  }

  /* Month value cells + the file-count load badge */
  table.matrix td.day { text-align: center; padding: 6px; min-width: 64px; }
  table.matrix td.day.hit { cursor: pointer; }
  .mk {
    position: relative;
    display: inline-flex; align-items: center; justify-content: center;
    min-width: 22px; height: 22px; border-radius: 6px; padding: 0 8px;
    background: #e3f4ea; color: #1c7a44; font-weight: 700; font-size: 12px;
    box-shadow: inset 0 0 0 1px rgba(28,122,68,.28);
    transition: transform .08s ease;
    font-variant-numeric: tabular-nums;
  }
  table.matrix td.day.hit:hover .mk { background: #cdecd8; transform: scale(1.12); }
  /* red corner dot when the cell contains any unconfigured file */
  .mk.has-unconf::after {
    content: ""; position: absolute; top: -3px; right: -3px;
    width: 8px; height: 8px; border-radius: 50%;
    background: #d64426; box-shadow: 0 0 0 1.5px #fff;
  }
  .mx-empty { padding: 28px; text-align: center; color: var(--muted); }

  /* Client group (expandable) row */
  tr.client-row td { background: #f5f9fd; border-bottom: 1px solid #dce7f3; }
  tr.client-row td.rowhdr { cursor: pointer; background: #eef4fb; }
  tr.client-row:hover td { background: #e6effb; }
  tr.client-row:hover td.rowhdr { background: #dfeafb; }
  tr.client-row td.c-client { font-weight: 700; color: var(--accent-dark); box-shadow: inset 3px 0 0 var(--accent); }
  tr.client-row .mk { background: var(--accent); color: #fff; box-shadow: none; }   /* rollup badge */
  tr.client-row td.day.hit:hover .mk { background: var(--accent-dark); }
  .tri { display: inline-block; width: 12px; margin-right: 5px; color: var(--accent); font-size: 10px; }

  /* Contract child rows */
  tr.contract-row:hover td { background: #f4f8fd; }
  tr.contract-row:hover td.rowhdr { background: #eef4fb; }
  tr.contract-row td.c-contract { padding-left: 24px; }
  tr.contract-row td.c-contract::before {
    content: ""; position: absolute; left: 13px; top: 50%; width: 5px; height: 5px;
    margin-top: -3px; border-radius: 50%; background: #c2ccd8;
  }

  #tooltip {
    position: fixed; z-index: 9999; pointer-events: none;
    background: #1f2a37; color: #fff; border-radius: 6px;
    padding: 8px 10px; font-size: 12px; line-height: 1.45;
    box-shadow: 0 4px 16px rgba(0,0,0,.3); max-width: 460px;
  }
  #tooltip .tt-file + .tt-file { margin-top: 6px; border-top: 1px solid #3a4757; padding-top: 6px; }
  #tooltip .tt-name { font-family: Consolas, "Courier New", monospace; word-break: break-all; }
  #tooltip .tt-label { color: #9fb4cc; }
  #tooltip .tt-uc { color: #ff9b8a; font-weight: 700; }
</style>
</head>
<body>
<header>
  <h1>COBC / TRR / ABII File Report</h1>
  <div class="meta">Generated __GENERATED__ &middot; source: RAMP FileLogger + Unconfigured Files (http://ramp) &middot; since __SINCE__ &middot; __ROW_COUNT__ files &middot; <span style="color:#ff9b8a">__UNCONF_COUNT__ unconfigured</span></div>
</header>
<main>
  <div class="tabs">
    <div class="tab active" data-view="monthly" data-ftype="COBC">COBC</div>
    <div class="tab" data-view="monthly" data-ftype="TRR">TRR</div>
    <div class="tab" data-view="monthly" data-ftype="ABII">ABII</div>
    <div class="tab" data-view="detail">Detail</div>
  </div>

  <section id="view-monthly" class="view active">
    <section class="filters">
      <div class="field">
        <label for="m-year">Year</label>
        <select id="m-year"></select>
      </div>
      <div class="field">
        <label for="m-client">Client</label>
        <select id="m-client"><option value="">All clients</option></select>
      </div>
      <div class="field">
        <label for="m-contract">Contract</label>
        <select id="m-contract"><option value="">All contracts</option></select>
      </div>
      <button class="secondary" id="m-expand">Expand all</button>
      <button class="secondary" id="m-collapse">Collapse all</button>
      <div class="field" style="justify-content:flex-end">
        <label>&nbsp;</label>
        <span style="font-size:12px;color:var(--muted)">Cell = files logged that month &middot; a <b style="color:#d64426">&bull;</b> dot marks a month containing an unconfigured file &middot; click a client to expand &middot; hover a <b>count</b> for details.</span>
      </div>
    </section>
    <div class="matrix-wrap">
      <table class="matrix" id="matrix">
        <thead id="matrix-head"></thead>
        <tbody id="matrix-body"></tbody>
      </table>
    </div>
  </section>

  <section id="view-detail" class="view">
  <section class="kpis">
    <div class="kpi"><div class="label">Files</div><div class="value" id="kpi-files">&mdash;</div></div>
    <div class="kpi"><div class="label">Clients</div><div class="value" id="kpi-clients">&mdash;</div></div>
    <div class="kpi"><div class="label">Contracts</div><div class="value" id="kpi-contracts">&mdash;</div></div>
    <div class="kpi warn"><div class="label">Unconfigured</div><div class="value" id="kpi-unconf">&mdash;</div></div>
    <div class="kpi"><div class="label">Total Size</div><div class="value" id="kpi-size">&mdash;</div></div>
    <div class="kpi"><div class="label">Latest Logged</div><div class="value" id="kpi-latest">&mdash;</div></div>
  </section>

  <section class="filters">
    <div class="field">
      <label for="f-type">File Type</label>
      <select id="f-type"><option value="">All types</option><option>COBC</option><option>TRR</option><option>ABII</option></select>
    </div>
    <div class="field">
      <label for="f-client">Client</label>
      <select id="f-client"><option value="">All clients</option></select>
    </div>
    <div class="field">
      <label for="f-contract">Contract</label>
      <select id="f-contract"><option value="">All contracts</option></select>
    </div>
    <div class="field">
      <label for="f-filename">File Name</label>
      <input type="text" id="f-filename" placeholder="contains&hellip;">
    </div>
    <div class="field">
      <label for="f-ldfrom">Logged From</label>
      <input type="date" id="f-ldfrom">
    </div>
    <div class="field">
      <label for="f-ldto">Logged To</label>
      <input type="date" id="f-ldto">
    </div>
    <div class="field check">
      <input type="checkbox" id="f-unconf">
      <label for="f-unconf">Unconfigured only</label>
    </div>
    <button class="secondary" id="btn-reset">Reset</button>
  </section>

  <table id="grid">
    <thead>
      <tr>
        <th data-key="client">Client<span class="arrow">&#8597;</span></th>
        <th data-key="ftype">Type<span class="arrow">&#8597;</span></th>
        <th data-key="contract">Contract<span class="arrow">&#8597;</span></th>
        <th data-key="filename">File Name<span class="arrow">&#8597;</span></th>
        <th data-key="extracted">Extracted<span class="arrow">&#8597;</span></th>
        <th data-key="logged">Log Date<span class="arrow">&#8597;</span></th>
        <th data-key="feed">Feed<span class="arrow">&#8597;</span></th>
        <th data-key="size" class="num">File Size<span class="arrow">&#8597;</span></th>
      </tr>
    </thead>
    <tbody id="grid-body"></tbody>
  </table>
  <div class="pager">
    <span id="pager-info"></span>
    <button class="secondary" id="pg-prev">&laquo; Prev</button>
    <button class="secondary" id="pg-next">Next &raquo;</button>
  </div>
  </section>
</main>
<div id="tooltip" style="display:none"></div>

<script type="application/json" id="data">__DATA_JSON__</script>
<script>
(function() {
  const ROWS = JSON.parse(document.getElementById('data').textContent);
  const PAGE_SIZE = 100;
  let state = {
    ftype: '', client: '', contract: '', filename: '',
    ldfrom: '', ldto: '', unconf: false,
    sortKey: 'logged', sortDir: 'desc', page: 0,
  };

  const $ = (id) => document.getElementById(id);
  const fmtSize = (b) => {
    if (b == null || b === 0) return '0 B';
    const u = ['B', 'KB', 'MB', 'GB', 'TB'];
    let i = 0, n = b;
    while (n >= 1024 && i < u.length - 1) { n /= 1024; i++; }
    return (i === 0 ? n : n.toFixed(n < 10 ? 2 : 1)) + ' ' + u[i];
  };
  const ftPill = (t) => '<span class="ft ft-' + t + '">' + t + '</span>';

  const fillSelect = (id, values) => {
    for (const v of values) {
      const opt = document.createElement('option');
      opt.value = v; opt.textContent = v;
      $(id).appendChild(opt);
    }
  };
  fillSelect('f-client', [...new Set(ROWS.map(r => r.client).filter(Boolean))].sort());
  fillSelect('f-contract', [...new Set(ROWS.map(r => r.contract).filter(Boolean))].sort());

  function applyFilters() {
    const fnQ = state.filename.trim().toLowerCase();
    return ROWS.filter(r => {
      if (state.ftype && r.ftype !== state.ftype) return false;
      if (state.client && r.client !== state.client) return false;
      if (state.contract && r.contract !== state.contract) return false;
      if (state.unconf && !r.unconf) return false;
      if (fnQ && r.filename.toLowerCase().indexOf(fnQ) === -1) return false;
      const ld = r.logged ? r.logged.slice(0, 10) : '';
      if (state.ldfrom && (!ld || ld < state.ldfrom)) return false;
      if (state.ldto && (!ld || ld > state.ldto)) return false;
      return true;
    });
  }

  function sortRows(rows) {
    const k = state.sortKey;
    const dir = state.sortDir === 'asc' ? 1 : -1;
    const numeric = (k === 'size');
    return rows.slice().sort((a, b) => {
      let av = a[k], bv = b[k];
      if (numeric) { av = Number(av); bv = Number(bv); }
      if (av == null) av = '';
      if (bv == null) bv = '';
      if (av < bv) return -1 * dir;
      if (av > bv) return  1 * dir;
      return 0;
    });
  }

  function render() {
    const filtered = applyFilters();
    const sorted = sortRows(filtered);

    // KPIs
    let size = 0, latest = '', unconf = 0;
    const cset = new Set(), ctset = new Set();
    for (const r of filtered) {
      size += r.size || 0;
      if (r.client) cset.add(r.client);
      if (r.contract) ctset.add(r.contract);
      if (r.unconf) unconf++;
      if (r.logged && r.logged > latest) latest = r.logged;
    }
    $('kpi-files').textContent = filtered.length.toLocaleString();
    $('kpi-clients').textContent = cset.size.toLocaleString();
    $('kpi-contracts').textContent = ctset.size.toLocaleString();
    $('kpi-unconf').textContent = unconf.toLocaleString();
    $('kpi-size').textContent = fmtSize(size);
    $('kpi-latest').textContent = latest ? latest.slice(0, 10) : '—';

    // Pagination
    const pageCount = Math.max(1, Math.ceil(sorted.length / PAGE_SIZE));
    if (state.page >= pageCount) state.page = pageCount - 1;
    if (state.page < 0) state.page = 0;
    const start = state.page * PAGE_SIZE;
    const slice = sorted.slice(start, start + PAGE_SIZE);

    const body = $('grid-body');
    body.innerHTML = '';
    if (slice.length === 0) {
      const tr = document.createElement('tr');
      tr.innerHTML = '<td colspan="8" class="empty">No files match the current filters.</td>';
      body.appendChild(tr);
    } else {
      const esc = (s) => (s == null ? '' : String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'));
      for (const r of slice) {
        const tr = document.createElement('tr');
        if (r.unconf) tr.className = 'is-unconf';
        tr.innerHTML =
          '<td>' + esc(r.client) + '</td>' +
          '<td>' + ftPill(r.ftype) + '</td>' +
          '<td>' + esc(r.contract) + '</td>' +
          '<td>' + esc(r.filename) + (r.unconf ? '<span class="uc" title="File is unconfigured in RAMP">UNCONFIGURED</span>' : '') + '</td>' +
          '<td>' + esc(r.extracted) + '</td>' +
          '<td>' + esc(r.logged) + '</td>' +
          '<td>' + esc(r.feed) + '</td>' +
          '<td class="num' + (r.size ? '' : ' zero') + '">' + fmtSize(r.size) + '</td>';
        body.appendChild(tr);
      }
    }

    for (const th of document.querySelectorAll('#grid th[data-key]')) {
      th.classList.remove('sort-asc', 'sort-desc');
      const arrow = th.querySelector('.arrow');
      if (th.dataset.key === state.sortKey) {
        th.classList.add('sort-' + state.sortDir);
        arrow.textContent = state.sortDir === 'asc' ? '▲' : '▼';
      } else {
        arrow.textContent = '↕';
      }
    }

    $('pager-info').textContent = sorted.length === 0
      ? '0 files'
      : (start + 1).toLocaleString() + '–' + (start + slice.length).toLocaleString() + ' of ' + sorted.length.toLocaleString();
  }

  function bindFilters() {
    const ids = ['f-type', 'f-client', 'f-contract', 'f-filename', 'f-ldfrom', 'f-ldto'];
    const keys = ['ftype', 'client', 'contract', 'filename', 'ldfrom', 'ldto'];
    ids.forEach((id, i) => {
      $(id).addEventListener('input', () => {
        state[keys[i]] = $(id).value;
        state.page = 0;
        render();
      });
    });
    $('f-unconf').addEventListener('change', () => {
      state.unconf = $('f-unconf').checked;
      state.page = 0;
      render();
    });
    $('btn-reset').addEventListener('click', () => {
      ids.forEach(id => $(id).value = '');
      $('f-unconf').checked = false;
      Object.assign(state, { ftype: '', client: '', contract: '', filename: '', ldfrom: '', ldto: '', unconf: false, page: 0 });
      render();
    });
    document.querySelectorAll('#grid th[data-key]').forEach(th => {
      th.addEventListener('click', () => {
        const k = th.dataset.key;
        if (state.sortKey === k) {
          state.sortDir = state.sortDir === 'asc' ? 'desc' : 'asc';
        } else {
          state.sortKey = k;
          state.sortDir = (k === 'size' || k === 'logged' || k === 'extracted') ? 'desc' : 'asc';
        }
        render();
      });
    });
    $('pg-prev').addEventListener('click', () => { state.page--; render(); });
    $('pg-next').addEventListener('click', () => { state.page++; render(); });
  }

  bindFilters();
  render();

  // ===================== Monthly matrix view =====================
  const monthState = { year: '', ftype: 'COBC', client: '', contract: '', expanded: new Set() };

  const MONTH_ABBR = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

  // Distinct log-years (YYYY), newest first.
  const years = [...new Set(ROWS.map(r => r.logged).filter(Boolean).map(s => s.slice(0, 4)))].sort().reverse();
  for (const yy of years) {
    const opt = document.createElement('option');
    opt.value = yy; opt.textContent = yy;
    $('m-year').appendChild(opt);
  }
  monthState.year = years[0] || '';
  if (monthState.year) $('m-year').value = monthState.year;
  fillSelect('m-client', [...new Set(ROWS.map(r => r.client).filter(Boolean))].sort());

  // Contract dropdown is scoped to the selected client + type (cascading).
  function populateContracts() {
    const sel = $('m-contract');
    const keep = monthState.contract;
    sel.innerHTML = '<option value="">All contracts</option>';
    let src = ROWS;
    if (monthState.client) src = src.filter(r => r.client === monthState.client);
    if (monthState.ftype) src = src.filter(r => r.ftype === monthState.ftype);
    const cts = [...new Set(src.map(r => r.contract).filter(Boolean))].sort();
    for (const c of cts) {
      const opt = document.createElement('option');
      opt.value = c; opt.textContent = c;
      sel.appendChild(opt);
    }
    if (cts.indexOf(keep) !== -1) { sel.value = keep; }
    else { monthState.contract = ''; sel.value = ''; }
  }
  populateContracts();

  const tip = $('tooltip');
  let tipMap = {};   // cell id -> array of files
  const escT = (s) => (s == null ? '' : String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'));

  function showTip(evt, id) {
    const files = tipMap[id];
    if (!files) return;
    tip.innerHTML = files.map(f =>
      '<div class="tt-file">' +
        '<div class="tt-name">' + escT(f.filename) + (f.unconf ? ' <span class="tt-uc">&#9888; UNCONFIGURED</span>' : '') + '</div>' +
        '<div><span class="tt-label">Type:</span> ' + escT(f.ftype) + ' &middot; <span class="tt-label">Feed:</span> ' + (escT(f.feed) || '&mdash;') + '</div>' +
        '<div><span class="tt-label">Date Extracted:</span> ' + (escT(f.extracted) || '&mdash;') + '</div>' +
        '<div><span class="tt-label">Log Date:</span> ' + escT((f.logged || '').slice(0, 10)) + '</div>' +
        '<div><span class="tt-label">File Size:</span> ' + fmtSize(f.size) + '</div>' +
      '</div>'
    ).join('');
    tip.style.display = 'block';
    moveTip(evt);
  }
  function moveTip(evt) {
    const pad = 14, w = tip.offsetWidth, h = tip.offsetHeight;
    let x = evt.clientX + pad, y = evt.clientY + pad;
    if (x + w > window.innerWidth) x = evt.clientX - w - pad;
    if (y + h > window.innerHeight) y = evt.clientY - h - pad;
    tip.style.left = Math.max(4, x) + 'px';
    tip.style.top = Math.max(4, y) + 'px';
  }
  const hideTip = () => { tip.style.display = 'none'; };

  // Universe of client -> [contracts] for the current type filter, so every
  // client/contract is listed each month whether or not it logged a file.
  function buildPairs() {
    const pairs = {};
    const tmp = {};
    for (const r of ROWS) {
      if (monthState.ftype && r.ftype !== monthState.ftype) continue;
      (tmp[r.client] = tmp[r.client] || new Set()).add(r.contract);
    }
    for (const k of Object.keys(tmp)) pairs[k] = [...tmp[k]].sort((a, b) => a.localeCompare(b));
    return pairs;
  }

  let cid = 0;   // unique cell/tooltip id counter across a render
  function monthCells(monthMap) {
    let cells = '';
    for (let mo = 1; mo <= 12; mo++) {
      const files = monthMap[mo];
      if (files && files.length) {
        const id = 'x' + (cid++);
        tipMap[id] = files;
        const hasU = files.some(f => f.unconf);
        cells += '<td class="day hit" data-tip="' + id + '"><span class="mk' + (hasU ? ' has-unconf' : '') + '">' + files.length.toLocaleString('en-US') + '</span></td>';
      } else {
        cells += '<td class="day"></td>';
      }
    }
    return cells;
  }

  function renderMatrix() {
    const yr = monthState.year;
    const head = $('matrix-head'), body = $('matrix-body');
    head.innerHTML = ''; body.innerHTML = ''; tipMap = {}; cid = 0;
    if (!yr) { body.innerHTML = '<tr><td class="mx-empty">No data.</td></tr>'; return; }

    // Header: Client | Contract | Jan..Dec
    const htr = document.createElement('tr');
    let hh = '<th class="rowhdr c-client">Client</th><th class="rowhdr c-contract">Contract</th>';
    for (let mo = 1; mo <= 12; mo++) hh += '<th>' + MONTH_ABBR[mo - 1] + '</th>';
    htr.innerHTML = hh;
    head.appendChild(htr);

    const ALL_PAIRS = buildPairs();

    // clients[name] = { contracts: {ct: {month:[files]}}, agg: {month:[files]} }
    const clients = {};
    for (const r of ROWS) {
      if (!r.logged || r.logged.slice(0, 4) !== yr) continue;
      if (monthState.ftype && r.ftype !== monthState.ftype) continue;
      const c = clients[r.client] || (clients[r.client] = { contracts: {}, agg: {} });
      const ct = c.contracts[r.contract] || (c.contracts[r.contract] = {});
      const mo = Number(r.logged.slice(5, 7));
      (ct[mo] = ct[mo] || []).push(r);
      (c.agg[mo] = c.agg[mo] || []).push(r);
    }

    let clientNames = Object.keys(ALL_PAIRS);
    if (monthState.client) clientNames = clientNames.filter(n => n === monthState.client);
    if (monthState.contract) clientNames = clientNames.filter(n => ALL_PAIRS[n].indexOf(monthState.contract) !== -1);
    clientNames = clientNames.sort((a, b) => a.localeCompare(b));

    if (clientNames.length === 0) {
      body.innerHTML = '<tr><td class="mx-empty" colspan="14">No clients match the current filters.</td></tr>';
      return;
    }

    const forceOpen = !!monthState.contract;

    for (const name of clientNames) {
      const c = clients[name] || { contracts: {}, agg: {} };
      const open = forceOpen || monthState.expanded.has(name);
      const tri = open ? '&#9660;' : '&#9654;';   // ▼ / ▶

      const ctr = document.createElement('tr');
      ctr.className = 'client-row';
      ctr.dataset.client = name;
      ctr.innerHTML =
        '<td class="rowhdr c-client"><span class="tri">' + tri + '</span>' + escT(name) + '</td>' +
        '<td class="rowhdr c-contract"></td>' +
        monthCells(c.agg);
      body.appendChild(ctr);

      let contracts = ALL_PAIRS[name].slice();
      if (monthState.contract) contracts = contracts.filter(ct => ct === monthState.contract);
      for (const ct of contracts) {
        const rtr = document.createElement('tr');
        rtr.className = 'contract-row';
        rtr.dataset.client = name;
        if (!open) rtr.style.display = 'none';
        rtr.innerHTML =
          '<td class="rowhdr c-client"></td>' +
          '<td class="rowhdr c-contract">' + (escT(ct) || '<span style="color:var(--muted)">(none)</span>') + '</td>' +
          monthCells(c.contracts[ct] || {});
        body.appendChild(rtr);
      }
    }
  }

  $('matrix-body').addEventListener('mouseover', (e) => {
    const td = e.target.closest('td.hit');
    if (td && td.dataset.tip) showTip(e, td.dataset.tip);
  });
  $('matrix-body').addEventListener('mousemove', (e) => {
    if (tip.style.display === 'block') moveTip(e);
  });
  $('matrix-body').addEventListener('mouseout', (e) => {
    const to = e.relatedTarget;
    if (!to || !to.closest || !to.closest('td.hit')) hideTip();
  });

  $('matrix-body').addEventListener('click', (e) => {
    const row = e.target.closest('tr.client-row');
    if (!row) return;
    const name = row.dataset.client;
    if (monthState.expanded.has(name)) monthState.expanded.delete(name);
    else monthState.expanded.add(name);
    renderMatrix();
  });

  $('m-year').addEventListener('change', () => { monthState.year = $('m-year').value; renderMatrix(); });
  $('m-client').addEventListener('change', () => {
    monthState.client = $('m-client').value;
    populateContracts();
    renderMatrix();
  });
  $('m-contract').addEventListener('change', () => { monthState.contract = $('m-contract').value; renderMatrix(); });
  $('m-expand').addEventListener('click', () => {
    document.querySelectorAll('#matrix-body tr.client-row').forEach(r => monthState.expanded.add(r.dataset.client));
    renderMatrix();
  });
  $('m-collapse').addEventListener('click', () => { monthState.expanded.clear(); renderMatrix(); });

  // Tab switching.
  document.querySelectorAll('.tab').forEach(t => {
    t.addEventListener('click', () => {
      document.querySelectorAll('.tab').forEach(x => x.classList.remove('active'));
      document.querySelectorAll('.view').forEach(x => x.classList.remove('active'));
      t.classList.add('active');
      $('view-' + t.dataset.view).classList.add('active');
      hideTip();
      if (t.dataset.view === 'monthly') {
        monthState.ftype = t.dataset.ftype || '';
        monthState.expanded.clear();
        populateContracts();
        renderMatrix();
      }
    });
  });

  renderMatrix();
})();
</script>
</body>
</html>
"""


def generate_html(rows) -> str:
    unconf_n = sum(1 for r in rows if r.get("unconf"))
    return (HTML_TEMPLATE
            .replace("__GENERATED__", datetime.now().strftime("%Y-%m-%d %H:%M"))
            .replace("__SINCE__", START_DATE.strftime("%Y-%m-%d"))
            .replace("__ROW_COUNT__", f"{len(rows):,}")
            .replace("__UNCONF_COUNT__", f"{unconf_n:,}")
            .replace("__DATA_JSON__", json.dumps(rows, separators=(",", ":"))))


def main():
    full = "--full" in sys.argv
    mode = "FULL rebuild" if full else f"rolling {WINDOW_DAYS}-day window"
    print(f"[info] Building COBC/TRR/ABII report from RAMP ({RAMP_BASE}) "
          f"since {START_DATE} [{mode}]")
    if full:
        rows = build_rows_full()
        _save_cache(rows)
    else:
        rows = build_rows_rolling()
    by_type = {}
    for r in rows:
        by_type[r["ftype"]] = by_type.get(r["ftype"], 0) + 1
    unconf_n = sum(1 for r in rows if r.get("unconf"))
    print(f"[info] {len(rows)} files — " + ", ".join(f"{k}={v}" for k, v in sorted(by_type.items()))
          + f", unconfigured={unconf_n}")
    html = generate_html(rows)

    primary = OUTPUT_PATHS[0]
    try:
        os.makedirs(os.path.dirname(primary), exist_ok=True)
        with open(primary, "w", encoding="utf-8") as f:
            f.write(html)
        print(f"[done] Wrote {primary}")
    except (PermissionError, OSError) as e:
        print(f"[warn] Couldn't write primary path: {e}")

    for path in OUTPUT_PATHS[1:]:
        try:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            if os.path.exists(primary):
                shutil.copyfile(primary, path)
            else:
                with open(path, "w", encoding="utf-8") as f:
                    f.write(html)
            print(f"[done] Copy: {path}")
        except (PermissionError, OSError) as e:
            print(f"[warn] Couldn't write {path}: {e}")


if __name__ == "__main__":
    main()
