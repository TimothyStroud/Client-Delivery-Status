#!/usr/bin/env python
r"""CMSE Dashboard - MMSEA response-file loading, calendar + load list + reference.

Modeled on the "Paid Dates - Medical" static-HTML dashboards and on the
MMSEA-2026 tab of
  \\trgfile1\Shared\DIG\Data Business Delivery Team\Delivery Schedule\2026\ClientTracker.2026.xlsx

Three tabs:
  1. Calendar  - Client / File Type / Frequency / Handling + Jan..Dec grid, one
                 cell per month holding the ADO ticket number(s) of the loads
                 that landed that month ("E" = expected, carried over from the
                 tracker; pink = tracker outreach/late date).
  2. Loads     - every cmse_new..SourceLog row for the in-scope SourceIds,
                 SourceLogId desc, expandable to the per-file ImportStaging
                 breakdown (EntryName / DNDispositionCode / StagingStatus /
                 record counts).
  3. Reference - the File Type / Display Length table that sits to the right of
                 the calendar on the Excel tab, plus the SourceId->SourceName
                 key, the Client/ClientId key and the legend.

Data:
  TRGRepSQL3 / cmse_new    Source, SourceLog, ImportStaging, Client, ClientCodeLookup
  ADO / TFS                MMSEA work items, matched to a load by its PCN
                           (SourceLog.ProductionControlId) appearing in the
                           work-item description.

Everything expensive (the ImportStaging rollup, the per-PCN ADO lookups) is
cached in cmse_report_cache.json and only re-queried for new / recent rows, so
the 5am weekday refresh is cheap.  `--full` rebuilds from scratch.

Manual run:  python C:\Users\tls2\.claude\projects\H--\cmse_report.py
"""

import json
import os
import re
import shutil
import subprocess
import sys
import time
from collections import defaultdict, OrderedDict
from datetime import datetime, timedelta

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, r"H:\\")
try:
    from report_export import EXPORT_CSS, EXPORT_JS, EXPORT_UI
except ImportError:  # export dropdown is optional
    EXPORT_CSS = EXPORT_JS = EXPORT_UI = ""

SERVER = "TRGRepSQL3"
DATABASE = "cmse_new"
SOURCE_IDS = [4, 5, 6, 12, 13, 14, 15, 18, 19, 20]
WINDOW_START = "2025-01-01"

ADO_BASE = "https://devops.ado.rawlingslou.prod/TFS2012/Rawlings"
ADO_WEB = ADO_BASE + "/_workitems/edit/"

TRACKER = (r"\\trgfile1\Shared\DIG\Data Business Delivery Team\Delivery Schedule"
           r"\2026\ClientTracker.2026.xlsx")
TRACKER_SHEET = "MMSEA - 2026"
TRACKER_YEAR = 2026
LATE_FILL = "FFFFC7CE"          # pink "Late - Outreach Date" fill on the Excel tab

CACHE = os.path.join(HERE, "cmse_report_cache.json")

OUTPUT_PATHS = [
    r"\\trgfile1\Shared\DIG\Data Business Delivery Team\Delivery Schedule"
    r"\Daily Status Reports\CMSEReport.html",
    os.path.join(HERE, "CMSEReport.html"),
    r"C:\Users\tls2\OneDrive - Machinify\Documents\Reports\CMSEReport.html",
]

# SourceId -> (tracker row label, file-format spec label)
SOURCE_TYPE = {
    4:  ("HEW",  "HEW"),
    5:  ("MSP",  "MSP"),
    6:  ("NMSP", "Non-MSP"),
    12: ("NMSP", "Aetna NMSP"),
    13: ("NMSP", "Aetna NMSP"),
    14: ("MSP",  "Aetna MSP"),
    15: ("MSP",  "Aetna MSP"),
    18: ("HEW",  "Aetna HEW"),
    19: ("NMSP", "Aetna NMSP"),
    20: ("MSP",  "Aetna MSP"),
}

# File Type / Display Length table from the right of the Excel calendar (S1:X8)
FILE_SPECS = [
    ("HEW",        302,  300,  "Yes - %HO%",    "38",   "29"),
    ("MSP",        802,  800,  "Yes - %MSPR%",  "51",   "175"),
    ("Non-MSP",    502,  500,  "Yes",           "215",  "5"),
    ("Aetna HEW",  1203, 1201, "None",          "825",  "-"),
    ("Aetna MSP",  2050, 2048, "None",          "1244", "-"),
    ("Aetna NMSP", 1879, 1877, "None",          "1515", "-"),
]

# tracker client label -> cmse_new Client.ClientName
TRACKER_ALIAS = {
    "aetna": "Aetna",
    "bcbsfl": "BCBSFL",
    "bcbsks": "BCBSKS",
    "bcbsnc": "BCBS NC",
    "bcbssc": "BCBS SC",
    "bsca": "BSCA_FACETS",
    "carefirst": "CareFirst",
    "cigna": "Cigna",
    "emblem": "Emblem",
    "excellus": "Excellus",
    "harvardpilgrim": "HARVARDP",
    "healthnet": "HealthNet",
    "hne": "HealthNewEngland",
    "jhhc": "JohnsHopkins",
    "kaiser": "Kaiser",
    "kaiser_wa": "Kaiser_WA",
    "medica": "Medica",
    "mmoh": "Medical Mutual of Ohio",
    "oscar": "Oscar",
    "premera": "Premera Blue Cross",
    "tufts": "TUFTS",
}

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


# --------------------------------------------------------------------------- #
# SQL
# --------------------------------------------------------------------------- #

def sql(query, timeout=900):
    """Run a query via sqlcmd, return a list of column lists (no header)."""
    p = subprocess.run(
        ["sqlcmd", "-S", SERVER, "-d", DATABASE, "-E", "-W", "-w", "65535",
         "-s", "|", "-h", "-1", "-Q", "SET NOCOUNT ON; " + query],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=timeout)
    if p.returncode != 0:
        raise RuntimeError("sqlcmd failed: %s" % ((p.stderr or p.stdout or "")[:500]))
    out = []
    for line in (p.stdout or "").splitlines():
        line = line.rstrip("\r")
        if not line or line.startswith("(") or set(line) <= set("-|"):
            continue
        # char columns come back space-padded; NULL prints as the literal "NULL"
        out.append([("" if c.strip() == "NULL" else c.strip()) for c in line.split("|")])
    return out


def fetch_sources():
    rows = sql("SELECT SourceId, SourceName, ISNULL(SourceDescription,'') "
               "FROM dbo.Source ORDER BY SourceId")
    return [{"id": int(r[0]), "name": r[1], "desc": r[2]} for r in rows if r[0].isdigit()]


def fetch_clients():
    rows = sql("""
        SELECT C.ClientId, C.ClientName,
               STUFF((SELECT ', ' + X.SMARTClientCode
                        FROM dbo.ClientCodeLookup X (NOLOCK)
                       WHERE X.ClientId = C.ClientId
                    ORDER BY X.SMARTClientCode FOR XML PATH('')), 1, 2, '')
          FROM dbo.Client C (NOLOCK)
         WHERE C.ClientId IN (SELECT DISTINCT ClientId FROM dbo.SourceLog (NOLOCK)
                               WHERE SourceId IN (%s) AND ImportStartDate >= '%s')
      ORDER BY C.ClientName
    """ % (",".join(map(str, SOURCE_IDS)), WINDOW_START))
    return [{"id": int(r[0]), "name": r[1], "codes": (r[2] if len(r) > 2 else "")}
            for r in rows if r[0].isdigit()]


def fetch_loads():
    rows = sql("""
        SELECT L.SourceLogId, L.SourceId, ISNULL(L.ClientId,0), ISNULL(C.ClientName,''),
               L.EntryName,
               CONVERT(varchar(19), L.ImportStartDate, 120),
               ISNULL(CONVERT(varchar(19), L.ImportCompleteDate, 120), ''),
               ISNULL(L.RecordCount,0), ISNULL(L.ImportSuccessCount,0),
               ISNULL(L.ImportFailedcount,0),
               ISNULL(L.EntitlementAgeCount,0), ISNULL(L.EntitlementDisabilityCount,0),
               ISNULL(L.EntitlementEsrdCount,0),
               ISNULL(L.ProductionControlId,'')
          FROM dbo.SourceLog L (NOLOCK)
          LEFT JOIN dbo.Client C (NOLOCK) ON C.ClientId = L.ClientId
         WHERE L.SourceId IN (%s) AND L.ImportStartDate >= '%s'
      ORDER BY L.SourceLogId DESC
    """ % (",".join(map(str, SOURCE_IDS)), WINDOW_START))
    loads = []
    for r in rows:
        if not r[0].isdigit():
            continue
        loads.append({
            "sl": int(r[0]), "src": int(r[1]), "cid": int(r[2]), "client": r[3],
            "entry": r[4], "start": r[5], "done": r[6],
            "rec": int(r[7]), "ok": int(r[8]), "bad": int(r[9]),
            "age": int(r[10]), "dis": int(r[11]), "esrd": int(r[12]),
            "pcn": r[13],
        })
    return loads


def fetch_staging(source_log_ids):
    """ImportStaging rollup for the given SourceLogIds -> {sl: [[dn, status, n], ...]}."""
    out = defaultdict(list)
    ids = sorted(source_log_ids)
    for i in range(0, len(ids), 400):
        chunk = ids[i:i + 400]
        rows = sql("""
            SELECT I.SourceLogId, ISNULL(I.DNDispositionCode,''),
                   ISNULL(CONVERT(varchar(12), I.StagingStatus),''), COUNT(*)
              FROM dbo.ImportStaging I (NOLOCK)
             WHERE I.SourceLogId IN (%s)
          GROUP BY I.SourceLogId, I.DNDispositionCode, I.StagingStatus
          ORDER BY I.SourceLogId, I.DNDispositionCode, I.StagingStatus
        """ % ",".join(map(str, chunk)), timeout=1800)
        for r in rows:
            if r[0].isdigit():
                out[int(r[0])].append([r[1], r[2], int(r[3])])
    return out


# --------------------------------------------------------------------------- #
# ADO
# --------------------------------------------------------------------------- #

def _curl(args, timeout=180):
    p = subprocess.run(["curl.exe", "-s", "--negotiate", "-u", ":"] + args,
                       capture_output=True, text=True, encoding="utf-8",
                       errors="replace", timeout=timeout)
    if not p.stdout:
        raise RuntimeError("empty response from ADO")
    return json.loads(p.stdout.lstrip("\ufeff"))


def ado_find_pcn(pcn):
    """Work item ids whose description mentions this PCN, lowest (original) first."""
    q = ("SELECT [System.Id] FROM WorkItems WHERE [System.TeamProject]='Rawlings' "
         "AND [System.Description] CONTAINS '%s'" % pcn.replace("'", "''"))
    try:
        d = _curl([ADO_BASE + "/_apis/wit/wiql?api-version=5.0",
                   "-H", "Content-Type: application/json",
                   "-d", json.dumps({"query": q})])
    except Exception:
        return []
    return sorted(w["id"] for w in d.get("workItems", []))


def ado_details(ids):
    """{id: {title, state, tags}} for the given work item ids."""
    out = {}
    ids = sorted(set(ids))
    for i in range(0, len(ids), 180):
        chunk = ids[i:i + 180]
        try:
            d = _curl([ADO_BASE + "/_apis/wit/workitems?ids=" + ",".join(map(str, chunk)) +
                       "&fields=System.Id,System.Title,System.State,System.Tags"
                       "&api-version=5.0&errorPolicy=omit"], timeout=300)
        except Exception as e:
            print("[warn] ADO detail batch failed: %s" % e)
            continue
        for w in d.get("value", []) or []:
            if not w:
                continue
            f = w.get("fields", {})
            out[w["id"]] = {"title": f.get("System.Title", ""),
                            "state": f.get("System.State", ""),
                            "tags": f.get("System.Tags", "") or ""}
    return out


# --------------------------------------------------------------------------- #
# Tracker (Excel)
# --------------------------------------------------------------------------- #

def _norm(s):
    return re.sub(r"[^a-z0-9_]", "", (s or "").strip().lower())


def read_tracker():
    """-> ({(client, filetype): {"freq","hand","cells":{month:{"v","late"}}}}, note)."""
    try:
        import openpyxl
    except ImportError:
        return {}, "openpyxl not installed - tracker overlay skipped"
    try:
        wb = openpyxl.load_workbook(TRACKER, data_only=True, read_only=False)
        ws = wb[TRACKER_SHEET]
    except Exception as e:
        return {}, "tracker unavailable (%s)" % type(e).__name__

    out = OrderedDict()
    client = ftype = ""
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=17):
        a = row[0].value
        b = row[1].value
        if a is not None and str(a).strip():
            if str(a).strip().lower().startswith("e ="):   # legend row
                break
            client = str(a).strip()
            ftype = ""
        if b is not None and str(b).strip():
            ftype = str(b).strip()
        if not client:
            continue
        label = TRACKER_ALIAS.get(_norm(client), client)
        key = (label, ftype or "—")
        ent = out.setdefault(key, {"freq": "", "hand": "", "cells": {}})
        if row[2].value and not ent["freq"]:
            ent["freq"] = str(row[2].value).strip()
        if row[3].value and not ent["hand"]:
            ent["hand"] = str(row[3].value).strip()
        for m in range(12):                       # cols F..Q -> idx 5..16
            c = row[5 + m]
            v = c.value
            if v is None or str(v).strip() == "":
                continue
            if isinstance(v, datetime):
                txt = v.strftime("%m/%d/%Y")
            elif isinstance(v, float) and v == int(v):
                txt = str(int(v))
            else:
                txt = str(v).strip()
            fill = c.fill
            rgb = ""
            try:
                if fill is not None and fill.patternType:
                    rgb = str(fill.fgColor.rgb or "")
            except Exception:
                rgb = ""
            slot = ent["cells"].setdefault(m + 1, {"v": [], "late": False})
            slot["v"].append(txt)
            if rgb.upper() == LATE_FILL:
                slot["late"] = True
    for ent in out.values():
        for slot in ent["cells"].values():
            slot["v"] = " / ".join(slot["v"])
    return out, ""


# --------------------------------------------------------------------------- #
# Cache
# --------------------------------------------------------------------------- #

def load_cache(full):
    if full or not os.path.exists(CACHE):
        return {"staging": {}, "pcn": {}, "pcn_checked": {}, "wi": {}}
    try:
        with open(CACHE, "r", encoding="utf-8") as f:
            c = json.load(f)
    except Exception:
        return {"staging": {}, "pcn": {}, "pcn_checked": {}, "wi": {}}
    for k in ("staging", "pcn", "pcn_checked", "wi"):
        c.setdefault(k, {})
    return c


def save_cache(cache):
    tmp = CACHE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(cache, f, separators=(",", ":"))
    os.replace(tmp, CACHE)


# --------------------------------------------------------------------------- #
# Build
# --------------------------------------------------------------------------- #

def build(full=False):
    cache = load_cache(full)
    now = datetime.now()

    print("[info] %s / %s" % (SERVER, DATABASE))
    sources = fetch_sources()
    clients = fetch_clients()
    loads = fetch_loads()
    print("[info] %d loads, %d clients" % (len(loads), len(clients)))

    # -- ImportStaging: everything uncached, plus anything loaded in the last 30
    #    days (a recent load can still be re-summarized).
    recent_cut = (now - timedelta(days=30)).strftime("%Y-%m-%d")
    want = {l["sl"] for l in loads
            if str(l["sl"]) not in cache["staging"] or l["start"][:10] >= recent_cut}
    if want:
        print("[info] ImportStaging rollup for %d SourceLogIds..." % len(want))
        t0 = time.time()
        got = fetch_staging(want)
        for sl in want:
            cache["staging"][str(sl)] = got.get(sl, [])
        print("[info]   %.0fs" % (time.time() - t0))

    # -- ADO: resolve each PCN once.  Unresolved PCNs are retried while the load
    #    is still young (a ticket may be written after the file lands).
    todo = []
    for l in loads:
        # ProductionControlId is sometimes suffixed ("1079229_1", "801765-2");
        # ADO descriptions carry the base number, so key on the leading digits.
        m = re.match(r"\d+", l["pcn"].strip())
        pcn = m.group(0) if m else ""
        l["pcnk"] = pcn
        if not pcn:
            continue
        if pcn in cache["pcn"]:
            continue
        last = cache["pcn_checked"].get(pcn)
        if last and l["start"][:10] < (now - timedelta(days=45)).strftime("%Y-%m-%d"):
            continue                       # old and still ticket-less: stop asking
        todo.append(pcn)
    todo = sorted(set(todo))
    if todo:
        print("[info] ADO lookup for %d PCNs..." % len(todo))
        t0 = time.time()
        for n, pcn in enumerate(todo, 1):
            hits = ado_find_pcn(pcn)
            if hits:
                cache["pcn"][pcn] = hits
            cache["pcn_checked"][pcn] = now.strftime("%Y-%m-%d")
            if n % 50 == 0:
                print("[info]   %d/%d" % (n, len(todo)))
        print("[info]   %.0fs" % (time.time() - t0))

    # -- work item titles/states for everything we matched
    wanted_wi = set()
    for l in loads:
        for wi in cache["pcn"].get(l.get("pcnk") or "", []):
            wanted_wi.add(wi)
    missing_wi = [w for w in wanted_wi if str(w) not in cache["wi"]]
    if missing_wi:
        print("[info] ADO details for %d work items..." % len(missing_wi))
        for wi, det in ado_details(missing_wi).items():
            cache["wi"][str(wi)] = det
    # refresh state for anything not yet Closed
    stale = [int(w) for w in cache["wi"]
             if int(w) in wanted_wi and cache["wi"][w].get("state") not in ("Closed", "Removed")]
    if stale:
        for wi, det in ado_details(stale).items():
            cache["wi"][str(wi)] = det

    save_cache(cache)

    # -- attach ADO + staging to each load, prefer an MMSEA-tagged ticket
    for l in loads:
        hits = cache["pcn"].get(l.get("pcnk") or "", [])
        pick = None
        for wi in hits:
            det = cache["wi"].get(str(wi), {})
            blob = (det.get("tags", "") + " " + det.get("title", "")).lower()
            if "mmsea" in blob:
                pick = wi
                break
        if pick is None and hits:
            pick = hits[0]
        l["wi"] = pick
        det = cache["wi"].get(str(pick), {}) if pick else {}
        l["wit"] = det.get("title", "")
        l["wis"] = det.get("state", "")
        l["stg"] = [[(dn or "").strip(), (st or "").strip(), n]
                    for dn, st, n in cache["staging"].get(str(l["sl"]), [])]
        l["ft"] = SOURCE_TYPE.get(l["src"], ("?", "?"))[0]
        l["spec"] = SOURCE_TYPE.get(l["src"], ("?", "?"))[1]
        l["file"] = l["entry"].rsplit("\\", 1)[-1]

    tracker, tracker_note = read_tracker()
    if tracker_note:
        print("[warn] %s" % tracker_note)

    # -- calendar rows: union of what CMSE has loaded and what the tracker lists
    cal = OrderedDict()

    def cal_row(client, ftype, cid):
        key = (client, ftype)
        if key not in cal:
            t = tracker.get(key, {})
            cal[key] = {"client": client, "cid": cid, "ft": ftype,
                        "freq": t.get("freq", ""), "hand": t.get("hand", ""),
                        "m": {}, "tr": {str(k): v for k, v in t.get("cells", {}).items()}}
        elif cid and not cal[key]["cid"]:
            cal[key]["cid"] = cid
        return cal[key]

    name_by_id = {c["id"]: c["name"] for c in clients}
    for l in loads:
        client = l["client"] or name_by_id.get(l["cid"], "(unknown %d)" % l["cid"])
        row = cal_row(client, l["ft"], l["cid"])
        ym = l["start"][:7]
        row["m"].setdefault(ym, []).append(
            {"sl": l["sl"], "wi": l["wi"], "pcn": l["pcn"], "d": l["start"][8:10],
             "rec": l["rec"], "file": l["file"], "src": l["src"]})
    for (client, ftype) in tracker:
        cal_row(client, ftype, 0)

    id_by_name = {c["name"]: c["id"] for c in clients}
    for row in cal.values():
        if not row["cid"]:
            row["cid"] = id_by_name.get(row["client"], 0)

    years = sorted({l["start"][:4] for l in loads} | {str(TRACKER_YEAR)})
    cal_rows = sorted(cal.values(), key=lambda r: (r["client"].lower(), r["ft"]))

    # status/disposition inventory for the reference tab
    dn_ct, st_ct = defaultdict(int), defaultdict(int)
    for l in loads:
        for dn, st, n in l["stg"]:
            dn_ct[dn or "(blank)"] += n
            st_ct[st or "(blank)"] += n

    return {
        "generated": now.strftime("%Y-%m-%d %H:%M"),
        "years": years,
        "trackerYear": str(TRACKER_YEAR),
        "trackerNote": tracker_note,
        "cal": cal_rows,
        "loads": loads,
        "sources": sources,
        "scope": SOURCE_IDS,
        "clients": clients,
        "specs": [list(s) for s in FILE_SPECS],
        "srcType": {str(k): list(v) for k, v in SOURCE_TYPE.items()},
        "dn": sorted(dn_ct.items(), key=lambda kv: -kv[1]),
        "st": sorted(st_ct.items(), key=lambda kv: -kv[1]),
        "adoBase": ADO_WEB,
    }


# --------------------------------------------------------------------------- #
# HTML
# --------------------------------------------------------------------------- #

HTML_TEMPLATE = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>CMSE Dashboard</title>
<style>
  :root {
    --bg:#f4f6f9; --card:#fff; --border:#e3e7ec; --text:#24292f; --muted:#656d76;
    --accent:#2c5f8a; --accent-dark:#1f3d5c; --band:#f7f9fb; --hover:#eef4fb;
    --ok:#e8f5e9; --ok-line:#2e7d32; --late:#ffc7ce; --late-text:#9c0006;
    --exp:#f1f3f5; --exp-text:#6b7280;
  }
  * { box-sizing:border-box; }
  body { margin:0; font-family:"Segoe UI",-apple-system,BlinkMacSystemFont,sans-serif;
         background:var(--bg); color:var(--text); font-size:13px; }
  header { background:var(--accent-dark); color:#39ff14; padding:14px 24px; }
  header h1 { margin:0; font-size:18px; font-weight:600; }
  header .meta { font-size:12px; opacity:.85; margin-top:4px; color:#dbe6f0; }
  main { padding:14px 24px 40px; }
  .controls { display:flex; flex-wrap:wrap; gap:12px; align-items:center; margin-bottom:12px; }
  .seg { display:inline-flex; border:1px solid var(--border); border-radius:6px;
         overflow:hidden; background:#fff; }
  .seg button { border:0; background:#fff; color:var(--accent); padding:6px 14px;
                cursor:pointer; font-size:13px; border-right:1px solid var(--border); }
  .seg button:last-child { border-right:0; }
  .seg button.active { background:var(--accent); color:#fff; }
  select, input[type=text] { border:1px solid var(--border); border-radius:6px;
                             padding:6px 9px; font-size:13px; background:#fff;
                             color:var(--text); }
  input#search { min-width:230px; }
  .grow { flex:1; }
  button.export { background:var(--accent); color:#fff; border:0; border-radius:6px;
                  padding:6px 14px; cursor:pointer; font-size:13px; }
  button.export:hover { background:var(--accent-dark); }
  button.link { background:none; border:0; color:var(--accent); cursor:pointer;
                font-size:13px; text-decoration:underline; padding:0; }
  .kpis { display:flex; flex-wrap:wrap; gap:10px; margin-bottom:12px; }
  .kpi { background:var(--card); border:1px solid var(--border); border-radius:8px;
         padding:8px 14px; min-width:112px; }
  .kpi .n { font-size:19px; font-weight:600; font-variant-numeric:tabular-nums; }
  .kpi .l { font-size:11px; color:var(--muted); text-transform:uppercase;
            letter-spacing:.4px; }
  .wrap { background:var(--card); border:1px solid var(--border); border-radius:8px;
          overflow:auto; display:block; max-width:100%; max-height:72vh; }
  table { border-collapse:collapse; width:auto; }
  th, td { padding:4px 9px; white-space:nowrap; text-align:left;
           border-bottom:1px solid var(--border); }
  thead th { position:sticky; top:0; background:#fff; z-index:3; font-weight:600;
             color:var(--muted); border-bottom:2px solid var(--border); font-size:11.5px;
             text-transform:uppercase; letter-spacing:.3px; }
  thead th.sortable { cursor:pointer; }
  thead th.sortable:hover { color:var(--accent); }
  tbody tr:nth-child(even) td { background:var(--band); }
  tbody tr:hover td { background:var(--hover); }
  td.num, th.num { text-align:right; font-variant-numeric:tabular-nums; }
  td.mid, th.mid { text-align:center; }
  a { color:var(--accent); }
  /* calendar */
  #cal td.mo { text-align:center; min-width:74px; padding:3px 6px; }
  #cal td.mo.has { background:var(--ok) !important; }
  #cal td.mo.late { background:var(--late) !important; color:var(--late-text);
                    font-weight:600; }
  #cal td.mo.exp { color:var(--exp-text); }
  #cal td.mo a { display:block; line-height:1.35; text-decoration:none; }
  #cal td.mo a:hover { text-decoration:underline; }
  #cal td.mo .noado { display:block; line-height:1.35; color:var(--muted); }
  #cal td.cli { position:sticky; left:0; background:#fff; z-index:2; font-weight:600; }
  #cal tbody tr:nth-child(even) td.cli { background:var(--band); }
  #cal tbody tr:hover td.cli { background:var(--hover); }
  .pill { display:inline-block; padding:1px 7px; border-radius:9px; font-size:11px;
          border:1px solid var(--border); background:#fff; }
  .pill.closed { background:#e8f5e9; border-color:#c8e6c9; color:#1b5e20; }
  .pill.open { background:#fff4e5; border-color:#ffe0b2; color:#8a5300; }
  .pill.none { background:#fdecea; border-color:#f8c9c4; color:#8a1c13; }
  tr.det td { background:#fbfcfd !important; padding:0; }
  tr.det table { margin:6px 0 10px 28px; width:auto; }
  tr.det th, tr.det td { border-bottom:1px solid #eef1f4; font-size:12px;
                         padding:2px 10px; }
  tr.det thead th { position:static; text-transform:none; letter-spacing:0; }
  tr.load td.exp { cursor:pointer; color:var(--accent); user-select:none; width:18px; }
  .legend { display:flex; flex-wrap:wrap; gap:16px; align-items:center;
            margin:10px 0 0; font-size:12px; color:var(--muted); }
  .sw { display:inline-block; width:13px; height:13px; border-radius:3px;
        border:1px solid var(--border); vertical-align:-2px; margin-right:5px; }
  .cards { display:flex; flex-wrap:wrap; gap:16px; align-items:flex-start; }
  .card { background:var(--card); border:1px solid var(--border); border-radius:8px;
          overflow:hidden; }
  .card h2 { margin:0; padding:9px 14px; font-size:13px; background:#f7f9fb;
             border-bottom:1px solid var(--border); }
  .card .body { padding:0; max-height:60vh; overflow:auto; }
  .card.wide { flex:1 1 560px; max-width:820px; }
  #clikey td:last-child { white-space:normal; line-height:1.5; color:var(--muted);
                          font-size:12px; }
  .card p.note { margin:0; padding:8px 14px; font-size:12px; color:var(--muted);
                 border-top:1px solid var(--border); }
  tr.dim td { color:var(--muted); }
  .empty { padding:26px; text-align:center; color:var(--muted); }
  [hidden] { display:none !important; }
__EXPORT_CSS__
</style>
</head>
<body>
<header>
  <h1>CMSE Dashboard &mdash; MMSEA Response Files</h1>
  <div class="meta" id="meta"></div>
</header>
<main>
  <div class="controls">
    <div class="seg" id="tabs">
      <button data-tab="cal" class="active">Calendar</button>
      <button data-tab="loads">Loads</button>
      <button data-tab="ref">File Types &amp; Keys</button>
    </div>
    <select id="year" class="cal-only"></select>
    <select id="client"></select>
    <select id="source" class="loads-only"></select>
    <select id="ticket" class="loads-only">
      <option value="">All loads</option>
      <option value="y">Has ADO ticket</option>
      <option value="n">No ADO ticket</option>
    </select>
    <input id="search" type="text" placeholder="Find client, file, PCN, ticket&hellip;">
    <button class="link" id="clear">Reset</button>
    <span class="grow"></span>
    __EXPORT_UI__
  </div>

  <div class="kpis" id="kpis"></div>

  <section id="view-cal">
    <div class="wrap"><table id="cal"><thead id="cal-head"></thead><tbody id="cal-body"></tbody></table></div>
    <div class="legend">
      <span><span class="sw" style="background:var(--ok)"></span>Ticket number = loaded to CMSE</span>
      <span><span class="sw" style="background:var(--exp)"></span>E = expected</span>
      <span><span class="sw" style="background:var(--late)"></span>Late &ndash; outreach date</span>
      <span id="cal-note"></span>
    </div>
  </section>

  <section id="view-loads" hidden>
    <div class="wrap"><table id="loads"><thead id="loads-head"></thead><tbody id="loads-body"></tbody></table></div>
    <div class="legend"><span>Click a row to expand its ImportStaging breakdown
      (EntryName / DNDispositionCode / StagingStatus / records).</span>
      <button class="link" id="expand-all">Expand all</button></div>
  </section>

  <section id="view-ref" hidden>
    <div class="cards">
      <div class="card"><h2>File Type &mdash; layout</h2><div class="body">
        <table id="specs"></table></div>
        <p class="note">From the MMSEA&nbsp;-&nbsp;2026 tab of ClientTracker.2026.xlsx.</p></div>
      <div class="card"><h2>SourceId key</h2><div class="body">
        <table id="srckey"></table></div>
        <p class="note">Bold rows are the SourceIds this dashboard reports on.</p></div>
      <div class="card wide"><h2>Client key</h2><div class="body">
        <table id="clikey"></table></div>
        <p class="note">ClientId / ClientName from cmse_new..Client, joined to
          ClientCodeLookup for the SMART client codes.</p></div>
      <div class="card"><h2>DNDispositionCode</h2><div class="body">
        <table id="dnkey"></table></div>
        <p class="note">Record counts across every load in window.</p></div>
      <div class="card"><h2>StagingStatus</h2><div class="body">
        <table id="stkey"></table></div></div>
    </div>
  </section>
</main>

<script type="application/json" id="data">__DATA_JSON__</script>
<script>__EXPORT_JS__</script>
<script>
(function () {
  const D = JSON.parse(document.getElementById('data').textContent);
  const MN = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const SRC = {}; D.sources.forEach(s => SRC[s.id] = s.name);
  const $ = id => document.getElementById(id);
  const esc = s => (s == null ? '' : String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;')
                                              .replace(/>/g,'&gt;').replace(/"/g,'&quot;'));
  const nf = n => (n == null ? '' : Number(n).toLocaleString('en-US'));

  const latestYear = D.years[D.years.length - 1];
  let S = { tab:'cal', year:latestYear, client:'', source:'', ticket:'', q:'',
            sortK:'sl', sortD:-1, open:new Set() };

  // ---- filter widgets -----------------------------------------------------
  $('year').innerHTML = D.years.map(y => `<option${y===S.year?' selected':''}>${y}</option>`).join('');
  const clientNames = [...new Set(D.cal.map(r => r.client).concat(D.loads.map(l => l.client)))]
                        .filter(Boolean).sort((a,b) => a.localeCompare(b));
  $('client').innerHTML = '<option value="">All clients</option>' +
    clientNames.map(c => `<option>${esc(c)}</option>`).join('');
  $('source').innerHTML = '<option value="">All source types</option>' +
    D.scope.map(id => `<option value="${id}">${id} &middot; ${esc(SRC[id]||'')}</option>`).join('');

  $('meta').textContent = 'Generated ' + D.generated + ' \u00b7 TRGRepSQL3 / CMSE_New \u00b7 '
    + 'SourceId ' + D.scope.join(', ') + ' \u00b7 ' + D.loads.length.toLocaleString('en-US')
    + ' loads from ' + (D.years[0] || '') + ' forward';
  $('cal-note').textContent = D.trackerNote ? ('\u26a0 ' + D.trackerNote) : '';

  // ---- helpers ------------------------------------------------------------
  const hay = l => [l.client, l.file, l.entry, l.pcn, l.wi, l.wit, l.sl, SRC[l.src], l.ft]
                     .join(' ').toLowerCase();

  function filteredLoads() {
    const q = S.q.toLowerCase();
    return D.loads.filter(l =>
      (!S.client || l.client === S.client) &&
      (!S.source || String(l.src) === S.source) &&
      (S.ticket !== 'y' || l.wi) && (S.ticket !== 'n' || !l.wi) &&
      (!q || hay(l).includes(q)));
  }

  function filteredCal() {
    const q = S.q.toLowerCase();
    return D.cal.filter(r => {
      if (S.client && r.client !== S.client) return false;
      if (S.source && SRC[S.source] && D.srcType[S.source] && D.srcType[S.source][0] !== r.ft) return false;
      if (!q) return true;
      if ((r.client + ' ' + r.ft + ' ' + r.freq + ' ' + r.hand).toLowerCase().includes(q)) return true;
      return Object.keys(r.m).some(ym => ym.startsWith(S.year) &&
        r.m[ym].some(x => (x.file + ' ' + x.pcn + ' ' + (x.wi||'')).toLowerCase().includes(q)));
    });
  }

  // ---- calendar -----------------------------------------------------------
  function renderCal() {
    const rows = filteredCal();
    $('cal-head').innerHTML = '<tr><th>Client</th><th class="mid">Client Id</th>' +
      '<th>File Type</th><th>Frequency</th><th>Handling</th>' +
      MN.map(m => `<th class="mid">${m}</th>`).join('') + '</tr>';

    if (!rows.length) { $('cal-body').innerHTML =
      '<tr><td colspan="17" class="empty">No rows match.</td></tr>'; return; }

    const isTrackerYear = S.year === D.trackerYear;
    const html = rows.map(r => {
      const cells = MN.map((_, i) => {
        const ym = S.year + '-' + String(i+1).padStart(2,'0');
        const loads = r.m[ym] || [];
        if (loads.length) {
          const seen = new Set();
          const inner = loads.map(x => {
            const t = `${x.file}\n${x.d} ${MN[i]} ${S.year} \u00b7 SourceLogId ${x.sl}`
                    + `\nPCN ${x.pcn||'-'} \u00b7 ${nf(x.rec)} records`;
            if (x.wi) {
              if (seen.has(x.wi)) return '';
              seen.add(x.wi);
              return `<a href="${D.adoBase}${x.wi}" target="_blank" title="${esc(t)}">${x.wi}</a>`;
            }
            return `<span class="noado" title="${esc(t)}">\u25aa ${x.sl}</span>`;
          }).join('');
          return `<td class="mo has">${inner}</td>`;
        }
        const tr = isTrackerYear ? r.tr[String(i+1)] : null;
        if (tr) return `<td class="mo ${tr.late?'late':'exp'}">${esc(tr.v)}</td>`;
        return '<td class="mo"></td>';
      }).join('');
      return `<tr><td class="cli">${esc(r.client)}</td><td class="mid">${r.cid||''}</td>` +
             `<td>${esc(r.ft)}</td><td>${esc(r.freq)}</td><td>${esc(r.hand)}</td>${cells}</tr>`;
    }).join('');
    $('cal-body').innerHTML = html;
  }

  // ---- loads --------------------------------------------------------------
  const LCOLS = [
    ['', '', null],
    ['sl', 'SourceLogId', 'num'],
    ['wi', 'ADO', 'mid'],
    ['pcn', 'PCN', 'mid'],
    ['client', 'Client Name', null],
    ['cid', 'Client Id', 'mid'],
    ['src', 'SourceId', 'mid'],
    ['srcname', 'Source Name', null],
    ['ft', 'File Type', null],
    ['file', 'Entry Name', null],
    ['start', 'Import Start', null],
    ['done', 'Import Complete', null],
    ['rec', 'Records', 'num'],
    ['ok', 'Success', 'num'],
    ['bad', 'Failed', 'num'],
    ['age', 'Age', 'num'],
    ['dis', 'Disability', 'num'],
    ['esrd', 'ESRD', 'num'],
    ['wis', 'Ticket State', null],
  ];

  function sortLoads(rows) {
    const k = S.sortK, d = S.sortD;
    const val = l => k === 'srcname' ? (SRC[l.src] || '') : (l[k] == null ? '' : l[k]);
    return rows.slice().sort((a, b) => {
      const x = val(a), y = val(b);
      if (typeof x === 'number' && typeof y === 'number') return (x - y) * d;
      return String(x).localeCompare(String(y), undefined, {numeric:true}) * d;
    });
  }

  function renderLoads() {
    const rows = sortLoads(filteredLoads());
    $('loads-head').innerHTML = '<tr>' + LCOLS.map(([k, lbl, cls]) =>
      k ? `<th class="${cls||''} sortable" data-k="${k}">${lbl}${S.sortK===k?(S.sortD>0?' \u25b2':' \u25bc'):''}</th>`
        : '<th></th>').join('') + '</tr>';

    if (!rows.length) { $('loads-body').innerHTML =
      `<tr><td colspan="${LCOLS.length}" class="empty">No loads match.</td></tr>`; return; }

    const out = [];
    for (const l of rows) {
      const open = S.open.has(l.sl);
      const wiCell = l.wi
        ? `<a href="${D.adoBase}${l.wi}" target="_blank" title="${esc(l.wit)}">${l.wi}</a>`
        : '<span class="pill none">none</span>';
      const st = l.wis ? `<span class="pill ${l.wis==='Closed'?'closed':'open'}">${esc(l.wis)}</span>` : '';
      out.push(`<tr class="load" data-sl="${l.sl}">` +
        `<td class="exp">${open?'\u2212':'+'}</td>` +
        `<td class="num">${l.sl}</td><td class="mid">${wiCell}</td>` +
        `<td class="mid">${esc(l.pcn)}</td><td>${esc(l.client)}</td>` +
        `<td class="mid">${l.cid||''}</td><td class="mid">${l.src}</td>` +
        `<td>${esc(SRC[l.src]||'')}</td><td>${esc(l.ft)}</td>` +
        `<td title="${esc(l.entry)}">${esc(l.file)}</td>` +
        `<td>${esc(l.start)}</td><td>${esc(l.done)}</td>` +
        `<td class="num">${nf(l.rec)}</td><td class="num">${nf(l.ok)}</td>` +
        `<td class="num">${nf(l.bad)}</td><td class="num">${nf(l.age)}</td>` +
        `<td class="num">${nf(l.dis)}</td><td class="num">${nf(l.esrd)}</td>` +
        `<td>${st}</td></tr>`);
      if (open) {
        const tot = l.stg.reduce((s, r) => s + r[2], 0);
        const body = l.stg.length
          ? l.stg.map(([dn, sts, n]) =>
              `<tr><td>${esc(l.file)}</td><td class="mid">${esc(dn)||'\u2014'}</td>` +
              `<td class="mid">${esc(sts)||'\u2014'}</td><td class="num">${nf(n)}</td></tr>`).join('')
            + `<tr class="dim"><td colspan="3"><b>Total</b></td><td class="num"><b>${nf(tot)}</b></td></tr>`
          : '<tr class="dim"><td colspan="4">No ImportStaging rows for this SourceLogId.</td></tr>';
        out.push(`<tr class="det"><td colspan="${LCOLS.length}"><table><thead><tr>` +
          '<th>EntryName</th><th class="mid">DNDispositionCode</th>' +
          '<th class="mid">StagingStatus</th><th class="num">Records</th>' +
          `</tr></thead><tbody>${body}</tbody></table></td></tr>`);
      }
    }
    $('loads-body').innerHTML = out.join('');
  }

  // ---- reference ----------------------------------------------------------
  function renderRef() {
    $('specs').innerHTML = '<thead><tr><th>File Type</th><th class="num">Display Length</th>' +
      '<th class="num">File Length</th><th>Header &amp; Trailer</th>' +
      '<th class="num">Entitlement Reason Position</th><th class="num">SSN Location</th></tr></thead><tbody>' +
      D.specs.map(s => `<tr><td><b>${esc(s[0])}</b></td><td class="num">${s[1]}</td>` +
        `<td class="num">${s[2]}</td><td>${esc(s[3])}</td>` +
        `<td class="num">${esc(s[4])}</td><td class="num">${esc(s[5])}</td></tr>`).join('') +
      '</tbody>';

    $('srckey').innerHTML = '<thead><tr><th class="num">SourceId</th><th>SourceName</th>' +
      '<th>File Type</th><th>Description</th></tr></thead><tbody>' +
      D.sources.map(s => {
        const inScope = D.scope.includes(s.id);
        const t = D.srcType[s.id];
        return `<tr class="${inScope?'':'dim'}"><td class="num">${inScope?'<b>'+s.id+'</b>':s.id}</td>` +
               `<td>${inScope?'<b>'+esc(s.name)+'</b>':esc(s.name)}</td>` +
               `<td>${t?esc(t[1]):''}</td><td>${esc(s.desc)}</td></tr>`;
      }).join('') + '</tbody>';

    $('clikey').innerHTML = '<thead><tr><th class="num">Client Id</th><th>Client Name</th>' +
      '<th>SMART Client Codes</th></tr></thead><tbody>' +
      D.clients.map(c => `<tr><td class="num">${c.id}</td><td><b>${esc(c.name)}</b></td>` +
        `<td>${esc(c.codes||'')}</td></tr>`).join('') + '</tbody>';

    const kv = (id, arr, lbl) => $(id).innerHTML =
      `<thead><tr><th>${lbl}</th><th class="num">Records</th></tr></thead><tbody>` +
      arr.map(([k, n]) => `<tr><td>${esc(k)}</td><td class="num">${nf(n)}</td></tr>`).join('') +
      '</tbody>';
    kv('dnkey', D.dn, 'DNDispositionCode');
    kv('stkey', D.st, 'StagingStatus');
  }

  // ---- KPIs ---------------------------------------------------------------
  function renderKpis() {
    let cards;
    if (S.tab === 'cal') {
      const rows = filteredCal();
      let loaded = 0, cells = 0;
      rows.forEach(r => MN.forEach((_, i) => {
        const ym = S.year + '-' + String(i+1).padStart(2,'0');
        if (r.m[ym]) { loaded++; cells += r.m[ym].length; }
      }));
      cards = [['Rows', rows.length], ['Months with a load', loaded],
               ['Files loaded ' + S.year, cells],
               ['Clients', new Set(rows.map(r => r.client)).size]];
    } else if (S.tab === 'loads') {
      const rows = filteredLoads();
      cards = [['Loads', rows.length],
               ['Records', rows.reduce((s, l) => s + l.rec, 0)],
               ['Failed', rows.reduce((s, l) => s + l.bad, 0)],
               ['With ADO ticket', rows.filter(l => l.wi).length],
               ['No ADO ticket', rows.filter(l => !l.wi).length],
               ['Clients', new Set(rows.map(l => l.client)).size]];
    } else {
      cards = [['File types', D.specs.length], ['Source types in scope', D.scope.length],
               ['Clients', D.clients.length]];
    }
    $('kpis').innerHTML = cards.map(([l, n]) =>
      `<div class="kpi"><div class="n">${nf(n)}</div><div class="l">${l}</div></div>`).join('');
  }

  // ---- render -------------------------------------------------------------
  function render() {
    $('view-cal').hidden = S.tab !== 'cal';
    $('view-loads').hidden = S.tab !== 'loads';
    $('view-ref').hidden = S.tab !== 'ref';
    document.querySelectorAll('.cal-only').forEach(e => e.hidden = S.tab !== 'cal');
    document.querySelectorAll('.loads-only').forEach(e => e.hidden = S.tab !== 'loads');
    $('client').hidden = S.tab === 'ref';
    $('search').hidden = S.tab === 'ref';
    renderKpis();
    if (S.tab === 'cal') renderCal();
    else if (S.tab === 'loads') renderLoads();
    else renderRef();
  }

  // ---- events -------------------------------------------------------------
  $('tabs').addEventListener('click', e => {
    const b = e.target.closest('button'); if (!b) return;
    S.tab = b.dataset.tab;
    [...$('tabs').children].forEach(x => x.classList.toggle('active', x === b));
    render();
  });
  $('year').addEventListener('change', e => { S.year = e.target.value; render(); });
  $('client').addEventListener('change', e => { S.client = e.target.value; render(); });
  $('source').addEventListener('change', e => { S.source = e.target.value; render(); });
  $('ticket').addEventListener('change', e => { S.ticket = e.target.value; render(); });
  let t; $('search').addEventListener('input', e => {
    clearTimeout(t); const v = e.target.value;
    t = setTimeout(() => { S.q = v; render(); }, 160);
  });
  $('clear').addEventListener('click', () => {
    S.client = S.source = S.ticket = S.q = ''; S.year = latestYear; S.open.clear();
    $('client').value = ''; $('source').value = ''; $('ticket').value = '';
    $('search').value = ''; $('year').value = latestYear; render();
  });
  $('loads-head').addEventListener('click', e => {
    const th = e.target.closest('th[data-k]'); if (!th) return;
    const k = th.dataset.k;
    if (S.sortK === k) S.sortD = -S.sortD; else { S.sortK = k; S.sortD = 1; }
    renderLoads();
  });
  $('loads-body').addEventListener('click', e => {
    const tr = e.target.closest('tr.load'); if (!tr) return;
    if (e.target.closest('a')) return;
    const sl = Number(tr.dataset.sl);
    if (S.open.has(sl)) S.open.delete(sl); else S.open.add(sl);
    renderLoads();
  });
  $('expand-all').addEventListener('click', () => {
    const rows = filteredLoads();
    if (rows.every(l => S.open.has(l.sl))) S.open.clear();
    else rows.forEach(l => S.open.add(l.sl));
    $('expand-all').textContent = S.open.size ? 'Collapse all' : 'Expand all';
    renderLoads();
  });

  // ---- export -------------------------------------------------------------
  function buildExport() {
    if (S.tab === 'cal') {
      const rows = filteredCal();
      const isT = S.year === D.trackerYear;
      return {
        name: 'CMSE_Calendar_' + S.year,
        title: 'CMSE Dashboard \u2014 MMSEA Calendar ' + S.year,
        subtitle: (S.client || 'All clients') + ' \u00b7 generated ' + D.generated,
        headers: ['Client', 'Client Id', 'File Type', 'Frequency', 'Handling', ...MN],
        rows: rows.map(r => [r.client, r.cid || '', r.ft, r.freq, r.hand,
          ...MN.map((_, i) => {
            const ym = S.year + '-' + String(i+1).padStart(2,'0');
            const l = r.m[ym];
            if (l) return [...new Set(l.map(x => x.wi || ('SL ' + x.sl)))].join(' / ');
            const tr = isT ? r.tr[String(i+1)] : null;
            return tr ? tr.v : '';
          })]),
        note: 'TRGRepSQL3 / CMSE_New \u00b7 ticket number = loaded to CMSE, E = expected',
        rowsPerSlide: 12, fontSz: 800,
      };
    }
    if (S.tab === 'loads') {
      const rows = sortLoads(filteredLoads());
      return {
        name: 'CMSE_Loads',
        title: 'CMSE Dashboard \u2014 SourceLog',
        subtitle: (S.client || 'All clients') + ' \u00b7 ' + rows.length +
                  ' loads \u00b7 generated ' + D.generated,
        headers: ['SourceLogId','ADO','PCN','Client Name','Client Id','SourceId',
                  'Source Name','File Type','EntryName','Import Start','Import Complete',
                  'Records','Success','Failed','Age','Disability','ESRD','Ticket State'],
        rows: rows.map(l => [l.sl, l.wi||'', l.pcn, l.client, l.cid, l.src,
          SRC[l.src]||'', l.ft, l.entry, l.start, l.done, l.rec, l.ok, l.bad,
          l.age, l.dis, l.esrd, l.wis||'']),
        note: 'cmse_new..SourceLog, SourceId ' + D.scope.join(', '),
        rowsPerSlide: 12, fontSz: 700,
      };
    }
    return {
      name: 'CMSE_FileTypes',
      title: 'CMSE Dashboard \u2014 File Types',
      subtitle: 'Generated ' + D.generated,
      headers: ['File Type','Display Length','File Length','Header & Trailer',
                'Entitlement Reason Position','SSN Location'],
      rows: D.specs.map(s => s.slice()),
      note: 'From ClientTracker.2026.xlsx, MMSEA - 2026 tab',
    };
  }
  if (window.RptExport) RptExport.wire('btn-export', 'exp-menu', buildExport);

  render();
})();
</script>
</body>
</html>
"""


def generate_html(data):
    return (HTML_TEMPLATE
            .replace("__EXPORT_CSS__", EXPORT_CSS)
            .replace("__EXPORT_UI__", EXPORT_UI)
            .replace("__EXPORT_JS__", EXPORT_JS)
            .replace("__DATA_JSON__", json.dumps(data, separators=(",", ":"))))


def main():
    full = "--full" in sys.argv
    data = build(full=full)
    html = generate_html(data)

    primary = OUTPUT_PATHS[0]
    written = None
    for path in OUTPUT_PATHS:
        try:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            if written and os.path.exists(written):
                shutil.copyfile(written, path)
            else:
                with open(path, "w", encoding="utf-8") as f:
                    f.write(html)
                written = path
            print("[done] %s" % path)
        except (PermissionError, OSError) as e:
            print("[warn] couldn't write %s: %s" % (path, e))
    if not written:
        raise SystemExit("[error] no output written")
    if written != primary:
        print("[warn] primary path %s was not written" % primary)


if __name__ == "__main__":
    main()
