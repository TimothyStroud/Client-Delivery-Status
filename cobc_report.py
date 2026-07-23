"""Generate a static HTML COBC / TRR / ABII File Report dashboard.

Self-contained single .html file: embedded JSON data + vanilla JS filtering.
No external dependencies, no server, no auth. Styled to match the
MissingPayDates / CAQH Traffic / MSPI reports. Drop it on a shared drive and
users open it in a browser.

Data source: TRGUtil10 / Ramp.ramp.FileLog.  This is the RAMP file-movement
log, so a single physical file is logged once per pipeline stage
(Completed / Staged / Downloaded / ...).  We therefore DEDUPE to one row per
file (one canonical row per FileName + received-date, preferring the
"Completed" stage) and keep only INBOUND CLIENT files (SourcePath under a
``\\Clients\\`` folder), which is where the client name is derivable.

    SELECT ... FROM ramp.FileLog
    WHERE DateCreated > '2024-12-31 23:59:00.000'
      AND ([FileName] LIKE '%MARX%'   -- COBC (MARXCOB)
        OR [FileName] LIKE '%TRR%'    -- TRR
        OR [FileName] LIKE '%ABII%')  -- ABII
      AND SourcePath LIKE '%\\Clients\\%'

Note: the OR group is parenthesised on purpose.  Written flat
(``... AND a OR b OR c``) SQL binds it as ``(AND a) OR b OR c`` and the date
filter would silently not apply to the TRR/ABII arms.

Each row carries FileSize (bytes).  Client comes from the SourcePath (the
segment after ``\\Clients\\``, skipping a ``users\\`` level, with staging
suffixes like ``_Temp`` / ``_Decrypt`` stripped and a few known aliases
folded, e.g. ``Machinify\\Inbound\\UHC_COBC`` -> UHC, ``BCBSArkansas`` ->
BCBSARRx).  File Type (COBC / TRR / ABII) is derived from the FileName.
Contract and the data date are parsed out of the leaf file name where present.

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
from datetime import datetime

SERVER = "TRGUtil10"
DATABASE = "Ramp"
SINCE = "2024-12-31 23:59:00.000"

# sqlcmd's -s only honors a single character, so use a tab (never present in
# these path / filename / date fields).
SEP = "\t"

# One canonical row per (FileName, received-date): prefer the Completed stage,
# then Staged/Downloaded, then anything else; tie-break on FileLogId.  Only
# inbound client files (SourcePath under \Clients\) so the client parses out.
SQL = """SET NOCOUNT ON;
WITH F AS (
  SELECT  FileName,
          SourcePath,
          FileSize,
          CONVERT(varchar(19), DateCreated, 120) AS Received,
          Status,
          ROW_NUMBER() OVER (
            PARTITION BY FileName, CAST(DateCreated AS date)
            ORDER BY CASE Status
                       WHEN 'Completed'  THEN 0
                       WHEN 'Staged'     THEN 1
                       WHEN 'Downloaded' THEN 2
                       ELSE 3 END,
                     FileLogId
          ) AS rn
  FROM ramp.FileLog
  WHERE DateCreated > '{since}'
    AND ([FileName] LIKE '%MARX%' OR [FileName] LIKE '%TRR%' OR [FileName] LIKE '%ABII%')
    AND SourcePath LIKE '%\\Clients\\%'
)
SELECT FileName, SourcePath, FileSize, Received, Status
FROM F
WHERE rn = 1
ORDER BY Received DESC;
""".format(since=SINCE)

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

# Data-date tokens, tried in priority order against filename segments.
DATE_D_RE   = re.compile(r"^D(\d{2})(\d{2})(\d{2})$")          # D260508      -> 2026-05-08
DATE_DTRR_RE = re.compile(r"DTRRD_(\d{2})(\d{2})(\d{2})")      # DTRRD_260721 -> 2026-07-21 (YYMMDD)
DATE_TRR_RE = re.compile(r"TRR(\d{2})(\d{2})(\d{4})")          # TRR07232026  -> 2026-07-23 (MMDDYYYY)
DATE_YMD_RE = re.compile(r"(20\d{2})(\d{2})(\d{2})")           # 20260722     -> 2026-07-22 (YYYYMMDD)
DATE_MDY_RE = re.compile(r"^(\d{2})(\d{2})(20\d{2})$")         # 01282026     -> 2026-01-28 (MMDDYYYY)
DATE_YM_RE  = re.compile(r"^(20\d{2})(\d{2})$")                # 202501       -> 2025-01-01 (YYYYMM)

# Leaf-name wrappers that are pure encryption layers over an otherwise
# identical file; strip them so encrypted+decrypted twins dedupe to one file.
ENC_EXT_RE = re.compile(r"\.(pgp|gpg)$", re.IGNORECASE)

# Client-folder staging suffixes to strip (case-insensitive), e.g.
# BCBSNC_Decrypt -> BCBSNC, bcbsks_TempDecrypt -> bcbsks, MMOH_Temp -> MMOH.
CLIENT_SUFFIX_RE = re.compile(r"(?:_?Temp)?(?:_?Decrypt)?$", re.IGNORECASE)

# Fold folder tokens (lowercased key) to a clean display name.  Merges case
# variants and known same-client folder aliases; keeps genuinely distinct
# medical/Rx feed lines separate.
CLIENT_ALIASES = {
    "coventry": "Coventry",
    "anthem": "Anthem",
    "aetna": "Aetna",
    "aetnarx": "AetnaRx",
    "aetnaqnxt": "AetnaQNXT",
    "aetnaqnxtrx": "AetnaQNXTRx",
    "cigna": "Cigna",
    "healthnet": "HealthNet",
    "healthnow": "HealthNow",
    "wellpointrx": "WellpointRx",
    "wellcarerx": "WellcareRx",
    "centene": "Centene",
    "centenerx": "CenteneRx",
    "centenefidelis": "CenteneFidelis",
    "centenefidelisrx": "CenteneFidelisRx",
    "hip_facets": "HIP_Facets",
    "emblemrx": "EmblemRx",
    "evernorth": "Evernorth",
    "mmoh": "MMOH",
    "mmohrx": "MMOHRx",
    "bcbsnc": "BCBSNC",
    "bcbsnc_rx": "BCBSNC_RX",
    "bcbsks": "BCBSKS",
    "excellus": "Excellus",
    "excellusrx": "ExcellusRx",
    "hap": "HAP",
    "bcbsarkansas": "BCBSARRx",   # DMZ folder for the same client as BCBSARRx
    "bcbsarrx": "BCBSARRx",
    "mcs": "MCS",
    "carefirst": "CareFirst",
    "elixir": "ElixirSolutions",
    "elixirsolutions": "ElixirSolutions",
    "elevancemmmrx": "ElevanceMMMRx",
    "premera": "Premera",
    "premera_medadvrx": "Premera_MedAdvRx",
}


def leaf_name(path: str) -> str:
    return re.split(r"[\\/]", path or "")[-1]


def dedup_key(fname: str) -> str:
    """Filename with an encryption wrapper (.pgp/.gpg) stripped, upper-cased."""
    return ENC_EXT_RE.sub("", (fname or "")).upper()


def file_type(fname: str) -> str:
    u = (fname or "").upper()
    if "MARX" in u:
        return "COBC"
    if "ABII" in u:
        return "ABII"
    if "TRR" in u:
        return "TRR"
    return "Other"


def parse_client(source_path: str) -> str:
    """Client from the SourcePath: the segment after \\Clients\\ (skipping a
    users\\ level), staging suffix stripped, known aliases folded."""
    parts = [p for p in re.split(r"[\\/]", source_path or "") if p]
    lower = [p.lower() for p in parts]
    if "clients" not in lower:
        return "Unknown"
    i = lower.index("clients")
    seg = parts[i + 1] if i + 1 < len(parts) else ""
    if seg.lower() == "users" and i + 2 < len(parts):
        seg = parts[i + 2]
    if not seg:
        return "Unknown"
    # \Clients\Machinify\Inbound\UHC_COBC -> UHC (the deepest token, minus _COBC)
    if seg.lower() == "machinify":
        tail = parts[-1]
        tail = re.sub(r"_?COBC$", "", tail, flags=re.IGNORECASE)
        return (tail or "Machinify").upper()
    seg = CLIENT_SUFFIX_RE.sub("", seg) or seg
    return CLIENT_ALIASES.get(seg.lower(), seg)


def parse_contract(fname: str) -> str:
    """Contract token in the leaf file name, leading routing 'R' dropped."""
    for seg in re.split(r"[._ ]", (fname or "").upper()):
        if CONTRACT_SEG_RE.match(seg):
            return LEAD_R_RE.sub("", seg)
    return ""


def parse_extracted(fname: str) -> str:
    """Best-effort data (extracted) date from the leaf name; '' if none."""
    up = (fname or "").upper()
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


def fetch_rows():
    cmd = [
        "sqlcmd", "-S", SERVER, "-d", DATABASE, "-E",
        "-h", "-1", "-W", "-s", SEP, "-Q", SQL,
    ]
    out = subprocess.run(cmd, capture_output=True, text=True, check=True).stdout
    seen = {}          # dedup_key(name) + received-date -> row (keep first / Completed)
    for line in out.splitlines():
        line = line.rstrip("\r\n")
        if not line.strip():
            continue
        parts = line.split(SEP)
        if len(parts) != 5:
            continue
        fname, spath, size, received, status = [p.strip() for p in parts]
        if not fname or fname.lower() == "filename":
            continue
        ftype = file_type(fname)
        if ftype == "Other":
            continue
        rec = received if received and received != "NULL" else ""
        key = dedup_key(fname) + "|" + rec[:10]
        if key in seen:
            continue                       # keep the first (Completed-first from SQL order)
        seen[key] = {
            "client": parse_client(spath),
            "ftype": ftype,
            "contract": parse_contract(fname),
            "filename": leaf_name(fname),
            "extracted": parse_extracted(fname),
            "received": rec,
            "status": status,
            "size": int(size) if size.isdigit() else 0,
        }
    return list(seen.values())


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
  .pager { display: flex; align-items: center; gap: 12px; margin-top: 12px; justify-content: flex-end; font-size: 13px; color: var(--muted); }
  .empty { padding: 24px; text-align: center; color: var(--muted); }

  /* File-type pills */
  .ft { display: inline-block; padding: 1px 8px; border-radius: 10px; font-size: 11px; font-weight: 700;
        letter-spacing: .4px; }
  .ft-COBC { background: #e3f0fb; color: #1f5b8a; }
  .ft-TRR  { background: #eae6fb; color: #5b3a9e; }
  .ft-ABII { background: #e4f6ea; color: #1c7a44; }

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
    display: inline-flex; align-items: center; justify-content: center;
    min-width: 22px; height: 22px; border-radius: 6px; padding: 0 8px;
    background: #e3f4ea; color: #1c7a44; font-weight: 700; font-size: 12px;
    box-shadow: inset 0 0 0 1px rgba(28,122,68,.28);
    transition: transform .08s ease;
    font-variant-numeric: tabular-nums;
  }
  table.matrix td.day.hit:hover .mk { background: #cdecd8; transform: scale(1.12); }
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
</style>
</head>
<body>
<header>
  <h1>COBC / TRR / ABII File Report</h1>
  <div class="meta">Generated __GENERATED__ &middot; TRGUtil10 / Ramp.ramp.FileLog &middot; inbound client files since __SINCE__ &middot; __ROW_COUNT__ files (deduped)</div>
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
        <span style="font-size:12px;color:var(--muted)">Cell = files received that month &middot; click a client to expand its contracts &middot; hover a <b>count</b> for file details.</span>
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
    <div class="kpi"><div class="label">Total Size</div><div class="value" id="kpi-size">&mdash;</div></div>
    <div class="kpi"><div class="label">Latest Received</div><div class="value" id="kpi-latest">&mdash;</div></div>
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
      <label for="f-ldfrom">Received From</label>
      <input type="date" id="f-ldfrom">
    </div>
    <div class="field">
      <label for="f-ldto">Received To</label>
      <input type="date" id="f-ldto">
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
        <th data-key="received">Received<span class="arrow">&#8597;</span></th>
        <th data-key="status">Status<span class="arrow">&#8597;</span></th>
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
    ldfrom: '', ldto: '',
    sortKey: 'received', sortDir: 'desc', page: 0,
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
      if (fnQ && r.filename.toLowerCase().indexOf(fnQ) === -1) return false;
      const ld = r.received ? r.received.slice(0, 10) : '';
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
    let size = 0, latest = '';
    const cset = new Set(), ctset = new Set();
    for (const r of filtered) {
      size += r.size || 0;
      if (r.client) cset.add(r.client);
      if (r.contract) ctset.add(r.contract);
      if (r.received && r.received > latest) latest = r.received;
    }
    $('kpi-files').textContent = filtered.length.toLocaleString();
    $('kpi-clients').textContent = cset.size.toLocaleString();
    $('kpi-contracts').textContent = ctset.size.toLocaleString();
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
        tr.innerHTML =
          '<td>' + esc(r.client) + '</td>' +
          '<td>' + ftPill(r.ftype) + '</td>' +
          '<td>' + esc(r.contract) + '</td>' +
          '<td>' + esc(r.filename) + '</td>' +
          '<td>' + esc(r.extracted) + '</td>' +
          '<td>' + esc(r.received) + '</td>' +
          '<td>' + esc(r.status) + '</td>' +
          '<td class="num' + (r.size ? '' : ' zero') + '">' + fmtSize(r.size) + '</td>';
        body.appendChild(tr);
      }
    }

    for (const th of document.querySelectorAll('th[data-key]')) {
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
    $('btn-reset').addEventListener('click', () => {
      ids.forEach(id => $(id).value = '');
      Object.assign(state, { ftype: '', client: '', contract: '', filename: '', ldfrom: '', ldto: '', page: 0 });
      render();
    });
    document.querySelectorAll('#grid th[data-key]').forEach(th => {
      th.addEventListener('click', () => {
        const k = th.dataset.key;
        if (state.sortKey === k) {
          state.sortDir = state.sortDir === 'asc' ? 'desc' : 'asc';
        } else {
          state.sortKey = k;
          state.sortDir = (k === 'size' || k === 'received' || k === 'extracted') ? 'desc' : 'asc';
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

  // Distinct received-years (YYYY), newest first.
  const years = [...new Set(ROWS.map(r => r.received).filter(Boolean).map(s => s.slice(0, 4)))].sort().reverse();
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
        '<div class="tt-name">' + escT(f.filename) + '</div>' +
        '<div><span class="tt-label">Type:</span> ' + escT(f.ftype) + '</div>' +
        '<div><span class="tt-label">Date Extracted:</span> ' + (escT(f.extracted) || '&mdash;') + '</div>' +
        '<div><span class="tt-label">Date Received:</span> ' + escT((f.received || '').slice(0, 10)) + '</div>' +
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
  // client/contract is listed each month whether or not it received a file.
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
        cells += '<td class="day hit" data-tip="' + id + '"><span class="mk">' + files.length.toLocaleString('en-US') + '</span></td>';
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
      if (!r.received || r.received.slice(0, 4) !== yr) continue;
      if (monthState.ftype && r.ftype !== monthState.ftype) continue;
      const c = clients[r.client] || (clients[r.client] = { contracts: {}, agg: {} });
      const ct = c.contracts[r.contract] || (c.contracts[r.contract] = {});
      const mo = Number(r.received.slice(5, 7));
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
    return (HTML_TEMPLATE
            .replace("__GENERATED__", datetime.now().strftime("%Y-%m-%d %H:%M"))
            .replace("__SINCE__", SINCE[:10])
            .replace("__ROW_COUNT__", f"{len(rows):,}")
            .replace("__DATA_JSON__", json.dumps(rows, separators=(",", ":"))))


def main():
    print(f"[info] Fetching COBC/TRR/ABII files from {SERVER}/{DATABASE} (DateCreated > {SINCE})")
    rows = fetch_rows()
    by_type = {}
    for r in rows:
        by_type[r["ftype"]] = by_type.get(r["ftype"], 0) + 1
    print(f"[info] {len(rows)} files (deduped) — " + ", ".join(f"{k}={v}" for k, v in sorted(by_type.items())))
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
