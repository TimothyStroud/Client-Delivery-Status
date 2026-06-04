"""Sync MissingPayDates frmMain/rptClientReport data to Notion.

Source: trgintp3.DataOperations.dbo.vw_PayAmount (the SQL Server view the
.accdb pass-through hits). Logic mirrors qryReport from
M:\\DIG\\MissingPayDatesNpt3.accdb: LEFT JOIN a calendar to per-(client, day)
payment sums, emit "Weekend" for Sat/Sun with no payment, "Not Received"
for Mon-Fri with no payment, otherwise blank with the dollar amount.

Window: 2026-01-01 through today (inclusive).

Auth: reads Notion integration token from env var NOTION_INTEGRATION_TOKEN
(or from secrets file at C:\\Users\\tls2\\.claude\\secrets\\notion_token.txt
as a fallback).

Run:
    python C:\\Users\\tls2\\.claude\\projects\\H--\\missing_pay_dates_sync.py
"""

import os
import sys
import json
import time
import calendar
import subprocess
from datetime import date, timedelta, datetime
from urllib import request as urlrequest, error as urlerror

# --- Configuration -----------------------------------------------------------

NOTION_DB_URL = "https://www.notion.so/a817676241994d4bb08b09dd5b7249bc"
DATA_SOURCE_ID = "6e5fde80-abef-443d-a263-f3b7aa6b6e21"
WINDOW_START = date(2025, 10, 1)  # Q4 2025 forward per user 2026-06-02
SQL_SERVER = "trgintp3"
SQL_DATABASE = "DataOperations"
SECRETS_FALLBACK = r"C:\Users\tls2\.claude\secrets\notion_token.txt"


def load_token() -> str:
    tok = os.environ.get("NOTION_INTEGRATION_TOKEN", "").strip()
    if tok:
        return tok
    if os.path.exists(SECRETS_FALLBACK):
        with open(SECRETS_FALLBACK, "r", encoding="utf-8") as f:
            return f.read().strip()
    raise SystemExit(
        "No Notion token found. Set NOTION_INTEGRATION_TOKEN env var or "
        f"create {SECRETS_FALLBACK}"
    )


# --- SQL Server data fetch ---------------------------------------------------

def fetch_payments(window_start: date, window_end: date):
    """Return list of (client, date, amount) tuples from SQL Server."""
    sql = (
        f"SET NOCOUNT ON; "
        f"SELECT DatabaseName, CONVERT(varchar(10), PayDate, 23) AS d, "
        f"SUM(TotalPayAmount) AS amt "
        f"FROM dbo.vw_PayAmount "
        f"WHERE PayDate >= '{window_start:%Y-%m-%d}' "
        f"  AND PayDate <= '{window_end:%Y-%m-%d}' "
        f"GROUP BY DatabaseName, PayDate"
    )
    cp = subprocess.run(
        ["sqlcmd", "-S", SQL_SERVER, "-d", SQL_DATABASE, "-E",
         "-h", "-1", "-W", "-s", "|", "-Q", sql],
        capture_output=True, text=True, timeout=300,
    )
    if cp.returncode != 0:
        raise RuntimeError(f"sqlcmd failed: {cp.stderr or cp.stdout}")
    rows = []
    for line in cp.stdout.splitlines():
        line = line.strip()
        if not line or "rows affected" in line or "|" not in line:
            continue
        parts = line.split("|")
        if len(parts) != 3:
            continue
        client, d_str, amt_str = (p.strip() for p in parts)
        try:
            d = datetime.strptime(d_str, "%Y-%m-%d").date()
            amt = float(amt_str)
        except ValueError:
            continue
        rows.append((client, d, amt))
    return rows


def fetch_clients():
    """All client DatabaseName values from the Access tblClient list.
    We replicate the list inline since tblClient is a small static table —
    avoids needing OLE DB at sync time."""
    return [
        "AetnaHMO", "AetnaHRP", "AetnaRCE", "BCBSFL", "BCBSKS",
        "BCBSKSMedAdv", "BCBSNC", "BCBSNorthCarolinaFEP", "BCBSSC",
        "BSCA_Facets", "BSCA_Medicare", "Cambia", "CambiaFEP",
        "CareFirstDC", "CareFirstFacets", "CareFirstNasco", "CareSource",
        "Centene", "Chickering", "Christus", "CignaFacets",
        "CignaHealthSpring", "CignaPower", "CignaProClaims", "ConnectiCare",
        "Coventry", "CTCARE_Medicare", "EDW_ANE", "EDW_ASE", "EDW_ASEGNC",
        "EDW_C_FAC", "EDW_C_NAS", "EDW_EMPIRE", "EDW_WGS", "EmblemFacets",
        "GEHA", "GHI", "HAP_Medical", "HarvardPilgrim", "HealthNetCA",
        "HealthSpring_FWA", "HIP", "HIP_Montefiore", "HMSA",
        "Kaiser_AmbM", "Kaiser_AmbS", "Kaiser_CO", "Kaiser_GA",
        "Kaiser_HealthConnect", "Kaiser_HI", "Kaiser_MASTapestry",
        "Kaiser_NCTapestry", "Kaiser_NW", "Kaiser_SCTapestry", "Kaiser_WA",
        "MedicalMutual_Gen", "MedicalMutualOH", "Orthonet", "Oscar",
        "Oxford_ARO", "PHS_NICE_ACP", "Premera", "PremeraMedAdvVIS",
        "Tufts_Audit_CIT", "Tufts_PublicPlan", "TuftsMedPref", "UnitedASO",
        "UnitedCosmos", "UnitedCSP", "UnitedUNET", "WellMark",
        "KaiserAMBHI", "KaiserAMBCO", "KaiserAMBGA", "KaiserAMBNW",
        "KaiserAMBN", "KaiserNCPareo", "KaiserSCPareo",
    ]


def us_federal_holidays(year):
    """Return {date: name} for the 11 US federal holidays in `year`.
    Fixed-date holidays falling on Saturday are observed the prior Friday;
    on Sunday, the following Monday."""
    out = {}

    def nth_weekday(month, weekday, n):
        cal = calendar.monthcalendar(year, month)
        days = [w[weekday] for w in cal if w[weekday] != 0]
        return date(year, month, days[n - 1])

    def last_weekday(month, weekday):
        cal = calendar.monthcalendar(year, month)
        days = [w[weekday] for w in cal if w[weekday] != 0]
        return date(year, month, days[-1])

    def observed(d):
        if d.weekday() == 5:
            return d - timedelta(days=1)
        if d.weekday() == 6:
            return d + timedelta(days=1)
        return d

    fixed = [
        (date(year,  1,  1), "New Year's Day"),
        (date(year,  6, 19), "Juneteenth"),
        (date(year,  7,  4), "Independence Day"),
        (date(year, 11, 11), "Veterans Day"),
        (date(year, 12, 25), "Christmas Day"),
    ]
    for d, name in fixed:
        d_obs = observed(d)
        out[d_obs] = name + (" (observed)" if d_obs != d else "")

    out[nth_weekday(1,  0, 3)]  = "MLK Day"
    out[nth_weekday(2,  0, 3)]  = "Presidents' Day"
    out[last_weekday(5, 0)]     = "Memorial Day"
    out[nth_weekday(9,  0, 1)]  = "Labor Day"
    out[nth_weekday(10, 0, 2)]  = "Columbus Day"
    out[nth_weekday(11, 3, 4)]  = "Thanksgiving"
    return out


def build_rows(window_start: date, window_end: date):
    """Return list of dicts mirroring qryReport columns for every active
    (client, day) cell in the window."""
    payments = fetch_payments(window_start, window_end)
    paid_by_client = {}
    active_clients = set()
    for client, d, amt in payments:
        paid_by_client[(client, d)] = amt
        active_clients.add(client)

    # Only emit rows for clients that have any payment activity in the window
    # OR are in tblClient (so a never-paying client still surfaces as
    # "Not Received" every weekday — same as the Access report).
    clients = sorted(set(fetch_clients()) | active_clients)

    weekday_names = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday",
                     "Saturday", "Sunday"]
    # Federal-holiday index covering every year in the window.
    holidays = {}
    for yr in range(window_start.year, window_end.year + 1):
        holidays.update(us_federal_holidays(yr))

    rows = []
    d = window_start
    while d <= window_end:
        wd = weekday_names[d.weekday()]
        is_weekend = d.weekday() >= 5
        holiday_name = holidays.get(d)
        for client in clients:
            amt = paid_by_client.get((client, d))
            parts = []
            if amt is None:
                parts.append("Weekend" if is_weekend else "Not Received")
            elif amt < 0:
                parts.append("Negative Paid")
            if holiday_name:
                parts.append(f"Holiday: {holiday_name}")
            comment = "; ".join(parts) if parts else None
            rows.append({
                "client": client,
                "date": d,
                "weekday": wd,
                "amount": amt,
                "comment": comment,
                "year": d.year,
                "month": d.month,
            })
        d += timedelta(days=1)
    return rows


# --- Notion API client (no SDK to keep deps minimal) -------------------------

NOTION_API = "https://api.notion.com/v1"
NOTION_VERSION = "2022-06-28"


def notion_request(token, method, path, body=None):
    url = NOTION_API + path
    data = None
    headers = {
        "Authorization": f"Bearer {token}",
        "Notion-Version": NOTION_VERSION,
        "Content-Type": "application/json",
    }
    if body is not None:
        data = json.dumps(body).encode("utf-8")
    req = urlrequest.Request(url, data=data, method=method, headers=headers)
    for attempt in range(5):
        try:
            with urlrequest.urlopen(req, timeout=60) as resp:
                return json.loads(resp.read().decode("utf-8"))
        except urlerror.HTTPError as e:
            body_text = e.read().decode("utf-8", errors="replace")
            if e.code in (429, 502, 503, 504) and attempt < 4:
                sleep_s = 2 ** attempt
                print(f"[warn] HTTP {e.code} — retrying in {sleep_s}s")
                time.sleep(sleep_s)
                continue
            raise RuntimeError(f"Notion API {e.code}: {body_text}")
    raise RuntimeError("Notion API: retries exhausted")


def fetch_existing(token, data_source_id):
    """Return {(client, iso_date): page_id} for all live pages in the DB."""
    out = {}
    cursor = None
    while True:
        body = {"page_size": 100}
        if cursor:
            body["start_cursor"] = cursor
        resp = notion_request(token, "POST",
                              f"/data_sources/{data_source_id}/query", body)
        for page in resp.get("results", []):
            props = page.get("properties", {})
            name = ""
            tnode = props.get("DatabaseName", {})
            for t in tnode.get("title", []):
                name += t.get("plain_text", "")
            d_node = props.get("Date", {}).get("date") or {}
            d_iso = d_node.get("start", "")
            if name and d_iso:
                out[(name, d_iso[:10])] = page["id"]
        if not resp.get("has_more"):
            break
        cursor = resp.get("next_cursor")
    return out


def row_properties(row):
    props = {
        "DatabaseName": {"title": [{"text": {"content": row["client"]}}]},
        "Date": {"date": {"start": row["date"].isoformat()}},
        "DayOfWeekName": {"select": {"name": row["weekday"]}},
        "Year": {"number": row["year"]},
        "Month": {"number": row["month"]},
    }
    if row["amount"] is not None:
        props["SumOfTotalPayAmount"] = {"number": round(row["amount"], 2)}
    if row["comment"]:
        props["Comment"] = {"select": {"name": row["comment"]}}
    return props


def create_pages(token, data_source_id, rows):
    """Create new pages one at a time (Notion has no bulk-create endpoint).
    Light rate limiting to stay under 3 req/sec average."""
    created = 0
    for row in rows:
        notion_request(token, "POST", "/pages", {
            "parent": {"type": "data_source_id",
                       "data_source_id": data_source_id},
            "properties": row_properties(row),
        })
        created += 1
        if created % 50 == 0:
            print(f"[info]   created {created}/{len(rows)}")
        time.sleep(0.34)
    return created


def update_page(token, page_id, row):
    notion_request(token, "PATCH", f"/pages/{page_id}", {
        "properties": row_properties(row),
    })


def archive_page(token, page_id):
    notion_request(token, "PATCH", f"/pages/{page_id}", {"archived": True})


# --- Main --------------------------------------------------------------------

def sync(initial_seed=False, lookback_days=30):
    token = load_token()
    today = date.today()
    win_end = today
    if initial_seed:
        win_start = WINDOW_START
    else:
        win_start = max(WINDOW_START, today - timedelta(days=lookback_days))
    print(f"[info] Sync window: {win_start} → {win_end}")

    print("[info] Building rows from SQL Server…")
    rows = build_rows(win_start, win_end)
    print(f"[info]   {len(rows)} rows")

    print("[info] Loading existing Notion pages…")
    existing = fetch_existing(token, DATA_SOURCE_ID)
    print(f"[info]   {len(existing)} existing pages")

    want_keys = set()
    to_create = []
    to_update = []
    for row in rows:
        key = (row["client"], row["date"].isoformat())
        want_keys.add(key)
        if key in existing:
            to_update.append((existing[key], row))
        else:
            to_create.append(row)

    # Archive obsolete pages that fall in our sync window but no longer
    # match any row (e.g. tblClient list changed). Only archive pages
    # whose date falls within [win_start, win_end].
    to_archive = []
    for (client, d_iso), pid in existing.items():
        try:
            d = datetime.strptime(d_iso, "%Y-%m-%d").date()
        except ValueError:
            continue
        if win_start <= d <= win_end and (client, d_iso) not in want_keys:
            to_archive.append(pid)

    print(f"[info] Plan: create={len(to_create)}, update={len(to_update)}, archive={len(to_archive)}")

    if to_create:
        print("[info] Creating new pages…")
        create_pages(token, DATA_SOURCE_ID, to_create)

    if to_update:
        print("[info] Updating changed pages…")
        for i, (pid, row) in enumerate(to_update, 1):
            update_page(token, pid, row)
            if i % 50 == 0:
                print(f"[info]   updated {i}/{len(to_update)}")
            time.sleep(0.34)

    if to_archive:
        print("[info] Archiving obsolete pages…")
        for pid in to_archive:
            archive_page(token, pid)
            time.sleep(0.34)

    print("[done] Sync complete.")


def export_batches(window_start: date, window_end: date, out_path: str):
    """Write Notion-create-pages-ready batches as JSON.

    Used by the Claude cron (no integration token available): Claude reads
    this file, dedupes against existing pages via notion-query-data-sources,
    and pushes new rows via notion-create-pages. Output schema:
        {"batches": [[{properties:{}}, ...], [...], ...],
         "window_start": "YYYY-MM-DD", "window_end": "YYYY-MM-DD"}
    """
    rows = build_rows(window_start, window_end)
    pages = []
    for r in rows:
        props = {
            "DatabaseName": r["client"],
            "date:Date:start": r["date"].isoformat(),
            "DayOfWeekName": r["weekday"],
            "Year": r["year"],
            "Month": r["month"],
        }
        if r["amount"] is not None:
            props["SumOfTotalPayAmount"] = round(r["amount"], 2)
        if r["comment"]:
            props["Comment"] = r["comment"]
        pages.append({"properties": props})
    batches = [pages[i:i + 100] for i in range(0, len(pages), 100)]
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump({
            "window_start": window_start.isoformat(),
            "window_end": window_end.isoformat(),
            "row_count": len(pages),
            "batches": batches,
        }, f)
    print(f"[done] wrote {len(pages)} rows in {len(batches)} batches -> {out_path}")


def main():
    args = sys.argv[1:]
    if "--export" in args:
        # --export [--start YYYY-MM-DD] [--end YYYY-MM-DD] [--out path]
        start = None
        end = None
        out = r"C:\Users\tls2\.claude\projects\H--\missing_pay_dates_today.json"
        for i, a in enumerate(args):
            if a == "--start" and i + 1 < len(args):
                start = datetime.strptime(args[i + 1], "%Y-%m-%d").date()
            elif a == "--end" and i + 1 < len(args):
                end = datetime.strptime(args[i + 1], "%Y-%m-%d").date()
            elif a == "--out" and i + 1 < len(args):
                out = args[i + 1]
        today = date.today()
        if end is None:
            end = today
        if start is None:
            start = today
        export_batches(start, end, out)
        return
    seed = "--seed" in args or "--initial" in args
    sync(initial_seed=seed)


if __name__ == "__main__":
    main()
