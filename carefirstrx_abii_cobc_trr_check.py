r"""CareFirstRx — ABII, COBC & TRR file delivery check (modeled on the
WellpointEdwardRx TRR report).

Source: TRGETL3.CareFirstRx.etl.tape joined to config.Table.
  TableID 5400            -> TRR   (daily)   \Data\TRR\P.R<contract>.DTRRD.D<YYMMDD>...
  TableID 5000/5100/5200  -> COBC  (daily)   \Data\COBC\P.R<contract>.MARXCOB.D<YYMMDD>...
                            (DTL/PRM/SUP are three layouts of ONE physical COBC file;
                             dedup by TapeID so a file counts once.)
  TableID 5500/5600       -> ABII  (monthly) <contract>_Dialysis|Transplant_Report_<YYYYMM>...

Contracts are parsed from the FileName. TRR/COBC use a filename-embedded daily
date (YYMMDD); ABII uses a monthly period (YYYYMM). Contracts are auto-discovered
and tracked as expected from their first-seen date forward (never retroactively).

Emails an HTML report to the user via the interactive Outlook sender.
"""
import re
import subprocess
import calendar as _cal
from collections import defaultdict
from datetime import date, datetime, timedelta

BASE = r'C:\Users\tls2\.claude\projects\H--'
import sys
sys.path.insert(0, BASE)
from send_via_outlook import send

EMAIL_TO = "timothy.stroud@machinify.com"
YEAR = 2026


# ---------------- US federal holidays (same as Wellpoint) ----------------
def us_federal_holidays(year):
    def nth_weekday(y, m, wd, n):
        days = [d for d in _cal.Calendar().itermonthdates(y, m)
                if d.month == m and d.weekday() == wd]
        return days[n - 1]

    def last_weekday(y, m, wd):
        days = [d for d in _cal.Calendar().itermonthdates(y, m)
                if d.month == m and d.weekday() == wd]
        return days[-1]

    def observed(d):
        if d.weekday() == 5:
            return d - timedelta(days=1)
        if d.weekday() == 6:
            return d + timedelta(days=1)
        return d

    return {
        observed(date(year, 1, 1)):   "New Year's Day",
        nth_weekday(year, 1, 0, 3):   "MLK Day",
        nth_weekday(year, 2, 0, 3):   "Presidents' Day",
        last_weekday(year, 5, 0):     "Memorial Day",
        observed(date(year, 6, 19)):  "Juneteenth",
        observed(date(year, 7, 4)):   "Independence Day",
        nth_weekday(year, 9, 0, 1):   "Labor Day",
        nth_weekday(year, 10, 0, 2):  "Columbus Day",
        observed(date(year, 11, 11)): "Veterans Day",
        nth_weekday(year, 11, 3, 4):  "Thanksgiving",
        observed(date(year, 12, 25)): "Christmas Day",
    }


HOLIDAYS = us_federal_holidays(YEAR)
HOLIDAY_ABBR = {"New Year's Day": "NYD", "MLK Day": "MLK", "Presidents' Day": "PD",
                "Memorial Day": "MEM", "Juneteenth": "JTH", "Independence Day": "IND",
                "Labor Day": "LAB", "Columbus Day": "COL", "Veterans Day": "VET",
                "Thanksgiving": "TGV", "Christmas Day": "XMS"}

SEP = "\x1f"


def query(sql):
    r = subprocess.run(
        ["sqlcmd", "-S", "TRGETL3", "-E", "-d", "CareFirstRx",
         "-Q", "SET NOCOUNT ON;\n" + sql, "-W", "-s", SEP, "-h", "-1"],
        capture_output=True, text=True, check=False)
    out = []
    for line in r.stdout.splitlines():
        if not line.strip() or "rows affected" in line:
            continue
        out.append(line.split(SEP))
    return out


# ---------------- Parse helpers ----------------
# Production files start with "P."; test files "T." are ignored.
DAILY_RE = re.compile(r"P\.R([A-Z]\d{4})\.(?:DTRRD|MARXCOB)\.D(\d{6})", re.I)
ABII_RE = re.compile(r"([A-Z]\d{4})_(Dialysis|Transplant)_Report_(\d{6})", re.I)


def parse_daily(rows):
    """rows = [[TapeID, FileName], ...] -> per_day[date][contract]=count (deduped by TapeID)."""
    per_day = defaultdict(lambda: defaultdict(int))
    seen_tape = set()
    for parts in rows:
        if len(parts) < 2:
            continue
        tape, fname = parts[0].strip(), parts[1].strip()
        if tape in seen_tape:          # dedup: one physical file counted once
            continue
        m = DAILY_RE.search(fname)
        if not m:
            continue
        contract = m.group(1).upper()
        y6 = m.group(2)
        try:
            d = date(2000 + int(y6[:2]), int(y6[2:4]), int(y6[4:6]))
        except ValueError:
            continue
        if d.year != YEAR:
            continue
        seen_tape.add(tape)
        per_day[d][contract] += 1
    return per_day


# ---------------- Contract lifecycle detection (offboarding / new) ----------------
OFFBOARD_DAYS = 10   # no delivery in this many days (while others deliver) => offboarding
NEW_DAYS = 15        # first-seen within this many days => newly added


def detect_lifecycle(daily_feeds, today):
    """daily_feeds = {'TRR': per_day, 'COBC': per_day}. Return (offboarding, new)
    lists of (contract, detail) using combined first/last-seen across feeds."""
    first_seen, last_seen = {}, {}
    for per_day in daily_feeds.values():
        for d in per_day:
            for c in per_day[d]:
                if c not in first_seen or d < first_seen[c]:
                    first_seen[c] = d
                if c not in last_seen or d > last_seen[c]:
                    last_seen[c] = d
    offboarding, added = [], []
    for c in sorted(first_seen):
        gap = (today - last_seen[c]).days
        span = (last_seen[c] - first_seen[c]).days
        if gap > OFFBOARD_DAYS and span > 0:      # had a history, now silent
            offboarding.append((c, f"last delivery {last_seen[c]:%Y-%m-%d} ({gap} days ago)"))
        if (today - first_seen[c]).days <= NEW_DAYS:
            added.append((c, f"first delivery {first_seen[c]:%Y-%m-%d}"))
    return offboarding, added


# ---------------- Daily calendar table (TRR / COBC) ----------------
def first_seen_map(per_day):
    fs = {}
    for d in sorted(per_day):
        for c in per_day[d]:
            fs.setdefault(c, d)
    return fs


def build_daily_table(per_day, cadence_note):
    """Return (html_table, stats_dict) for a Wellpoint-style daily grid."""
    first_seen = first_seen_map(per_day)
    DISPLAY = sorted(first_seen)                 # all auto-discovered, alpha

    def expected_today(c, d):
        return first_seen.get(c, date.max) <= d

    today = date.today()
    start = date(YEAR, 1, 1)
    all_days = []
    d = start
    while d <= today:
        if d.weekday() != 6:                     # skip Sundays
            all_days.append(d)
        d += timedelta(days=1)

    def fmt_count(n):
        return str(n) if n else "<span style='color:#999'>0</span>"

    def month_total_row(month, year, totals, days_in_month):
        cells = [f"<td class='month-total'>{_cal.month_name[month]} {year} total</td>",
                 f"<td class='month-total'>({days_in_month} days)</td>"]
        for c in DISPLAY:
            cells.append(f"<td class='num month-total'>{totals.get(c, 0)}</td>")
        cells.append("<td class='month-total'></td>")
        return "<tr class='monthtotal'>" + "".join(cells) + "</tr>"

    rows = []
    prev_month = None
    mtot = defaultdict(int)
    mdays = 0
    for d in all_days:
        if prev_month is not None and d.month != prev_month:
            rows.append(month_total_row(prev_month, d.year, mtot, mdays))
            rows.append(f"<tr class='monthsep'><td colspan='{len(DISPLAY) + 3}'></td></tr>")
            mtot = defaultdict(int)
            mdays = 0
        prev_month = d.month
        day_data = per_day.get(d, {})
        for c in DISPLAY:
            mtot[c] += day_data.get(c, 0)
        mdays += 1
        received = {c for c, n in day_data.items() if n > 0}
        missing = [c for c in DISPLAY if expected_today(c, d) and c not in received]
        holiday = HOLIDAYS.get(d)
        if holiday:
            row_cls = " class='holiday'"
        elif d.weekday() == 5:
            row_cls = " class='satday'"
        else:
            row_cls = ""
        date_label = f"{d:%Y-%m-%d}"
        if holiday:
            date_label += f" ({HOLIDAY_ABBR.get(holiday, holiday)})"
        cells = [f"<td{row_cls}>{date_label}</td>", f"<td{row_cls}>{d:%a}</td>"]
        for c in DISPLAY:
            n = day_data.get(c, 0)
            cls = "num ok" if n > 0 else ("num missing" if expected_today(c, d) else "num")
            cells.append(f"<td class='{cls}'>{fmt_count(n)}</td>")
        cells.append(f"<td class='missing'>{', '.join(missing)}</td>" if missing
                     else "<td class='ok'>&mdash;</td>")
        rows.append("<tr>" + "".join(cells) + "</tr>")
    if prev_month is not None:
        rows.append(month_total_row(prev_month, all_days[-1].year, mtot, mdays))

    th = "".join(f"<th{' class=' + chr(39) + 'grp-start' + chr(39) if i == 0 else ''}>{c}</th>"
                 for i, c in enumerate(DISPLAY))
    total_files = sum(sum(v.values()) for v in per_day.values())
    days_missing = sum(1 for d in all_days
                       if any(expected_today(c, d) and c not in per_day.get(d, {}) for c in DISPLAY))
    table = f"""<table><thead><tr>
<th>Date</th><th>Day</th>{th}<th class='grp-start'>Missing Contracts</th>
</tr></thead><tbody>{''.join(rows)}</tbody></table>"""
    stats = dict(contracts=DISPLAY, total_files=total_files, days=len(all_days),
                 days_missing=days_missing, cadence=cadence_note)
    return table, stats


# ---------------- Monthly table (ABII) ----------------
def _ym_int(ym):
    return int(ym[:4]) * 100 + int(ym[4:])


def _ym_add(ym_int, months):
    y, m = divmod(ym_int, 100)
    total = (y * 12 + (m - 1)) + months
    return (total // 12) * 100 + (total % 12) + 1


def _ym_label(ym_int):
    return f"{ym_int // 100}-{ym_int % 100:02d}"


def build_abii_table(rows, dialysis_last):
    """rows = [[TapeID, FileName, TableName], ...] -> monthly grid per ABII type.
    Dialysis: highlighted if no YEAR files. Transplant: every-other-month cadence
    with missing-month flags."""
    seen = set()
    data = {"Dialysis": defaultdict(lambda: defaultdict(int)),
            "Transplant": defaultdict(lambda: defaultdict(int))}
    for parts in rows:
        if len(parts) < 2:
            continue
        tape, fname = parts[0].strip(), parts[1].strip()
        m = ABII_RE.search(fname)
        if not m:
            continue
        contract, kind, ym = m.group(1).upper(), m.group(2).capitalize(), m.group(3)
        if not ym.startswith(str(YEAR)):
            continue
        key = (tape, kind, contract, ym)
        if key in seen:
            continue
        seen.add(key)
        data[kind][ym][contract] += 1

    out = []

    # --- Dialysis: highlight when nothing received this year ---
    dmd = data["Dialysis"]
    if not dmd:
        last_note = f" Most recent file on record: <b>{dialysis_last}</b>." if dialysis_last else ""
        out.append(
            "<h4>ABII Dialysis</h4>"
            f"<p class='alert'>&#9888; ABII Dialysis &mdash; <b>NO files received in {YEAR}.</b>{last_note}</p>")
    else:
        contracts = sorted({c for mm in dmd.values() for c in mm})
        th = "".join(f"<th>{c}</th>" for c in contracts)
        body = []
        zero = "<span style='color:#999'>0</span>"
        for ym in sorted(dmd):
            cells = [f"<td>{ym[:4]}-{ym[4:]}</td>"]
            for c in contracts:
                n = dmd[ym].get(c, 0)
                cells.append(f"<td class='{'num ok' if n else 'num'}'>{n if n else zero}</td>")
            body.append("<tr>" + "".join(cells) + "</tr>")
        out.append(f"<h4>ABII Dialysis</h4><table><thead><tr><th>Month</th>{th}</tr></thead>"
                   f"<tbody>{''.join(body)}</tbody></table>")

    # --- Transplant: every-other-month; flag missing due months ---
    tmd = data["Transplant"]
    if not tmd:
        out.append(f"<h4>ABII Transplant</h4><p class='alert'>&#9888; ABII Transplant &mdash; "
                   f"<b>NO files received in {YEAR}.</b> (expected every other month)</p>")
    else:
        present = sorted(_ym_int(ym) for ym in tmd)
        contracts = sorted({c for mm in tmd.values() for c in mm})
        cur_ym = date.today().year * 100 + date.today().month
        # Expected sequence: from first present month, stepping +2, up to now.
        expected = []
        ym = present[0]
        while ym <= cur_ym:
            expected.append(ym)
            ym = _ym_add(ym, 2)
        # Due = expected months strictly before the current month (current month
        # not yet overdue). Missing = due months with no file.
        missing = [m for m in expected if m < cur_ym and m not in present]
        th = "".join(f"<th>{c}</th>" for c in contracts)
        body = []
        zero = "<span style='color:#999'>0</span>"
        # Render every expected month row (present or missing) so gaps are visible.
        for m in expected:
            is_missing = m in missing
            is_pending = (m == cur_ym and m not in present)
            if is_missing:
                lbl = f"{_ym_label(m)} <b>&#9888; MISSING</b>"
                rcls = " class='missing'"
            elif is_pending:
                lbl = f"{_ym_label(m)} (pending)"
                rcls = " class='satday'"
            else:
                lbl = _ym_label(m)
                rcls = ""
            cells = [f"<td{rcls}>{lbl}</td>"]
            for c in contracts:
                n = tmd.get(f"{m//100}{m%100:02d}", {}).get(c, 0)
                if n:
                    cells.append(f"<td class='num ok'>{n}</td>")
                elif is_missing:
                    cells.append("<td class='num missing'>0</td>")
                else:
                    cells.append(f"<td class='num'>{zero}</td>")
            body.append(f"<tr>" + "".join(cells) + "</tr>")
        miss_note = (f"<p class='alert'>&#9888; ABII Transplant missing expected month(s): "
                     f"<b>{', '.join(_ym_label(m) for m in missing)}</b> "
                     f"(delivered every other month).</p>" if missing else "")
        out.append(f"<h4>ABII Transplant <span style='font-weight:normal;font-size:11px'>"
                   f"(expected every other month)</span></h4>{miss_note}"
                   f"<table><thead><tr><th>Month</th>{th}</tr></thead>"
                   f"<tbody>{''.join(body)}</tbody></table>")
    return "".join(out)


# ---------------- Pull data ----------------
trr_rows = query("""SELECT t.[TapeID], t.[FileName]
FROM [CareFirstRx].[etl].[tape] T (NOLOCK)
WHERE t.[TableID] = 5400 AND t.[FileName] LIKE '%.D26%'""")

cobc_rows = query("""SELECT t.[TapeID], t.[FileName]
FROM [CareFirstRx].[etl].[tape] T (NOLOCK)
WHERE t.[TableID] IN (5000,5100,5200) AND t.[FileName] LIKE '%.D26%'""")

abii_rows = query("""SELECT t.[TapeID], t.[FileName], f.[TableName]
FROM [CareFirstRx].[etl].[tape] T (NOLOCK)
JOIN [CareFirstRx].[config].[Table] F (NOLOCK) ON t.TableID = f.TableID
WHERE t.[TableID] IN (5500,5600) AND t.[FileName] LIKE '%_20%'""")

# Most recent Dialysis file on record (any year) for the no-2026 highlight.
_dl = query("""SELECT TOP 1 t.[FileName] FROM [CareFirstRx].[etl].[tape] T (NOLOCK)
WHERE t.[TableID] = 5500 ORDER BY t.[TapeID] DESC""")
dialysis_last = None
if _dl:
    _m = ABII_RE.search(_dl[0][0])
    if _m:
        dialysis_last = f"{_m.group(3)[:4]}-{_m.group(3)[4:]}"

today = date.today()

trr_per_day = parse_daily(trr_rows)
cobc_per_day = parse_daily(cobc_rows)

trr_table, trr_stats = build_daily_table(trr_per_day, "daily")
cobc_table, cobc_stats = build_daily_table(cobc_per_day, "daily")
abii_html = build_abii_table(abii_rows, dialysis_last)

# Contract lifecycle callout (offboarding / newly added) across TRR + COBC.
offboarding, added = detect_lifecycle({"TRR": trr_per_day, "COBC": cobc_per_day}, today)
callout = ""
if offboarding or added:
    items = []
    for c, detail in offboarding:
        items.append(f"<li><b>{c}</b> appears to be <b>offboarding</b> &mdash; {detail} "
                     f"(no delivery in &gt;{OFFBOARD_DAYS} days while other contracts continue).</li>")
    for c, detail in added:
        items.append(f"<li><b>{c}</b> is a <b>newly added</b> contract &mdash; {detail}.</li>")
    callout = ("<div class='callout'><b>&#9888; Contract changes detected</b>"
               f"<ul>{''.join(items)}</ul></div>")

STYLE = """<style>
body { font-family: 'Segoe UI', sans-serif; font-size: 12px; }
table { border-collapse: collapse; margin-bottom: 18px; }
th, td { border: 1px solid #999; padding: 3px 8px; }
th { background: #305f9c; color: white; }
td.num { text-align: center; font-variant-numeric: tabular-nums; }
.ok { background: #d9f4d9; }
.missing { background: #fde4e4; color: #a40000; font-weight: bold; }
.satday { background: #fff8e1; font-style: italic; }
.holiday { background: #f0e4ff !important; font-weight: bold; }
tr.monthsep td { background: #305f9c; height: 4px; padding: 0; border: none; }
tr.monthtotal td.month-total { background: #d6e0f0 !important; font-weight: bold;
    border-top: 2px solid #305f9c; border-bottom: 2px solid #305f9c; }
.grp-start { border-left: 2px solid #000 !important; }
h3 { color: #305f9c; margin-top: 26px; }
h4 { color: #305f9c; margin: 8px 0 4px; }
.callout { background: #fff3cd; border: 2px solid #e0a800; border-radius: 5px;
           padding: 10px 14px; margin: 12px 0; color: #6b4e00; }
.callout ul { margin: 6px 0 0 18px; }
p.alert { background: #fde4e4; border-left: 4px solid #a40000; color: #a40000;
          font-weight: bold; padding: 6px 10px; margin: 6px 0; }
</style>"""


def section(title, table, stats):
    return f"""<h3>{title}</h3>
<p>Expected contracts (auto-discovered, tracked from first delivery):
<code>{', '.join(stats['contracts']) or '(none seen)'}</code><br>
<b>{stats['total_files']}</b> files across <b>{stats['days']}</b> Mon-Sat days;
<b>{stats['days_missing']}</b> days missing one or more expected contracts.</p>
{table}"""


html = f"""<html><head>{STYLE}</head><body>
<h2>CareFirstRx &mdash; ABII, COBC &amp; TRR delivery check ({YEAR}, through {today:%Y-%m-%d})</h2>
<p>Source: <code>TRGETL3.CareFirstRx.etl.tape</code> + <code>config.Table</code>.
Contract and file date are parsed from the FileName. TRR (TableID 5400) and COBC
(5000/5100/5200, deduped to one row per physical file) are <b>daily</b>; ABII
(5500 Dialysis / 5600 Transplant) is <b>monthly</b>. Only production (<code>P.</code>) files counted.</p>

{callout}

{section('TRR &mdash; daily (P.R&lt;contract&gt;.DTRRD.D&lt;YYMMDD&gt;)', trr_table, trr_stats)}
{section('COBC &mdash; daily (P.R&lt;contract&gt;.MARXCOB.D&lt;YYMMDD&gt;)', cobc_table, cobc_stats)}

<h3>ABII &mdash; monthly (&lt;contract&gt;_Dialysis|Transplant_Report_&lt;YYYYMM&gt;)</h3>
{abii_html}

<p style='font-size:11px;color:#666'>Green = file present. Red bold = expected contract missing that day.
Saturdays cream, holidays purple. First draft &mdash; reply with any layout / expected-contract changes.</p>
</body></html>"""

subject = f"CareFirstRx — ABII, COBC & TRR delivery check {YEAR} (through {today:%Y-%m-%d})"
result = send(to=EMAIL_TO, subject=subject, body=html)
print(f"[send] {result}")
print(f"[TRR]  contracts={trr_stats['contracts']} files={trr_stats['total_files']}")
print(f"[COBC] contracts={cobc_stats['contracts']} files={cobc_stats['total_files']}")
print(f"[ABII] rows_pulled={len(abii_rows)}")
