"""BCBSARRx COBC — H3554 dates NOT received (2025 & 2026), excluding Sundays.

Pulls every H3554 COBC file (TableID 5100, FileName LIKE '%H3554%') from
TRGETL3, parses the delivery date from the filename, and emails the list of
non-Sunday calendar dates from 2025-01-01 through today that had NO H3554 file.
Saturdays and federal holidays are flagged for context (still listed, since
only Sundays were asked to be excluded).
"""
import re
import calendar as _cal
import subprocess
from collections import defaultdict
from datetime import date, timedelta


EMAIL_TO = "timothy.stroud@machinify.com; RDPOperations@Machinify.com"
SERVER = "TRGETL3"
TODAY = date(2026, 6, 4)
START = date(2025, 1, 1)


def us_federal_holidays(year):
    def nth_weekday(year, month, weekday, n):
        days = [d for d in _cal.Calendar().itermonthdates(year, month)
                if d.month == month and d.weekday() == weekday]
        return days[n - 1]

    def last_weekday(year, month, weekday):
        days = [d for d in _cal.Calendar().itermonthdates(year, month)
                if d.month == month and d.weekday() == weekday]
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


HOLIDAYS = {**us_federal_holidays(2025), **us_federal_holidays(2026)}

# ---------- Pull data ----------
SQL = """SET NOCOUNT ON;
SELECT [FileName]
FROM [BCBSARRx].[etl].[tape] T (NOLOCK)
JOIN [BCBSARRx].[config].[Table] F (NOLOCK) ON t.TableID = f.TableID
WHERE t.[TableID] IN (5100) AND [FileName] LIKE '%H3554%';
"""
r = subprocess.run(
    ["sqlcmd", "-S", SERVER, "-E", "-d", "BCBSARRx",
     "-Q", SQL, "-W", "-h", "-1"],
    capture_output=True, text=True, check=False,
)
if r.returncode != 0:
    raise SystemExit(f"sqlcmd failed: {r.stderr}")

RE = re.compile(r"P\.RH3554\.MARXCOBA?\.D(\d{6})", re.I)
received = set()
for line in r.stdout.splitlines():
    line = line.strip()
    if not line or "rows affected" in line:
        continue
    m = RE.search(line)
    if not m:
        continue
    yy, mm, dd = int(m.group(1)[:2]), int(m.group(1)[2:4]), int(m.group(1)[4:6])
    try:
        received.add(date(2000 + yy, mm, dd))
    except ValueError:
        pass

last_received = max(received) if received else None

# ---------- Find missing non-Sunday dates ----------
missing = []
d = START
while d <= TODAY:
    if d.weekday() != 6 and d not in received:   # 6 = Sunday
        missing.append(d)
    d += timedelta(days=1)

by_month = defaultdict(list)
for d in missing:
    by_month[(d.year, d.month)].append(d)

miss_2025 = [d for d in missing if d.year == 2025]
miss_2026 = [d for d in missing if d.year == 2026]
# Split 2026 into within-window vs the continuous gap after the last load.
gap_after = [d for d in miss_2026 if last_received and d > last_received]
miss_2026_inwindow = [d for d in miss_2026 if not (last_received and d > last_received)]

# ---------- Build HTML ----------
def flag(d):
    h = HOLIDAYS.get(d)
    if h:
        return f"<span style='color:#5b2c8c;font-weight:bold'>{h}</span>"
    if d.weekday() == 5:
        return "<span style='color:#8a6d00'>Saturday</span>"
    return ""

month_blocks = []
for (y, m) in sorted(by_month):
    days = by_month[(y, m)]
    rows = "".join(
        f"<tr><td>{d:%Y-%m-%d}</td><td>{d:%a}</td><td>{flag(d)}</td></tr>"
        for d in days
    )
    month_blocks.append(
        f"<tr class='mhdr'><td colspan='3'>{_cal.month_name[m]} {y} "
        f"&mdash; {len(days)} date(s) not received</td></tr>{rows}"
    )

gap_note = ""
if gap_after:
    gap_note = (
        f"<p style='background:#fde4e4;color:#a40000;padding:8px;border:1px solid #a40000;'>"
        f"<b>Note:</b> The last H3554 COBC file loaded was for delivery date "
        f"<b>{last_received:%Y-%m-%d}</b>. Every non-Sunday date after that "
        f"({len(gap_after)} dates, {gap_after[0]:%Y-%m-%d} through {gap_after[-1]:%Y-%m-%d}) "
        f"is one continuous gap — likely H3554 COBC stopped loading rather than "
        f"{len(gap_after)} separate misses. Worth confirming whether this is a real "
        f"delivery stoppage or unloaded data.</p>"
    )

html = f"""<html><head><style>
body {{ font-family:'Segoe UI',sans-serif; font-size:12px; }}
table {{ border-collapse:collapse; }}
th,td {{ border:1px solid #999; padding:3px 10px; }}
th {{ background:#305f9c; color:white; }}
tr.mhdr td {{ background:#d6e0f0; font-weight:bold; border-top:2px solid #305f9c; }}
tr:nth-child(even) td {{ background:#f4f6fa; }}
</style></head><body>

<h3>BCBSARRx COBC &mdash; contract H3554: dates NOT received (2025 &amp; 2026, Sundays excluded)</h3>
<p>Source: <code>{SERVER}.BCBSARRx.etl.tape</code> joined to <code>config.Table</code>
WHERE TableID = 5100 (COBC) AND FileName LIKE '%H3554%'.<br>
Delivery date parsed from the filename (<code>P.RH3554.MARXCOB[A].D&lt;YYMMDD&gt;</code>);
any H3554 COBC file (daily or supplemental) counts as received.<br>
Window: {START:%Y-%m-%d} through {TODAY:%Y-%m-%d}. Sundays excluded.
Last H3554 COBC delivery on file: <b>{last_received:%Y-%m-%d}</b>.</p>

<p><b>{len(missing)}</b> non-Sunday date(s) with no H3554 file &mdash;
<b>{len(miss_2025)}</b> in 2025, <b>{len(miss_2026)}</b> in 2026
(of which {len(gap_after)} fall in the continuous post-{last_received:%Y-%m-%d} gap).</p>

{gap_note}

<table>
<thead><tr><th>Date</th><th>Day</th><th>Flag</th></tr></thead>
<tbody>
{''.join(month_blocks)}
</tbody></table>

<p style='font-size:11px;color:#666'>Saturdays and federal holidays are flagged but still listed
(only Sundays were excluded per request).</p>
</body></html>
"""

import win32com.client
outlook = win32com.client.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)
mail.To = EMAIL_TO
mail.Subject = (f"BCBSARRx COBC H3554 — dates NOT received 2025 & 2026 "
                f"({len(missing)} non-Sunday dates; last load {last_received:%Y-%m-%d})")
mail.HTMLBody = html
mail.Send()
print(f"[done] Sent to {EMAIL_TO}")
print(f"[info] last H3554 COBC delivery: {last_received}")
print(f"[info] missing non-Sunday dates: {len(missing)} "
      f"(2025={len(miss_2025)}, 2026={len(miss_2026)}, post-last-load gap={len(gap_after)})")
print(f"[info] 2025 missing: " + ", ".join(f"{d:%m-%d}" for d in miss_2025))
print(f"[info] 2026 in-window missing (<= {last_received}): "
      + (", ".join(f"{d:%m-%d}" for d in miss_2026_inwindow) or "none"))
