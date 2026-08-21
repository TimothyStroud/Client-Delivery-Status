"""CareFirstRx — 2026 COBC & TRR daily file check (CenteneRx-style).

One HTML email with two sections (COBC, then TRR). Each section is the CenteneRx
layout: roster callouts (current / expired / new), a per-contract missing-dates
pivot, and a Mon-Sat daily delivery calendar.

Source: TRGINTP3.CareFirstRx.etl.tape JOIN config.Table.

  TableID 5400 TRR  -> P.R<contract>.DTRRD.D<YYMMDD>.T...
  TableID 5000 DTL  -> P.R<contract>.MARXCOB.D<YYMMDD>.T...   (= COBC)

COBC gotcha (carried over from carefirstrx_abii_cobc_trr_check.py): one physical
COBC file is split into DTL/PRM/SUP layouts and EACH gets its own TapeID, so
querying TableName IN ('DTL','PRM','SUP') over-counts 3x. Count DTL (5000) only.

Only production `P.` files count; `T.` (test) and legacy `EFTO.` are ignored.

Expected-contract logic is fully data-driven (no curated roster):
  * a contract is expected from its first production delivery forward;
  * it stops being expected only if it looks OFFBOARDED, i.e. its own last
    delivery is more than OFFBOARD_DAYS behind the feed's newest delivery.
    If the WHOLE feed goes quiet, nobody is offboarded and every contract flags
    red — which is what we want (that is an outage, not an offboarding).

Per standing policy, the most recent MISSING_DELAY_DAYS (7) days are shown but
never assessed as missing.
"""
import os
import re
import subprocess
import calendar as _cal
from collections import defaultdict, Counter
from datetime import date, timedelta

EMAIL_TO = "timothy.stroud@machinify.com"     # review copy only
EMAIL_FROM = "DataOperations@machinify.com"
YEAR = 2026
OFFBOARD_DAYS = 10          # contract quiet this much longer than the feed => offboarded
MISSING_DELAY_DAYS = 7      # rolling reporting delay (standing COBC/TRR email policy)
RECENT_N = 6                # ~one Mon-Sat week of delivery days for the "current" roster
ESTABLISH_N = 3             # first N delivery days define the initial roster
DRY_RUN = os.environ.get("REPORT_DRYRUN") == "1"

today = date.today()
start = date(YEAR, 1, 1)
_missing_cutoff = today - timedelta(days=MISSING_DELAY_DAYS)


# ---------- Federal holidays ----------
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
        observed(date(year, 1, 1)):  "New Year's Day",
        nth_weekday(year, 1, 0, 3):  "MLK Day",
        nth_weekday(year, 2, 0, 3):  "Presidents' Day",
        last_weekday(year, 5, 0):    "Memorial Day",
        observed(date(year, 6, 19)): "Juneteenth",
        observed(date(year, 7, 4)):  "Independence Day",
        nth_weekday(year, 9, 0, 1):  "Labor Day",
        nth_weekday(year, 10, 0, 2): "Columbus Day",
        observed(date(year, 11, 11)): "Veterans Day",
        nth_weekday(year, 11, 3, 4): "Thanksgiving",
        observed(date(year, 12, 25)): "Christmas Day",
    }


HOLIDAYS = us_federal_holidays(YEAR)
HOLIDAY_ABBR = {
    "New Year's Day": "NYD", "MLK Day": "MLK", "Presidents' Day": "PD",
    "Memorial Day": "MEM", "Juneteenth": "JTH", "Independence Day": "IND",
    "Labor Day": "LAB", "Columbus Day": "COL", "Veterans Day": "VET",
    "Thanksgiving": "TGV", "Christmas Day": "XMS",
}

# ---------- Pull data ----------
SQL = """SET NOCOUNT ON;
SELECT [TapeID], [ProdCtrlNo], [FileName], t.[TableID], f.[TableName], f.[TableType]
FROM [CareFirstRx].[etl].[tape] T (nolock)
JOIN [CareFirstRx].[config].[Table] F (nolock)
ON t.TableID = f.TableID
WHERE f.[TableName] in ('TRR','DTL','PRM','SUP')
"""
SEP = "\x1f"
r = subprocess.run(
    ["sqlcmd", "-S", "TRGINTP3", "-E", "-d", "CareFirstRx",
     "-Q", SQL, "-W", "-s", SEP, "-h", "-1"],
    capture_output=True, text=True, check=False,
)
if r.returncode != 0:
    raise SystemExit(f"sqlcmd failed: {r.stderr[:500]}")

# P.R<contract>.<DTRRD|MARXCOB[A|S]>.D<YYMMDD>.T...
RE = re.compile(r"\\P\.R([SH]\d{4})\.(DTRRD|MARXCOB[AS]|MARXCOB)\.D(\d{6})", re.I)
KIND = {"MARXCOB": "COBC", "MARXCOBA": "COBC", "MARXCOBS": "COBC", "DTRRD": "TRR"}

per_day = {"COBC": defaultdict(Counter), "TRR": defaultdict(Counter)}
seen_tape = {"COBC": set(), "TRR": set()}
# MARXCOBA / MARXCOBS = one-off name variants (a mid-March bulk drop), NOT part of the
# daily cadence — kept out of the grid/roster so they can't skew first-seen or day counts.
backfill = defaultdict(Counter)          # contract -> date -> count
backfill_kinds = Counter()               # MARXCOBA / MARXCOBS -> count
unparsed = Counter()
rows_seen = 0

for line in r.stdout.splitlines():
    if not line.strip() or "rows affected" in line:
        continue
    parts = line.split(SEP)
    if len(parts) < 5:
        continue
    rows_seen += 1
    tape_id, fname, table_id = parts[0].strip(), parts[2].strip(), parts[3].strip()
    # COBC: count the one physical file once -> DTL (5000) only. TRR = 5400.
    if table_id not in ("5000", "5400"):
        continue
    m = RE.search(fname)
    if not m:
        if "\\P.R" in fname:
            unparsed["unmatched"] += 1
        continue                      # T./EFTO. legacy files
    contract, kind_raw, ymd = m.group(1).upper(), m.group(2).upper(), m.group(3)
    kind = KIND[kind_raw]
    try:
        d = date(2000 + int(ymd[:2]), int(ymd[2:4]), int(ymd[4:6]))
    except ValueError:
        unparsed[kind] += 1
        continue
    if d.year != YEAR:
        continue
    if kind_raw in ("MARXCOBA", "MARXCOBS"):
        backfill[contract][d] += 1
        backfill_kinds[kind_raw] += 1
        continue
    if tape_id in seen_tape[kind]:
        continue
    seen_tape[kind].add(tape_id)
    per_day[kind][d][contract] += 1

# ---------- Calendar days (Mon-Sat; no Sunday deliveries for either feed) ----------
all_days = []
d = start
while d <= today:
    if d.weekday() != 6:
        all_days.append(d)
    d += timedelta(days=1)

NCOLS = 5


def flush_month_total(month, year, files_m, days_m):
    return ("<tr class='monthtotal'>"
            f"<td class='month-total'>{_cal.month_name[month]} {year} total</td>"
            f"<td class='month-total'>({days_m} days)</td>"
            f"<td class='num month-total'>{files_m}</td>"
            "<td class='month-total'></td><td class='month-total'></td></tr>")


def misses_by_month_html(dates):
    bym = defaultdict(list)
    for x in dates:
        bym[x.month].append(x)
    return "<br>".join(
        f"<b>{_cal.month_abbr[m]}:</b> " + ", ".join(f"{x.month}/{x.day}" for x in sorted(bym[m]))
        for m in sorted(bym)
    )


def build_section(kind, label, source_note, extra=""):
    pd = per_day[kind]
    first_seen, last_seen = {}, {}
    for dd in sorted(pd):
        for c in pd[dd]:
            first_seen.setdefault(c, dd)
            last_seen[c] = dd

    days_with_data = sorted(pd)
    feed_last = days_with_data[-1] if days_with_data else None

    # Offboarded = quiet more than OFFBOARD_DAYS behind the feed's newest delivery.
    offboarded = {c: ld for c, ld in last_seen.items()
                  if feed_last and (feed_last - ld).days > OFFBOARD_DAYS}

    def expected_for(dd):
        return [c for c in sorted(first_seen)
                if first_seen[c] <= dd
                and not (c in offboarded and dd > offboarded[c])]

    # roster lists
    recent_days = days_with_data[-RECENT_N:]
    recent_start = recent_days[0] if recent_days else today
    recent_cnt = Counter()
    for dd in recent_days:
        for c in pd[dd]:
            recent_cnt[c] += 1
    current = sorted(c for c in recent_cnt if recent_cnt[c] >= 2 or len(recent_days) <= 1)
    exp_by_last = defaultdict(list)
    for c, ld in last_seen.items():
        if c not in set(current) and ld < recent_start:
            exp_by_last[ld].append(c)
    exp_events = [(ld, sorted(exp_by_last[ld])) for ld in sorted(exp_by_last)]
    establish_boundary = (days_with_data[ESTABLISH_N - 1] if len(days_with_data) >= ESTABLISH_N
                          else (days_with_data[-1] if days_with_data else today))
    new = sorted(c for c in first_seen if first_seen[c] > establish_boundary)

    # daily rows
    rows, day_missing = [], {}
    prev_month, month_files, month_days = None, 0, 0
    for dd in all_days:
        if prev_month is not None and dd.month != prev_month:
            rows.append(flush_month_total(prev_month, dd.year, month_files, month_days))
            rows.append(f"<tr class='monthsep'><td colspan='{NCOLS}'></td></tr>")
            month_files = month_days = 0
        prev_month = dd.month

        day_data = pd.get(dd, {})
        total = sum(day_data.values())
        month_files += total
        month_days += 1

        expected = expected_for(dd)
        nexp = len(expected)
        received = {c for c, n in day_data.items() if n > 0}
        base_recv = sum(1 for c in expected if c in received)
        holiday_name = HOLIDAYS.get(dd)
        is_pending = dd > _missing_cutoff
        is_holiday = holiday_name is not None
        excluded = is_pending or is_holiday or nexp == 0
        missing = [] if excluded else [c for c in expected if c not in received]
        day_missing[dd] = missing

        if holiday_name:
            row_cls = " class='holiday'"
        elif dd.weekday() == 5:
            row_cls = " class='satday'"
        else:
            row_cls = ""
        date_label = f"{dd:%Y-%m-%d}"
        if holiday_name:
            date_label += f" ({HOLIDAY_ABBR.get(holiday_name, holiday_name)})"
        recv_cls = "num" if excluded else ("num " + ("ok" if base_recv == nexp else "missing"))
        cells = [f"<td{row_cls}>{date_label}</td>",
                 f"<td{row_cls}>{dd:%a}</td>",
                 f"<td class='num'>{total}</td>",
                 f"<td class='{recv_cls}'>{base_recv} / {nexp}</td>"]
        if nexp == 0:
            cells.append("<td class='pending'>no contracts onboarded yet</td>")
        elif is_holiday:
            cells.append("<td class='holidaycell'>Federal holiday &mdash; excluded from missing</td>")
        elif is_pending:
            cells.append("<td class='pending'>within 1-week delay &mdash; not yet assessed</td>")
        elif not received:
            cells.append("<td class='missing'>ALL &mdash; no file</td>")
        elif missing:
            cells.append(f"<td class='missing'>{', '.join(missing)}</td>")
        else:
            cells.append("<td class='ok'>&mdash;</td>")
        rows.append("<tr>" + "".join(cells) + "</tr>")
    if prev_month is not None:
        rows.append(flush_month_total(prev_month, all_days[-1].year, month_files, month_days))

    # per-contract pivot of day_missing (so the two views always agree)
    contract_misses = defaultdict(list)
    for dd in all_days:
        for c in day_missing.get(dd, []):
            contract_misses[c].append(dd)
    cm_rows = [f"<tr><td>{c}</td><td class='num'>{len(contract_misses[c])}</td>"
               f"<td style='font-size:11px'>{misses_by_month_html(contract_misses[c])}</td></tr>"
               for c in sorted(contract_misses)]

    total_files = sum(sum(v.values()) for v in pd.values())
    n_days_data = sum(1 for dd in all_days if pd.get(dd))
    days_short = sum(1 for dd in all_days if day_missing.get(dd))

    expired_render = ("; ".join(f"<b>last {ld:%m/%d}:</b> " + ", ".join(grp)
                               for ld, grp in exp_events) if exp_events else "none")
    new_render = (", ".join(f"{c} <span style='color:#777'>(first {first_seen[c]:%m/%d/%y})</span>"
                            for c in new) if new
                  else "None &mdash; no contract first appeared after the initial roster.")
    offb_render = ("; ".join(f"{c} <span style='color:#777'>(last {ld:%m/%d/%y})</span>"
                             for c, ld in sorted(offboarded.items())) if offboarded
                   else "none &mdash; every contract is still within "
                        f"{OFFBOARD_DAYS} days of the feed's newest delivery")

    outage = ""
    if feed_last and (today - feed_last).days > MISSING_DELAY_DAYS:
        outage = (f"<p class='alert'>&#9888; <b>Feed-wide gap:</b> the newest {label} production file "
                  f"is dated <b>{feed_last:%Y-%m-%d}</b> &mdash; {(today - feed_last).days} days ago. "
                  f"The whole feed went quiet at once, so the <b>{len(expected_for(today))} contract(s) "
                  f"still expected ({', '.join(expected_for(today))}) are flagged missing every day</b> "
                  f"from {feed_last + timedelta(days=1):%Y-%m-%d} to the 1-week cutoff. "
                  f"This is an outage/onboarding question, not an offboarding.</p>")

    return f"""<h2 style='margin-top:26px;border-bottom:2px solid #305f9c'>{label}</h2>
<p>{source_note}<br>
Coverage: {start:%Y-%m-%d} &rarr; {today:%Y-%m-%d} (Sundays excluded &mdash; neither feed delivers Sunday;
Saturdays shaded cream; federal holidays purple).</p>
{outage}{extra}
<div class='callout' style='background:#eef7ee;border-color:#1a7a1a'>
<b style='color:#1a7a1a'>Current contracts ({len(current)})</b> &mdash; delivering in the last
{len(recent_days)} delivery days ({recent_start:%Y-%m-%d} &rarr; {feed_last if feed_last else today:%Y-%m-%d}).<br>
<span style='font-size:11px'>{', '.join(current) if current else 'none'}</span>
</div>
<div class='callout' style='background:#fff6f6;border-color:#a40000'>
<b style='color:#a40000'>Expired contracts ({len(exp_events and [c for _, g in exp_events for c in g] or [])})</b>
&mdash; delivered earlier in {YEAR} but absent from the recent window.<br>
<span style='font-size:11px'>{expired_render}</span>
</div>
<div class='callout'>
<b>Treated as offboarded ({len(offboarded)})</b> &mdash; no longer expected after their last delivery
(quiet &gt; {OFFBOARD_DAYS} days behind the feed).<br>
<span style='font-size:11px'>{offb_render}</span>
</div>
<div class='callout' style='background:#eef7ee;border-color:#1a7a1a'>
<b style='color:#1a7a1a'>New contracts ({len(new)})</b> &mdash; first delivering after the initial roster
established ({establish_boundary:%Y-%m-%d}).<br>
<span style='font-size:11px'>{new_render}</span>
</div>

<p><b>{total_files}</b> {label} files across <b>{len(all_days)}</b> Mon&ndash;Sat days;
<b>{n_days_data}</b> days have at least one file; <b>{days_short}</b> days were short of expected.<br>
<span style='font-size:11px;color:#556'><b>One-week reporting delay:</b> the most recent
{MISSING_DELAY_DAYS} days ({_missing_cutoff + timedelta(days=1):%Y-%m-%d} &rarr; {today:%Y-%m-%d})
are shown but <b>not assessed for missing files</b>.</span></p>

<h3>{label} &mdash; missing deliveries by contract</h3>
<p style='max-width:1100px;font-size:11px'>Pivot of the calendar below (they match exactly): days a contract
was expected but did not deliver. Feed-wide no-file days count as a miss for every expected contract.
Excluded: federal holidays and the last {MISSING_DELAY_DAYS} days.</p>
<table><thead><tr><th>Contract</th><th># Misses</th><th>Missing dates by month</th></tr></thead>
<tbody>{''.join(cm_rows) if cm_rows else '<tr><td colspan=3>none</td></tr>'}</tbody></table>

<h3>{label} &mdash; daily delivery calendar</h3>
<table><thead><tr><th>Date</th><th>Day</th><th>Total Files</th>
<th>Received (of expected)</th><th>Missing Contracts</th></tr></thead>
<tbody>{''.join(rows)}</tbody></table>""", day_missing, contract_misses


_bf_total = sum(sum(v.values()) for v in backfill.values())
_bf_dates = sorted({d for v in backfill.values() for d in v})
_bf_note = ("" if not _bf_total else
            "<div class='callout'><b>Excluded from the grid: "
            f"{_bf_total} non-standard COBC files</b> ("
            + ", ".join(f"{k}: {n}" for k, n in sorted(backfill_kinds.items()))
            + ") landing on "
            + ", ".join(f"{d:%Y-%m-%d}" for d in _bf_dates)
            + " &mdash; "
            + "; ".join(f"{c}: {sum(backfill[c].values())}" for c in sorted(backfill))
            + ". These use the <code>MARXCOBA</code>/<code>MARXCOBS</code> name variants instead of "
              "<code>MARXCOB</code> and appear on no other dates, so they are kept out of the daily "
              "cadence, the day counts and the "
              "first-seen roster &mdash; otherwise they would make H8854/S0375 look onboarded in "
              "mid-March and flag ~5 weeks of false misses.</div>")

cobc_html, cobc_dm, cobc_cm = build_section(
    "COBC", "COBC (MARXCOB)",
    "Source: <code>TRGINTP3.CareFirstRx.etl.tape</code> JOIN <code>config.Table</code>, "
    "<b>TableID 5000 (DTL) only</b> &mdash; one physical COBC file is split into DTL/PRM/SUP "
    "layouts, each with its own TapeID, so counting all three over-counts 3&times;. "
    "Production <code>P.R&lt;contract&gt;.MARXCOB.D&lt;YYMMDD&gt;</code> files only "
    "(<code>T.</code> test and legacy <code>EFTO.</code> excluded).", extra=_bf_note)
trr_html, trr_dm, trr_cm = build_section(
    "TRR", "TRR (DTRRD)",
    "Source: <code>TRGINTP3.CareFirstRx.etl.tape</code> JOIN <code>config.Table</code>, "
    "<b>TableID 5400 (TRR)</b>. Production "
    "<code>P.R&lt;contract&gt;.DTRRD.D&lt;YYMMDD&gt;</code> files only.")

html = f"""<html><head><style>
body {{ font-family: 'Segoe UI', sans-serif; font-size: 12px; }}
table {{ border-collapse: collapse; }}
th, td {{ border: 1px solid #999; padding: 3px 8px; }}
th {{ background: #305f9c; color: white; white-space: nowrap; }}
td.num {{ text-align: center; font-variant-numeric: tabular-nums; }}
.ok {{ background: #d9f4d9; }}
.missing {{ background: #fde4e4; color: #a40000; font-weight: bold; }}
.pending {{ background: #eef1f6; color: #667; font-style: italic; }}
.holidaycell {{ background: #f0e4ff; color: #5b2c8c; font-style: italic; }}
.satday {{ background: #fff8e1; font-style: italic; }}
.holiday {{ background: #f0e4ff !important; font-weight: bold; }}
tr.monthsep td {{ background: #305f9c; height: 4px; padding: 0; border: none; }}
tr.monthtotal td.month-total {{ background: #d6e0f0 !important; font-weight: bold;
        border-top: 2px solid #305f9c; border-bottom: 2px solid #305f9c; }}
.callout {{ background: #fff3da; border: 1px solid #c47f00; padding: 8px 12px; margin: 10px 0; }}
p.alert {{ background: #fde4e4; border: 1px solid #a40000; color: #a40000; padding: 8px 12px; }}
h2 {{ color: #305f9c; }}
</style></head><body>
<h1 style='font-size:16px'>CareFirstRx &mdash; COBC &amp; TRR daily file check, {YEAR}
<span style='font-weight:normal;font-size:12px;color:#666'>(through {today:%Y-%m-%d})</span></h1>
<p style='font-size:11px;color:#666'>Contracts and dates are parsed from the <b>FileName</b>
(ProdCtrlNo is an internal numeric ID, not the contract). Expected rosters are fully data-driven:
a contract is expected from its first production delivery forward, and stops being expected only if it
falls &gt; {OFFBOARD_DAYS} days behind the feed's newest delivery (i.e. it individually offboarded).
Modeled on the CenteneRx TRR / COBC SUP checks.</p>
{cobc_html}
{trr_html}
</body></html>
"""

# ---------- Self-verify: calendar vs by-contract pivot ----------
for _k, _dm, _cm in (("COBC", cobc_dm, cobc_cm), ("TRR", trr_dm, trr_cm)):
    a = {(c, d) for d in all_days for c in _dm.get(d, [])}
    b = {(c, d) for c, ds in _cm.items() for d in ds}
    print(f"[verify] {_k}: {'MATCH' if a == b else 'MISMATCH diff=' + str(len(a ^ b))} "
          f"({len(a)} missing pairs)")
print(f"[info] backfill(MARXCOBA)={_bf_total}; unparsed={dict(unparsed)}")
print(f"[info] {rows_seen} tape rows; "
      f"COBC {sum(sum(v.values()) for v in per_day['COBC'].values())} files / "
      f"{len(per_day['COBC'])} days; "
      f"TRR {sum(sum(v.values()) for v in per_day['TRR'].values())} files / "
      f"{len(per_day['TRR'])} days")

out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "carefirstrx_cobc_trr_check.html")
with open(out, "w", encoding="utf-8") as fh:
    fh.write(html)
print(f"[html] {out}")

if not DRY_RUN:
    import sys
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    import send_via_outlook
    send_via_outlook.send(EMAIL_TO,
                          f"CareFirstRx — COBC & TRR daily check {YEAR} (through {today:%Y-%m-%d})",
                          html, from_address=EMAIL_FROM)
    print(f"[done] Sent to {EMAIL_TO}")
else:
    print("[dry-run] send skipped")
