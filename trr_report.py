"""
TRR missing contracts report (2025 + 2026) — queries TRGINTP3, emails Timothy.Stroud@machinify.com
"""
import subprocess, re, sys
from collections import defaultdict
from datetime import datetime, date, timedelta

sys.path.insert(0, r'C:\Users\tls2\.claude\projects\H--')
import send_via_outlook

TO      = 'Timothy.Stroud@machinify.com; RDPOperations@Machinify.com'
FROM    = 'DataOperations@machinify.com'
SUBJECT = 'WellcareRx TRR Missing Contracts Report (2025 & 2026)'

EXPECTED = sorted([
    'H0029','H0074','H0111','H0174','H0270','H0712','H0913','H1032','H1112','H1416',
    'H1914','H2117','H2162','H2491','H2775','H2816','H3975','H4073','H4506','H4847',
    'H4868','H5087','H5475','H5779','H5965','H7323','H7326','H7518','H8711','H9364',
    'H9730','H9916','S4802'
])
EXPECTED_SET = set(EXPECTED)

QUERIES = {
    '2025': "SET NOCOUNT ON; SELECT [FileName] FROM [WellcareRx].[etl].[tape] T (nolock) JOIN [WellcareRx].[config].[Table] F (nolock) ON t.TableID = f.TableID WHERE f.[TableName] in ('TRR') and [FileName] like '%D25%'",
    '2026': "SET NOCOUNT ON; SELECT [FileName] FROM [WellcareRx].[etl].[tape] T (nolock) JOIN [WellcareRx].[config].[Table] F (nolock) ON t.TableID = f.TableID WHERE f.[TableName] in ('TRR') and [FileName] like '%D26%'",
}


def run_query(sql):
    result = subprocess.run(
        ['sqlcmd', '-S', 'TRGINTP3', '-E', '-Q', sql, '-W', '-s', '|', '-h', '-1'],
        capture_output=True, text=True
    )
    return result.stdout.splitlines()


def parse_filename(line):
    fname = line.strip().rsplit('\\', 1)[-1]
    m = re.match(r'([A-Z0-9]+)\.[^.]+\.D(\d{6})\.', fname)
    if m:
        return m.group(1), m.group(2)
    return None, None


def yymmdd_to_date(d):
    try:
        return datetime.strptime('20' + d, '%Y%m%d').date()
    except Exception:
        return None


def fmt_date(d):
    return d.strftime('%Y-%m-%d') if isinstance(d, date) else ''


def mon_sat_dates(year, present_dates):
    """All Mon-Sat dates in the range of data, excluding Sundays."""
    if not present_dates:
        return []
    parsed = [yymmdd_to_date(d) for d in present_dates if yymmdd_to_date(d)]
    if not parsed:
        return []
    start = min(parsed)
    end   = min(max(parsed), date.today())
    result = []
    cur = start
    while cur <= end:
        if cur.weekday() != 6:  # 6 = Sunday
            result.append(cur)
        cur += timedelta(days=1)
    return result


def build_year_section(year, present, unexpected):
    th         = 'background:#2c5f8a;color:#fff;padding:7px 11px;text-align:left;white-space:nowrap;'
    td         = 'padding:6px 11px;border-bottom:1px solid #e0e0e0;vertical-align:top;'
    td_a       = 'padding:6px 11px;border-bottom:1px solid #e0e0e0;vertical-align:top;background:#f5f8fc;'
    miss_style = 'background:#fdecea;color:#c0392b;padding:2px 5px;border-radius:3px;font-size:12px;margin:1px 2px;display:inline-block;'
    h2_style   = ('font-family:Segoe UI,Arial,sans-serif;font-size:16px;font-weight:700;color:#fff;'
                  'background:#2c5f8a;padding:8px 14px;margin:28px 0 0 0;border-radius:4px 4px 0 0;')

    all_dates = sorted(present.keys())
    dates_with_missing = [(d, EXPECTED_SET - present[d]) for d in all_dates if EXPECTED_SET - present[d]]
    dates_complete     = [d for d in all_dates if not (EXPECTED_SET - present[d])]

    # --- Missing contracts table ---
    rows = []
    for i, (d, missing) in enumerate(dates_with_missing):
        s      = td_a if i % 2 else td
        badges = ''.join(f'<span style="{miss_style}">{c}</span>' for c in sorted(missing))
        rows.append(f'<tr><td style="{s}">{fmt_date(yymmdd_to_date(d))}</td>'
                    f'<td style="{s}">{len(present[d])}</td>'
                    f'<td style="{s}">{len(missing)}</td>'
                    f'<td style="{s}">{badges}</td></tr>')

    thead         = ''.join(f'<th style="{th}">{h}</th>' for h in ['Date','Received','Missing Count','Missing Contracts'])
    missing_table = (
        f'<table style="border-collapse:collapse;width:100%;min-width:600px;margin-bottom:16px;">'
        f'<thead><tr>{thead}</tr></thead><tbody>{"".join(rows)}</tbody></table>'
    ) if rows else '<p style="color:#369e57;font-weight:600;margin-bottom:16px;">All dates have all 33 contracts.</p>'

    # --- Complete dates ---
    if dates_complete:
        ok_badge = ('background:#e8f5e9;color:#2e7d32;padding:2px 6px;border-radius:3px;'
                    'font-size:12px;margin:2px 3px;display:inline-block;')
        complete_badges = ''.join(
            f'<span style="{ok_badge}">{fmt_date(yymmdd_to_date(d))}</span>' for d in dates_complete
        )
        complete_section = (f'<p style="font-weight:700;color:#2e7d32;margin:16px 0 4px 0;">'
                            f'Dates with all 33 contracts ({len(dates_complete)}):</p><p>{complete_badges}</p>')
    else:
        complete_section = ''

    # --- No-file dates (Mon-Sat gaps) ---
    all_mon_sat     = mon_sat_dates(year, all_dates)
    present_as_date = {yymmdd_to_date(d) for d in all_dates}
    no_file_dates   = [d for d in all_mon_sat if d not in present_as_date]

    if no_file_dates:
        gap_badge = ('background:#f3e5f5;color:#6a1b9a;padding:2px 6px;border-radius:3px;'
                     'font-size:12px;margin:2px 3px;display:inline-block;')
        gap_badges = ''.join(f'<span style="{gap_badge}">{fmt_date(d)} ({d.strftime("%a")})</span>'
                             for d in no_file_dates)
        no_file_section = (f'<p style="font-weight:700;color:#6a1b9a;margin:16px 0 4px 0;">'
                           f'Mon–Sat dates with no files received ({len(no_file_dates)}):</p>'
                           f'<p>{gap_badges}</p>')
    else:
        no_file_section = '<p style="margin-top:16px;color:#555;">No Mon–Sat gaps found.</p>'

    # --- Unexpected contracts ---
    if unexpected:
        unexp_badge = ('background:#fff3cd;color:#856404;padding:2px 6px;border-radius:3px;'
                       'font-size:12px;margin:2px 3px;display:inline-block;')
        unexp_badges = ''.join(f'<span style="{unexp_badge}">{c}</span>' for c in sorted(unexpected))
        unexpected_section = (f'<p style="font-weight:700;color:#856404;margin:16px 0 4px 0;">'
                              f'Contracts received not in expected list ({len(unexpected)}):</p>'
                              f'<p>{unexp_badges}</p>')
    else:
        unexpected_section = '<p style="margin-top:16px;color:#555;">No unexpected contracts found.</p>'

    summary = (f'{len(all_dates)} date(s) found &nbsp;|&nbsp; '
               f'{len(dates_with_missing)} with missing contracts &nbsp;|&nbsp; '
               f'{len(dates_complete)} complete &nbsp;|&nbsp; '
               f'{len(no_file_dates)} Mon–Sat gap(s)')

    return f"""
<p style="{h2_style}">{year} TRR Data</p>
<div style="border:1px solid #2c5f8a;border-top:none;padding:14px 16px 8px 16px;margin-bottom:8px;border-radius:0 0 4px 4px;">
  <p style="color:#555;margin:0 0 14px 0;">{summary}</p>
  <p style="font-weight:700;color:#c0392b;margin:0 0 6px 0;">Dates with missing contracts ({len(dates_with_missing)}):</p>
  {missing_table}
  {complete_section}
  {no_file_section}
  {unexpected_section}
</div>
"""


def main():
    sections = []
    for year, sql in QUERIES.items():
        print(f'Running {year} query...')
        lines   = run_query(sql)
        present = defaultdict(set)
        for line in lines:
            contract, date_str = parse_filename(line)
            if contract and date_str:
                present[date_str].add(contract)
        all_seen   = set().union(*present.values()) if present else set()
        unexpected = all_seen - EXPECTED_SET
        print(f'  {year}: {len(present)} dates, {sum(len(v) for v in present.values())} files, '
              f'{len(unexpected)} unexpected contracts')
        sections.append(build_year_section(year, present, unexpected))

    html = f"""
<html><body style="font-family:Segoe UI,Arial,sans-serif;font-size:13px;color:#222;max-width:1100px;">
<p style="margin-bottom:4px;">
  <strong>WellcareRx TRR — Missing Contracts &amp; Gap Report</strong>
  &nbsp;&nbsp;<span style="color:#555;">as of {datetime.now().strftime('%Y-%m-%d %H:%M')}</span>
  &nbsp;&nbsp;|&nbsp;&nbsp;Expecting {len(EXPECTED)} contracts per date
</p>
{''.join(sections)}
</body></html>
"""
    print(f'Sending to {TO}...')
    result = send_via_outlook.send(TO, SUBJECT, html, from_address=FROM)
    print(f'Result: {result}')


if __name__ == '__main__':
    main()
