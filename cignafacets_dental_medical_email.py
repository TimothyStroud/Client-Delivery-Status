"""
One-off: email the Dental vs Medical PayAmount / record-count breakdown from
CignaFacets.dbo.vwMiningCache_Full (TRGINTP3) for TapeID >= 3618.

Dental is identified by a CDT procedure code in CPT_1 (D####) -- the mining
cache carries no claim-type flag (ClaimIndicator is unmapped/NULL in the
CignaFacets mapping document).
"""
import subprocess
import send_via_outlook

SERVER = 'TRGINTP3'
MIN_TAPE = 3618
EMAIL_TO = 'timothy.stroud@machinify.com'

DENTAL = "CPT_1 LIKE 'D[0-9][0-9][0-9][0-9]'"


def q(sql):
    r = subprocess.run(
        ['sqlcmd', '-S', SERVER, '-E', '-h', '-1', '-W', '-s', '|',
         '-Q', 'SET NOCOUNT ON; ' + sql, '-t', '1800'],
        capture_output=True, text=True)
    rows = []
    for line in r.stdout.splitlines():
        line = line.strip()
        if not line or line.startswith('('):
            continue
        rows.append(line.split('|'))
    return rows


summary = q(f"""
SELECT CASE WHEN {DENTAL} THEN 'Dental' ELSE 'Medical' END,
       COUNT_BIG(*),
       SUM(CAST(ISNULL(PayAmount,0) AS MONEY))
FROM CignaFacets.dbo.vwMiningCache_Full
WHERE TapeID >= {MIN_TAPE}
GROUP BY CASE WHEN {DENTAL} THEN 'Dental' ELSE 'Medical' END
ORDER BY 1;
""")

by_tape = q(f"""
SELECT t.TapeID, CONVERT(varchar(10), c.DataDescription),
       SUM(CASE WHEN {DENTAL} THEN 1 ELSE 0 END),
       SUM(CASE WHEN {DENTAL} THEN CAST(ISNULL(t.PayAmount,0) AS MONEY) ELSE 0 END),
       SUM(CASE WHEN {DENTAL} THEN 0 ELSE 1 END),
       SUM(CASE WHEN {DENTAL} THEN 0 ELSE CAST(ISNULL(t.PayAmount,0) AS MONEY) END)
FROM CignaFacets.dbo.vwMiningCache_Full t
LEFT JOIN CignaFacets.chimera.Tape c ON c.TapeID = t.TapeID
WHERE t.TapeID >= {MIN_TAPE}
GROUP BY t.TapeID, CONVERT(varchar(10), c.DataDescription)
ORDER BY t.TapeID;
""")

tot_rec = sum(int(r[1]) for r in summary)
tot_pay = sum(float(r[2]) for r in summary)


def money(v):
    return '${:,.2f}'.format(float(v))


def num(v):
    return '{:,}'.format(int(v))


TH = ('padding:6px 12px;border:1px solid #cbd5e1;background:#1e3a5f;'
      'color:#fff;text-align:left;font-weight:600;')
TD = 'padding:6px 12px;border:1px solid #cbd5e1;'
TDR = TD + 'text-align:right;'

srows = ''
for cat, rec, pay in summary:
    bg = '#fff7ed' if cat == 'Dental' else '#f8fafc'
    srows += (
        f'<tr style="background:{bg}">'
        f'<td style="{TD}font-weight:600">{cat}</td>'
        f'<td style="{TDR}">{num(rec)}</td>'
        f'<td style="{TDR}">{100.0*int(rec)/tot_rec:.2f}%</td>'
        f'<td style="{TDR}">{money(pay)}</td>'
        f'<td style="{TDR}">{100.0*float(pay)/tot_pay:.2f}%</td></tr>')
srows += (
    f'<tr style="background:#e2e8f0;font-weight:700">'
    f'<td style="{TD}">Total</td><td style="{TDR}">{num(tot_rec)}</td>'
    f'<td style="{TDR}">100.00%</td><td style="{TDR}">{money(tot_pay)}</td>'
    f'<td style="{TDR}">100.00%</td></tr>')

trows = ''
for tid, ddate, drec, dpay, mrec, mpay in by_tape:
    flag = ''
    if int(drec) < 5000:
        flag = ' <span style="color:#b91c1c;font-weight:700">&#9888;</span>'
    trows += (
        f'<tr><td style="{TD}">{tid}{flag}</td>'
        f'<td style="{TD}">{ddate}</td>'
        f'<td style="{TDR}">{num(drec)}</td>'
        f'<td style="{TDR}">{money(dpay)}</td>'
        f'<td style="{TDR}">{num(mrec)}</td>'
        f'<td style="{TDR}">{money(mpay)}</td></tr>')

html = f"""<html><body style="font-family:Segoe UI,Arial,sans-serif;font-size:13px;color:#0f172a">
<h2 style="color:#1e3a5f;margin-bottom:2px">CignaFacets &mdash; Dental vs. Medical Breakdown</h2>
<div style="color:#475569;margin-bottom:16px">
Source: <code>CignaFacets.dbo.vwMiningCache_Full</code> on <b>TRGINTP3</b> &nbsp;|&nbsp;
Filter: <code>TapeID &gt;= {MIN_TAPE}</code> (4 claims tapes, data dates 07/01/2026 &ndash; 07/22/2026)
</div>

<h3 style="color:#1e3a5f;margin-bottom:4px">Summary</h3>
<table style="border-collapse:collapse;font-size:13px">
<tr><th style="{TH}">Category</th><th style="{TH}">Record Count</th><th style="{TH}">% Recs</th>
<th style="{TH}">Pay Amount</th><th style="{TH}">% Pay</th></tr>
{srows}
</table>

<h3 style="color:#1e3a5f;margin-bottom:4px;margin-top:22px">By Tape</h3>
<table style="border-collapse:collapse;font-size:13px">
<tr><th style="{TH}">TapeID</th><th style="{TH}">Data Date</th>
<th style="{TH}">Dental Recs</th><th style="{TH}">Dental Pay</th>
<th style="{TH}">Medical Recs</th><th style="{TH}">Medical Pay</th></tr>
{trows}
</table>

<div style="margin-top:22px;padding:12px 14px;background:#fef2f2;border-left:4px solid #b91c1c">
<b>&#9888; Data quality note &mdash; dental volume dropped sharply on the last two tapes.</b>
<ul style="margin:8px 0 0 0;padding-left:20px">
<li><b>Tape 3626</b> (data date 07/15/2026): dental records fell to 23,390 from ~61,000 on the two
prior tapes. This tape also has <b>2,110,622 of 2,872,740 rows (73%) with no CPT/procedure code</b>,
versus ~5% on tapes 3618 and 3622 &mdash; so the dental undercount here is likely a symptom of
missing procedure codes rather than missing dental claims.</li>
<li><b>Tape 3631</b> (data date 07/22/2026): only <b>742</b> dental records, and dental pay is
<b>negative (&minus;$2,935.89)</b>, i.e. adjustments/reversals only. CPT population on this tape is
normal (5.8% null), but there are <b>zero</b> rows on a dental plan (<code>PlanID LIKE 'D%'</code>)
and <b>zero</b> rows with a tooth number &mdash; all three dental markers collapse together, which
points to dental claims being absent from the 07/22 file rather than a coding issue.</li>
</ul>
<div style="margin-top:8px">The dental totals above are therefore understated for the period.
Recommend confirming with Cigna whether the 07/15 and 07/22 claim extracts were complete.</div>
</div>

<h3 style="color:#1e3a5f;margin-bottom:4px;margin-top:22px">How dental is identified</h3>
<div style="color:#334155;max-width:820px">
The CignaFacets mapping document defines no claim-type flag &mdash; <code>ClaimIndicator</code> is
mapped to <code>NULL</code>, and <code>ServiceType</code> is null on ~99% of dental lines. Dental is
therefore identified by a <b>CDT procedure code in <code>CPT_1</code></b>
(<code>CPT_1 LIKE 'D[0-9][0-9][0-9][0-9]'</code>); everything else is classified as Medical.
This is corroborated by <code>ToothNumber</code> (populated only on CDT lines) and
<code>PlanID</code> (DPPO/DENT/DPPL prefixes). Note the CDT test intentionally also captures
dental services billed under a medical plan.
</div>

<div style="margin-top:20px;color:#64748b;font-size:11px">
Query run against TRGINTP3 &middot; CignaFacets.dbo.vwMiningCache_Full
(<code>Cache.MiningCacheHistory</code> UNION ALL <code>vwMiningCache_Full_1</code>).
</div>
</body></html>"""

print(send_via_outlook.send(
    EMAIL_TO,
    'CignaFacets - Dental vs Medical Pay Amount & Record Counts (TapeID >= 3618)',
    html))
