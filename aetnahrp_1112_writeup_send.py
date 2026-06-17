"""Send the AetnaHRP 11/12/2025 paid-claims investigation write-up to Tim."""
import sys
sys.path.insert(0, r'C:\Users\tls2\.claude\projects\H--')
from send_via_outlook import send

# November 2025 daily paid (all tapes), 11/12 flagged
nov = [
    ("2025-11-03","Mon","1,561,137","$119.51M",False),
    ("2025-11-04","Tue","1,765,637","$198.99M",False),
    ("2025-11-05","Wed","2,561,393","$220.94M",False),
    ("2025-11-06","Thu","2,288,821","$212.92M",False),
    ("2025-11-07","Fri","2,425,505","$244.40M",False),
    ("2025-11-08","Sat","2,469,552","$212.83M",False),
    ("2025-11-10","Mon","1,733,169","$132.47M",False),
    ("2025-11-11","Tue","2,167,975","$189.97M",False),
    ("2025-11-12","Wed","380,299","$26.70M",True),
    ("2025-11-13","Thu","2,147,090","$209.42M",False),
    ("2025-11-14","Fri","2,595,247","$223.21M",False),
    ("2025-11-15","Sat","2,049,919","$189.68M",False),
    ("2025-11-17","Mon","1,743,522","$130.35M",False),
    ("2025-11-18","Tue","2,109,564","$180.35M",False),
    ("2025-11-19","Wed","2,758,353","$230.62M",False),
    ("2025-11-20","Thu","2,252,971","$209.71M",False),
    ("2025-11-21","Fri","2,436,234","$208.68M",False),
]
rows = ""
for d, dow, lines, paid, flag in nov:
    style = ' style="background:#ffe0e0;font-weight:bold;color:#c00000"' if flag else ''
    rows += (f'<tr{style}><td style="padding:4px 10px;border:1px solid #ddd">{d}</td>'
             f'<td style="padding:4px 10px;border:1px solid #ddd">{dow}</td>'
             f'<td style="padding:4px 10px;border:1px solid #ddd;text-align:right">{lines}</td>'
             f'<td style="padding:4px 10px;border:1px solid #ddd;text-align:right">{paid}</td></tr>')

body = f"""<html><body style="font-family:Calibri,Arial,sans-serif;font-size:14px;color:#222;line-height:1.45">

<p style="font-size:16px"><b>AetnaHRP &ndash; Paid Claims for Payment Date 11/12/2025 (TRGINTP3)</b></p>

<p><b>Bottom line:</b> For payment date <b>11/12/2025</b>, AetnaHRP holds only
<b>380,299 paid claim lines totaling $26,695,383.30</b> across all tapes. That is
<b>not a cache.Mining / view defect</b> &mdash; the view is summing correctly. The
data for that single payment date is <b>under-loaded by roughly 87%</b> versus the
~$200M daily norm. The ~$28M you saw in cache.mining matches this real (incomplete)
total. The missing volume needs to be re-loaded at the source.</p>

<h3 style="color:#2c5f8a;margin-bottom:4px">1. Total volume for 11/12/2025</h3>
<table style="border-collapse:collapse;font-size:13px;margin-bottom:6px">
<tr style="background:#f0f0f0"><th style="padding:5px 10px;border:1px solid #ddd;text-align:left">Scope (PAYMENT_DATE = 2025-11-12)</th><th style="padding:5px 10px;border:1px solid #ddd">Claim lines</th><th style="padding:5px 10px;border:1px solid #ddd">Paid (SUM PAID_AMOUNT)</th><th style="padding:5px 10px;border:1px solid #ddd">Tapes</th></tr>
<tr><td style="padding:5px 10px;border:1px solid #ddd">All tapes</td><td style="padding:5px 10px;border:1px solid #ddd;text-align:right">380,299</td><td style="padding:5px 10px;border:1px solid #ddd;text-align:right"><b>$26,695,383.30</b></td><td style="padding:5px 10px;border:1px solid #ddd;text-align:right">1,364</td></tr>
<tr><td style="padding:5px 10px;border:1px solid #ddd">&nbsp;&nbsp;TapeID &ge; 12245 (your filter)</td><td style="padding:5px 10px;border:1px solid #ddd;text-align:right">39,035</td><td style="padding:5px 10px;border:1px solid #ddd;text-align:right">$7,425,249.01</td><td style="padding:5px 10px;border:1px solid #ddd;text-align:right">170 (12245&ndash;13107)</td></tr>
<tr><td style="padding:5px 10px;border:1px solid #ddd">&nbsp;&nbsp;TapeID &lt; 12245</td><td style="padding:5px 10px;border:1px solid #ddd;text-align:right">341,264</td><td style="padding:5px 10px;border:1px solid #ddd;text-align:right">$19,270,134.29</td><td style="padding:5px 10px;border:1px solid #ddd;text-align:right">1,194 (4934&ndash;12241)</td></tr>
</table>
<p style="font-size:12px;color:#555">Note: your <code>TapeID &ge; 12245</code> filter only captures $7.4M of it; the rest of that day's payments sit in older tapes (4934&ndash;12241). Even combined, the all-tape total is just ~$26.7M.</p>

<h3 style="color:#2c5f8a;margin-bottom:4px">2. Why ~$28M instead of ~$200M</h3>
<p>11/12/2025 is a clear one-day outlier. Surrounding November business days run
$180M&ndash;$244M; 11/12 alone collapses to $26.7M on ~380K lines (vs ~2&ndash;2.6M
lines/day). ~$173M and ~2.0M lines (~87%) of the expected volume are simply not
present in the database for that payment date.</p>
<table style="border-collapse:collapse;font-size:12px;margin-bottom:6px">
<tr style="background:#f0f0f0"><th style="padding:4px 10px;border:1px solid #ddd;text-align:left">Pay Date</th><th style="padding:4px 10px;border:1px solid #ddd">Day</th><th style="padding:4px 10px;border:1px solid #ddd">Lines</th><th style="padding:4px 10px;border:1px solid #ddd">Paid</th></tr>
{rows}
</table>

<h3 style="color:#2c5f8a;margin-bottom:4px">3. The cache.Mining view is functioning correctly</h3>
<p>Per the <b>Aetna HRP Medical Mapping Document</b> (ETL 2), the view maps
<code>PayAmount = [tcl].[PAID_AMOUNT]</code> (Claims) and
<code>PayDate = [p].[PAYMENT_DATE]</code> (Payment). I verified the claim layer is
fully intact for TapeID&ge;12245: <code>claim.TRGClaim</code> = <b>$33.78B</b> over
397,955,069 lines, and the view's inner join
(<code>cache.TRGCache &harr; claim.TRGClaim</code> on LineKeyID+TapeID) returns the
<b>exact same $33.78B</b> &mdash; i.e., the join is not dropping value.
<code>history.claim</code> holds $34.99B / 410,138,227 lines. So the money is present
where claims exist; the 11/12/2025 shortfall is an <b>upstream load gap for that pay
date</b>, not a view or join problem.</p>

<h3 style="color:#2c5f8a;margin-bottom:4px">How I found this (TRGINTP3, AetnaHRP)</h3>
<ol style="margin-top:4px">
<li>Mapped the lineage: <code>cache.Mining.PayAmount = claim.TRGClaim.PAID_AMOUNT</code>,
<code>PayDate = client.Payment.PAYMENT_DATE</code> (via <code>OBJECT_DEFINITION</code>
of the view + the Aetna HRP Medical Mapping Document, ETL 2 / "MC View Mapping" tab).</li>
<li>Joined <code>client.Payment</code> &harr; claim on <b>TapeID + LineKeyID</b> (the correct
per-line grain &mdash; joining on TapeID alone cross-joins every line to every payment in
the tape) and filtered <code>PAYMENT_DATE = '2025-11-12'</code>: 39,035 lines / $7.43M
for TapeID&ge;12245, 380,299 lines / $26.70M across all tapes.</li>
<li>Confirmed the claim layer total ($33.78B) equals the view's inner-join total
($33.78B) &mdash; no value lost in the view.</li>
<li>Pulled November 2025 day-by-day to show 11/12 is an isolated ~87% shortfall while
neighboring days are normal.</li>
</ol>
<p style="font-size:12px;color:#555">Tooling note: <code>cache.Mining</code> embeds
<code>WITH (FORCESEEK)</code>, so it only runs under full ANSI SET options
(QUOTED_IDENTIFIER/ARITHABORT ON, etc.); I queried the base tables directly to avoid that.</p>

<h3 style="color:#2c5f8a;margin-bottom:4px">Recommendation</h3>
<p>Treat 11/12/2025 as an <b>incomplete payment load</b> for AetnaHRP and re-load /
reprocess that pay cycle's file(s). Once the ~2.0M missing lines land, cache.mining and
any PayDate-based daily totals will reflect the full ~$200M for that date automatically.</p>

<p style="color:#999;font-size:11px">Investigation run against TRGINTP3 on 2026-06-17.
Source: AetnaHRP.history.claim, claim.TRGClaim, cache.TRGCache, client.Payment, cache.Mining.</p>
</body></html>"""

print(send(to='timothy.stroud@machinify.com',
           subject='AetnaHRP - Paid Claims 11/12/2025 Investigation (~$27M actual vs ~$200M expected)',
           body=body))
