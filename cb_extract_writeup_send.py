"""Email the VENDOR.CB-CLAIMS-EXTRACT 11/12/2025 review to Tim."""
import sys
sys.path.insert(0, r'C:\Users\tls2\.claude\projects\H--')
from send_via_outlook import send

body = """<html><body style="font-family:Calibri,Arial,sans-serif;font-size:14px;color:#222;line-height:1.45">

<p style="font-size:16px"><b>Review &ndash; VENDOR.CB-CLAIMS-EXTRACT.260429220537 &middot; Pay Dates for 11/12/2025</b></p>
<p style="font-size:12px;color:#555">File: <code>\\\\trgdatacap2\\EDrive\\TLS2\\VENDOR.CB-CLAIMS-EXTRACT.260429220537.csv</code> (6.48 GB, 231 columns, 1,984,856 data rows)</p>

<p><b>Answer:</b> Yes &mdash; this extract <u>is</u> the 11/12/2025 processing cycle. <b>Every one of its
1,984,856 claim lines has MOST_RECENT_PROCESS_DATE = 11/12/2025</b>, and <b>1,868,215 lines
(94.1%) are PAYABLE on 11/12/2025, totaling $227,527,616.71</b> &mdash; right in line with the
~$200M expected daily volume.</p>

<p><b>Important distinction:</b> this is the <b>Claim</b> feed, and it has <u>no PAYMENT_DATE column</u>.
In the AetnaHRP model (confirmed in the Aetna HRP Medical Mapping Document, ETL 2), the report's
<b>PayDate maps to client.Payment.PAYMENT_DATE</b>, which is delivered in a <b>separate Payment raw
file</b> (mapping Config.Table: Claim = TableID 1000, <b>Payment = TableID 1100</b>). So the only
&ldquo;pay dates&rdquo; present in this Claims extract are <b>PAYABLE_DATE</b> and
<b>MOST_RECENT_PROCESS_DATE</b> &mdash; not the PAYMENT_DATE that drives the report.</p>

<h3 style="color:#2c5f8a;margin-bottom:4px">What the file contains for 11/12/2025</h3>
<table style="border-collapse:collapse;font-size:13px">
<tr style="background:#f0f0f0"><th style="padding:5px 10px;border:1px solid #ddd;text-align:left">Date field (= 2025-11-12)</th><th style="padding:5px 10px;border:1px solid #ddd">Lines</th><th style="padding:5px 10px;border:1px solid #ddd">% of file</th><th style="padding:5px 10px;border:1px solid #ddd">SUM(PAID_AMOUNT)</th></tr>
<tr><td style="padding:5px 10px;border:1px solid #ddd">MOST_RECENT_PROCESS_DATE</td><td style="padding:5px 10px;border:1px solid #ddd;text-align:right">1,984,856</td><td style="padding:5px 10px;border:1px solid #ddd;text-align:right">100.0%</td><td style="padding:5px 10px;border:1px solid #ddd;text-align:right">$237,248,664.04</td></tr>
<tr><td style="padding:5px 10px;border:1px solid #ddd"><b>PAYABLE_DATE</b></td><td style="padding:5px 10px;border:1px solid #ddd;text-align:right"><b>1,868,215</b></td><td style="padding:5px 10px;border:1px solid #ddd;text-align:right">94.1%</td><td style="padding:5px 10px;border:1px solid #ddd;text-align:right"><b>$227,527,616.71</b></td></tr>
<tr><td style="padding:5px 10px;border:1px solid #ddd">CLEAN_CLAIM_DATE</td><td style="padding:5px 10px;border:1px solid #ddd;text-align:right">1,254,749</td><td style="padding:5px 10px;border:1px solid #ddd;text-align:right">63.2%</td><td style="padding:5px 10px;border:1px solid #ddd;text-align:right">$86,231,295.41</td></tr>
</table>
<p style="font-size:12px;color:#555">Other November payable dates in the file are negligible (a few hundred lines each), confirming this is a single 11/12 cycle, not a multi-day file.</p>

<h3 style="color:#2c5f8a;margin-bottom:4px">Source vs. what's loaded in AetnaHRP (TRGINTP3)</h3>
<table style="border-collapse:collapse;font-size:13px">
<tr style="background:#f0f0f0"><th style="padding:5px 10px;border:1px solid #ddd;text-align:left">Measure for 11/12/2025</th><th style="padding:5px 10px;border:1px solid #ddd">Lines</th><th style="padding:5px 10px;border:1px solid #ddd">Paid</th></tr>
<tr><td style="padding:5px 10px;border:1px solid #ddd">Source file &mdash; PAYABLE_DATE</td><td style="padding:5px 10px;border:1px solid #ddd;text-align:right">1,868,215</td><td style="padding:5px 10px;border:1px solid #ddd;text-align:right">$227.53M</td></tr>
<tr><td style="padding:5px 10px;border:1px solid #ddd">DB history.claim &mdash; PAYABLE_DATE (all tapes)</td><td style="padding:5px 10px;border:1px solid #ddd;text-align:right">4,034,563</td><td style="padding:5px 10px;border:1px solid #ddd;text-align:right">$455.71M</td></tr>
<tr style="color:#c00000;font-weight:bold"><td style="padding:5px 10px;border:1px solid #ddd">DB &mdash; PAYMENT_DATE (Payment-feed PayDate, from yesterday)</td><td style="padding:5px 10px;border:1px solid #ddd;text-align:right">380,299</td><td style="padding:5px 10px;border:1px solid #ddd;text-align:right">$26.70M</td></tr>
</table>

<h3 style="color:#2c5f8a;margin-bottom:4px">Conclusion</h3>
<p>The 11/12/2025 <b>claim</b> data is healthy &mdash; present both in this source extract ($227.5M payable)
and in the database ($455.7M payable, all tapes). The deficit found yesterday is strictly on the
<b>PAYMENT_DATE</b> (remittance/payment) side: only $26.7M of ~$200M expected. Because this file is the
<b>Claim</b> feed and carries no PAYMENT_DATE, <b>it does not contain &mdash; and cannot backfill &mdash; the
missing 11/12/2025 PayDate data.</b></p>
<p><b>Recommendation:</b> to close the 11/12/2025 PayDate gap, obtain/re-load the corresponding
<b>11/12/2025 Payment (remittance) extract</b> (mapping Table 1100), not this Claims extract. The Claims
side for that date is already complete.</p>

<h3 style="color:#2c5f8a;margin-bottom:4px">How I found this</h3>
<ol style="margin-top:4px">
<li>Confirmed the file layout: 231 columns, no PAYMENT_DATE; pay/date fields are PAYABLE_DATE [50] and MOST_RECENT_PROCESS_DATE [116] (PAID_AMOUNT [46]).</li>
<li>Streamed the full 6.48 GB CSV through a proper CSV parser (fields contain embedded commas), tallying lines and SUM(PAID_AMOUNT) by date field for 2025-11-12, plus a November distribution.</li>
<li>Cross-checked against TRGINTP3.AetnaHRP (history.claim) on the same fields, and against yesterday's PAYMENT_DATE finding.</li>
<li>Used the Aetna HRP Medical Mapping Document (ETL 2) to confirm PayDate = client.Payment.PAYMENT_DATE and that Claim (1000) and Payment (1100) are separate raw files.</li>
</ol>
<p style="color:#999;font-size:11px">Reviewed 2026-06-18. File scanned in full (1,984,856 rows, 0 unparseable PAID_AMOUNT values).</p>
</body></html>"""

print(send(to='timothy.stroud@machinify.com',
           subject='Review: VENDOR.CB-CLAIMS-EXTRACT 11/12/2025 - Claims cycle present ($227.5M), but no PAYMENT_DATE (PayDate gap is the Payment feed)',
           body=body))
