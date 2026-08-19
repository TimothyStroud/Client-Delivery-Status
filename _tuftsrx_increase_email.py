import sys
sys.path.insert(0, r'C:\Users\tls2\.claude\projects\H--')
import send_via_outlook

body = """<html><body style="font-family:Segoe UI,Arial,sans-serif;font-size:13px;color:#222">

<h2 style="margin-bottom:2px">TuftsRx &ndash; Why Records / Paid Claims Jumped (Aug 17 2026 load)</h2>
<p style="margin-top:0;color:#666">Source: TRGINTP3.TuftsRx &mdash; etl.tape, claim.TRGRxClaim, mining.RxClaim, history.Eligibility.
Current cycle = TapeID 11711&ndash;11715 (claim file TapeID <b>11712</b>), compared to the July / early-August weekly cycles.</p>

<h3 style="color:#0b5">Bottom line</h3>
<p>Nothing broke and nothing was double-loaded. <b>Point32Health started sending Harvard Pilgrim (HPHC) carriers in the weekly Rx claim extract for the first time on 8/17.</b>
Those new carriers account for essentially 100% of the record increase &mdash; and, because we have <b>no HPHC eligibility feed</b>, also for essentially 100% of the new Null Eligibility.</p>

<h3>1. Size of the change</h3>
<table cellpadding="5" cellspacing="0" border="1" style="border-collapse:collapse;font-size:12px">
<tr style="background:#eee"><th align="left">Claim file (TableID 1100)</th><th>File size</th><th>Records loaded</th><th>Paid claims<br>(TotalAmountPaid&gt;0)</th><th>Total paid $</th></tr>
<tr><td>11689 &nbsp;7/27/2026</td><td align="right">96.7 MB</td><td align="right">59,374</td><td align="right">21,309</td><td align="right">&mdash;</td></tr>
<tr><td>11691 &nbsp;8/03/2026</td><td align="right">95.5 MB</td><td align="right">58,664</td><td align="right">21,304</td><td align="right">&mdash;</td></tr>
<tr><td>11699 &nbsp;8/10/2026</td><td align="right">94.5 MB</td><td align="right">58,016</td><td align="right">20,840</td><td align="right">$6,915,141</td></tr>
<tr style="background:#fff3cd;font-weight:bold"><td>11712 &nbsp;8/17/2026</td><td align="right">279.7 MB</td><td align="right">171,801</td><td align="right">63,404</td><td align="right">$23,776,660</td></tr>
</table>
<p>Records <b>+113,785 (2.96x)</b>; paid claims <b>+42,564 (3.04x)</b>; the raw file itself is <b>2.96x larger</b>, so the extra volume was actually sent to us &mdash; it is not a load/duplication artifact.
(Confirmed: 0 duplicate, 0 reversal and 0 adjustment-flagged rows in the tape, and <code>mining.RxClaim</code> carries the same 171,801 rows.)</p>

<h3>2. What type of claim increased &mdash; new Carriers</h3>
<p>The breakdown by <code>Carrier</code> makes the cause unambiguous. The pre-existing Point32Health carriers are <b>flat</b>; five brand-new HPHC carriers appeared out of nowhere:</p>
<table cellpadding="5" cellspacing="0" border="1" style="border-collapse:collapse;font-size:12px">
<tr style="background:#eee"><th align="left">Carrier</th><th>8/10 (11699)</th><th>8/17 (11712)</th><th>No eligibility<br>match on 8/17</th><th>Paid claims<br>8/17</th></tr>
<tr><td>P32H1144</td><td align="right">52,719</td><td align="right">52,812</td><td align="right">905</td><td align="right">19,450</td></tr>
<tr><td>P32T7301</td><td align="right">5,293</td><td align="right">5,123</td><td align="right">0</td><td align="right">1,649</td></tr>
<tr><td>P32T7300</td><td align="right">4</td><td align="right">0</td><td align="right">0</td><td align="right">0</td></tr>
<tr style="background:#f8d7da"><td><b>HPHCCOMM</b> (HPHC Commercial)</td><td align="right">0</td><td align="right"><b>89,405</b></td><td align="right">89,405</td><td align="right">32,634</td></tr>
<tr style="background:#f8d7da"><td><b>HPHCCOMMS</b> (HPHC Commercial &ndash; secondary/self-funded)</td><td align="right">0</td><td align="right"><b>15,107</b></td><td align="right">15,107</td><td align="right">5,857</td></tr>
<tr style="background:#f8d7da"><td><b>HPHCHIXME</b> (HPHC Exchange &ndash; Maine)</td><td align="right">0</td><td align="right"><b>4,828</b></td><td align="right">4,828</td><td align="right">2,189</td></tr>
<tr style="background:#f8d7da"><td><b>HPHCHIXMA</b> (HPHC Exchange &ndash; Massachusetts)</td><td align="right">0</td><td align="right"><b>4,086</b></td><td align="right">4,086</td><td align="right">1,372</td></tr>
<tr style="background:#f8d7da"><td><b>HPHCHIXNH</b> (HPHC Exchange &ndash; New Hampshire)</td><td align="right">0</td><td align="right"><b>440</b></td><td align="right">440</td><td align="right">253</td></tr>
<tr style="background:#eee;font-weight:bold"><td>Total new HPHC</td><td align="right">0</td><td align="right">113,866</td><td align="right">113,866</td><td align="right">42,305</td></tr>
</table>
<p><b>113,866 of the 113,785 net new records</b> are the new HPHC carriers, and <b>42,305 of the +42,564 paid claims (99.4%)</b> are HPHC.
They bring <b>$16,541,236</b> in new paid dollars, spread across <b>3,919 distinct Accounts</b> &mdash; versus only <b>14</b> Accounts in the entire prior week. So this is a new book of business (thousands of Harvard Pilgrim employer groups), not a spike in an existing group.</p>

<p style="background:#f4f4f4;padding:8px"><b>Not a backfill.</b> By DateFilled, 164,595 of the 171,801 rows are August 2026 fills and 6,961 are July 2026 &mdash; the same recent-service-date profile as prior weeks, just ~3x the volume.
All 171,801 rows carry a DatePaid in August 2026. Only ~120 rows have DateFilled older than 2026-02, so there is no material retro/history dump.
Other attributes are unchanged: RecordSource is still <code>TuftsRx</code>, DatabaseIndicator still <code>Optum</code>, PartDContractNumber still null (no Part D), and OtherCoverageCode is still 99.7% "00".</p>

<h3>3. Why Null Eligibility exploded</h3>
<table cellpadding="5" cellspacing="0" border="1" style="border-collapse:collapse;font-size:12px">
<tr style="background:#eee"><th align="left">Claim tape</th><th>Records</th><th>No eligibility match<br>(cEligibilityOrder NULL)</th><th>% of tape</th></tr>
<tr><td>11689 &nbsp;7/27</td><td align="right">59,374</td><td align="right">12</td><td align="right">0.02%</td></tr>
<tr><td>11691 &nbsp;8/03</td><td align="right">58,664</td><td align="right">42</td><td align="right">0.07%</td></tr>
<tr><td>11699 &nbsp;8/10</td><td align="right">58,016</td><td align="right">767</td><td align="right">1.3%</td></tr>
<tr style="background:#fff3cd;font-weight:bold"><td>11712 &nbsp;8/17</td><td align="right">171,801</td><td align="right">114,771</td><td align="right">66.8%</td></tr>
</table>
<p><b>113,866 of the 114,771 unmatched rows (99.2%) are the new HPHC carriers &mdash; and every single HPHC claim is unmatched (100%).</b>
The remaining 905 are P32H1144, a mild continuation of the 767 seen the week before (i.e. normal member-lag noise, not the story).
The same 114,771 rows carry no subscriber and no demographic match either (cSubscriberOrder / cEligibilityDemo are equally null), which is the signature of a population we have simply never been given eligibility for.</p>

<p><b>Root cause:</b> the only eligibility sources loaded in this database are <code>TuftsCIT_HealthPlan</code>, <code>Tufts_MedPref</code> and <code>Tufts_PublicPlan</code>
(files <code>Rawlings_active_eligible</code> / <code>Rawlings_cob_mem_chg_add</code>, <code>SHmember*.xtr</code>, <code>THPP_Rawlings_Member_Group_Extract</code>).
<b>There is no Harvard Pilgrim / HPHC eligibility feed at all.</b> Two things compound it in this cycle:</p>
<ul>
<li><b>No HPHC eligibility file exists</b> for any of the five new carriers &mdash; nothing in etl.tape has ever matched them.</li>
<li><b>This cycle also received a lighter-than-normal eligibility drop.</b> Cycle 11711&ndash;11715 loaded only <code>Rawlings_active_eligible_20260810.dat</code> (110 rows) and <code>Rawlings_cob_mem_chg_add.dat</code> (225 rows). The larger <code>SHmember</code> / <code>THPP</code> extracts that arrived in the 8/4 and 8/11 cycles (2,885 / 23,844 and 1,244 / 79,940 rows) did <b>not</b> arrive with the 8/17 claim file.</li>
</ul>

<h3>4. Recommended next steps</h3>
<ol>
<li>Confirm with Point32Health / the client that the HPHC (Harvard Pilgrim) book was <b>intentionally</b> added to the <code>Point32Health_Rawlings_6072_*_RXECHF70CL</code> extract effective 8/17 &mdash; and whether it is prospective-only or will also be backfilled.</li>
<li><b>Request the matching HPHC eligibility / group feed.</b> Without it, ~114K claims per week (2/3 of the file, ~$16.5M paid) are unmineable for COB.</li>
<li>Confirm whether the 3,919 new HPHC Accounts are in scope contractually, and get the BusLine / group key mapping loaded (<code>Rawlings_bus_line_keys.dat</code> is unchanged at 11,648 bytes, so it does not yet cover them).</li>
<li>Chase the missing <code>SHmember</code> / <code>THPP</code> eligibility extracts for the 8/17 cycle.</li>
<li>Expect the record/paid-claim volume and the Null Eligibility rate to <b>stay</b> at these levels going forward unless the client pulls HPHC back out &mdash; so any volume-based thresholds or SLA baselines for TuftsRx should be re-baselined.</li>
</ol>

<p style="color:#888;font-size:11px">Generated 8/19/2026 from TRGINTP3.TuftsRx.</p>
</body></html>"""

print(send_via_outlook.send('timothy.stroud@machinify.com',
    'TuftsRx - Cause of Record / Paid Claim Increase & Null Eligibility (Aug 17 2026 load)', body))
