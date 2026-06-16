"""One-off: send EmblemFacets SLA Update review email to Tim's Outlook."""
import sys
sys.path.insert(0, r'C:\Users\tls2\.claude\projects\H--')
from send_via_outlook import send

rows = [
    ("RESP_CCI_RAWLINGS_CCOMM_EMPGRP_FA_FM01_PROD_*", "Uploaded", "6/16/2026 8:36 AM", "#1a7f37"),
    ("RESP_EH_RAWLINGS_GPPO_GRP_FA_FM01_PROD_*",      "Uploaded", "6/16/2026 8:37 AM", "#1a7f37"),
    ("RESP_CCI_RAWLINGS_MULTI_MEMELIG_FA_FM01_PROD_*", "Pending",  "Inbound received 6/13; awaiting load", "#9a6700"),
    ("RESP_EH_RAWLINGS_GPPO_MEMELIG_FA_FM01_PROD_*",   "Pending",  "Inbound received 6/13; awaiting load", "#9a6700"),
    ("RESP_RWGS_CEE_FF_*",                             "Pending",  "Inbound received 6/13; awaiting load", "#9a6700"),
]

tr = ""
for fname, status, detail, color in rows:
    tr += (
        f'<tr>'
        f'<td style="padding:6px 10px;border:1px solid #ddd;font-family:Consolas,monospace;font-size:12px">{fname}</td>'
        f'<td style="padding:6px 10px;border:1px solid #ddd;color:{color};font-weight:bold">{status}</td>'
        f'<td style="padding:6px 10px;border:1px solid #ddd">{detail}</td>'
        f'</tr>'
    )

body = f"""<html><body style="font-family:Calibri,Arial,sans-serif;font-size:14px;color:#222">
<p><b>SLA Update &ndash; EmblemFacets Response Files</b><br>
Monthly cycle (data dated 2026-06-13) &middot; 3-day SLA &middot; Source: RAMP Dashboard</p>

<p>Status as of <b>6/16/2026</b>: <b>2 of 5</b> Emblem Facets response files have been
successfully uploaded. The remaining 3 inbound files have all been received and are
queued behind the <i>Emblem Facets 0110 Load</i> job (currently running); they remain
<b>within the 3-day SLA window</b>.</p>

<table style="border-collapse:collapse;font-size:13px">
<tr style="background:#f0f0f0">
  <th style="padding:6px 10px;border:1px solid #ddd;text-align:left">Response File</th>
  <th style="padding:6px 10px;border:1px solid #ddd;text-align:left">Status</th>
  <th style="padding:6px 10px;border:1px solid #ddd;text-align:left">Detail</th>
</tr>
{tr}
</table>

<p style="margin-top:14px"><b>Pipeline:</b> Emblem Facets 0100 Stage (complete) &rarr;
Emblem Facets 0110 Load (running) &rarr; SFTP Emblem Upload (2 files uploaded 6/16).</p>

<p style="color:#666;font-size:12px">Uploaded response files are confirmed in
\\\\trgllc\\Shares\\RawlingsOutbound\\DIG\\EmblemFacets\\Archive and delivered to Emblem
(/Home/Rawlings/Prod/ToEmblem). A follow-up update will confirm the remaining 3 once the
Load completes and they upload.</p>

<p style="color:#999;font-size:11px">Automated SLA update generated from RAMP.</p>
</body></html>"""

result = send(
    to='timothy.stroud@machinify.com',
    subject='SLA Update - EmblemFacets Response Files (June 2026) - 2 of 5 Uploaded',
    body=body,
)
print(result)
