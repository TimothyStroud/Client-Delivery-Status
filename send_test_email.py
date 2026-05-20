import json, subprocess
from datetime import datetime
import sys, os
sys.path.insert(0, r'C:\Users\tls2\.claude\projects\H--')
from send_via_outlook import send

STATE_FILE = r'C:\Users\tls2\.claude\projects\H--\known_unconfigured.json'
TO_ADDRESS = 'DataOperations@machinify.com; RDPOperations@machinify.com'

result = subprocess.run(
    ['curl', '-s', '--ntlm', '-u', ':', 'http://ramp/api/Ramp/ConfiguredFiles'],
    capture_output=True, text=True
)
data = json.loads(result.stdout)
items = data.get('Data', [[]])[0]
files = [i for i in items if not i.get('IsConfigured')]

import re

def highlight_test(filename):
    return re.sub(r'(?i)(test)', r'<span style="background-color: yellow;">\1</span>', filename)

by_client = {}
for f in files:
    client = (f.get('FileLog') or {}).get('ClientName') or 'Unknown'
    by_client.setdefault(client, []).append(f)

rows = []
for client, cfiles in sorted(by_client.items()):
    rows.append(f'<tr><td colspan="3" style="padding-top:10px;font-weight:bold;">{client} ({len(cfiles)})</td></tr>')
    for fi in cfiles:
        fl = fi.get('FileLog') or {}
        ts = fi.get('CreateDate', '')
        try:
            dt = datetime.fromisoformat(ts).strftime('%m/%d %I:%M %p')
        except Exception:
            dt = ts
        fname = highlight_test(fi['File'])
        job = fl.get('JobName', 'N/A')
        source = fl.get('SourcePath', 'N/A')
        rows.append(
            f'<tr><td style="padding-left:16px;padding-right:24px;">{fname}</td>'
            f'<td style="padding-right:24px;color:#555;">{dt}</td>'
            f'<td style="color:#555;">{job}<br><span style="font-size:0.9em;">{source}</span></td></tr>'
        )

body = f"""
<html><body style="font-family:Calibri,Arial,sans-serif;font-size:14px;">
<p>RAMP Unconfigured Files Summary as of {datetime.now().strftime("%m/%d/%Y %I:%M %p")} &mdash; <strong>{len(files)}</strong> total.</p>
<table cellpadding="4" cellspacing="0" style="border-collapse:collapse;">
{''.join(rows)}
</table>
<br><p style="color:#555;">Log in to RAMP to configure: <a href="http://ramp/Ramp/UnconfiguredFiles">View Unconfigured Files</a></p>
</body></html>
"""
subject = f'RAMP Unconfigured Files: {len(files)} Total'

# Update state baseline
with open(STATE_FILE, 'w') as fout:
    json.dump([{'Path': i['Path'], 'File': i['File'], 'CreateDate': i.get('CreateDate', '')} for i in files], fout)

result = send(TO_ADDRESS, subject, body)
print(result)
