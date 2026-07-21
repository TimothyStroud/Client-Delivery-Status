import json, subprocess, sys
from datetime import datetime

STATE_FILE = r'C:\Users\tls2\.claude\projects\H--\known_unconfigured.json'
TO_ADDRESS   = 'DataOperations@machinify.com; RDPOperations@Machinify.com'
FROM_ADDRESS = 'DataOperations@machinify.com'

# Fetch current unconfigured files
result = subprocess.run(
    ['curl', '-s', '--negotiate', '-u', ':', 'http://ramp/api/Ramp/ConfiguredFiles'],
    capture_output=True, text=True
)

try:
    data = json.loads(result.stdout)
except json.JSONDecodeError:
    sys.exit(0)

items = data.get('Data', [[]])[0]
current = [i for i in items if not i.get('IsConfigured')]

# Load known state
try:
    with open(STATE_FILE) as f:
        known = {i['Path'] + '|' + i['File'] for i in json.load(f)}
except Exception:
    known = set()

new_files = [i for i in current if i['Path'] + '|' + i['File'] not in known]

if not new_files:
    sys.exit(0)

# Update state file
with open(STATE_FILE, 'w') as f:
    json.dump(
        [{'Path': i['Path'], 'File': i['File'], 'CreateDate': i.get('CreateDate', '')}
         for i in current],
        f
    )

# Build HTML email body
import re

def highlight_test(filename):
    return re.sub(r'(?i)(test)', r'<span style="background-color: yellow;">\1</span>', filename)

# When RAMP doesn't report a ClientName, infer it from the file's location/job.
# Folder slugs (lowercase) -> display names matching the real ClientName values
# so inferred files merge into the same client group.
CLIENT_DISPLAY = {
    'bcbsarrx': 'BCBSARRx',
    'bcbsvt': 'BCBSVT',
    'excellus': 'Excellus',
    'oscar': 'Oscar',
    # Point32Health is the parent of Harvard Pilgrim & Tufts. Per preference,
    # surface the brand names rather than the corporate parent; refine to a
    # specific brand when a token identifies one (see _point32_brand).
    'p32health': 'Harvard Pilgrim / Tufts',
    'anthem': 'Anthem Wellpoint',
    'centene': 'Centene',
    'cigna': 'Cigna',
    'coventry': 'Coventry',
    'nyship': 'NyShip',
}

def _point32_brand(item):
    """Identify the specific Point32Health brand (Harvard Pilgrim vs Tufts)
    from any field, when possible. Returns None if it can't be determined."""
    fl = item.get('FileLog') or {}
    blob = ' '.join(str(x) for x in [
        item.get('File'), item.get('Path'), fl.get('JobName'),
        fl.get('SourcePath'), fl.get('FeedName'),
    ]).lower()
    if 'tufts' in blob:
        return 'Tufts Health Plan'
    if 'harvard' in blob or 'pilgrim' in blob:
        return 'Harvard Pilgrim'
    return None

def _slug_from_clients_path(p):
    """Pull the client folder out of a '...\\clients\\<client>\\...' path."""
    if not p:
        return None
    m = re.search(r'[\\/]clients[\\/]([^\\/]+)', p, re.I)
    if not m:
        return None
    seg = m.group(1).strip().lower()
    seg = re.sub(r'\s+incoming$', '', seg)   # 'centene incoming' -> 'centene'
    return seg or None

def _slug_from_job(j):
    """Pull a client hint out of a job name like 'Excellus Rx Logfile'."""
    if not j:
        return None
    s = re.sub(r'(?i)\b(incoming|log\s*file|logfile|rx)\b', '', j)
    s = s.strip()
    return s.lower() or None

def derive_client(item):
    """Return (client_name, inferred_bool). Uses ClientName when present;
    otherwise falls back to Current Path, then SourcePath, then JobName."""
    fl = item.get('FileLog') or {}
    cn = fl.get('ClientName')
    if cn:
        return cn, False
    for src in (item.get('Path'), fl.get('SourcePath')):
        slug = _slug_from_clients_path(src)
        if slug:
            if slug == 'p32health':
                return _point32_brand(item) or CLIENT_DISPLAY['p32health'], True
            return CLIENT_DISPLAY.get(slug, slug.title()), True
    slug = _slug_from_job(fl.get('JobName'))
    if slug:
        if slug == 'p32health':
            return _point32_brand(item) or CLIENT_DISPLAY['p32health'], True
        return CLIENT_DISPLAY.get(slug, slug.title()), True
    # Last resort: a brand token may still appear even without a clients\ path.
    brand = _point32_brand(item)
    if brand:
        return brand, True
    return 'Unknown', False

by_client = {}
any_inferred = False
for f in new_files:
    client, inferred = derive_client(f)
    f['_inferred_client'] = inferred
    any_inferred = any_inferred or inferred
    by_client.setdefault(client, []).append(f)

rows = []
for client, files in sorted(by_client.items()):
    rows.append(f'<tr><td colspan="3" style="padding-top:10px;font-weight:bold;">{client}</td></tr>')
    for fi in files:
        fl = fi.get('FileLog') or {}
        ts = fi.get('CreateDate', '')
        try:
            dt = datetime.fromisoformat(ts).strftime('%m/%d %I:%M %p')
        except Exception:
            dt = ts
        fname = highlight_test(fi['File'])
        if fi.get('_inferred_client'):
            fname += ' <span style="color:#c47f00;" title="client inferred from path/job/source">&dagger;</span>'
        job = fl.get('JobName') or 'N/A'
        source = fl.get('SourcePath') or fi.get('Path') or 'N/A'
        rows.append(
            f'<tr><td style="padding-left:16px;padding-right:24px;">{fname}</td>'
            f'<td style="padding-right:24px;color:#555;">{dt}</td>'
            f'<td style="color:#555;">{job}<br><span style="font-size:0.9em;">{source}</span></td></tr>'
        )

body = f"""
<html><body style="font-family:Calibri,Arial,sans-serif;font-size:14px;">
<p>RAMP detected <strong>{len(new_files)}</strong> new unconfigured file(s) as of {datetime.now().strftime("%m/%d/%Y %I:%M %p")}.</p>
<table cellpadding="4" cellspacing="0" style="border-collapse:collapse;">
{''.join(rows)}
</table>
{'<p style="color:#c47f00;font-size:0.9em;">&dagger; Client name not reported by RAMP &mdash; inferred from the file&rsquo;s path / job / source.</p>' if any_inferred else ''}
<br><p style="color:#555;">Log in to RAMP to configure: <a href="http://ramp/Ramp/UnconfiguredFiles">View Unconfigured Files</a></p>
</body></html>
"""
subject = f'RAMP Alert: {len(new_files)} New Unconfigured File(s)'

# Send via interactive Outlook task
import sys, os
sys.path.insert(0, r'C:\Users\tls2\.claude\projects\H--')
from send_via_outlook import send
result = send(TO_ADDRESS, subject, body, from_address=FROM_ADDRESS)
print(result)
