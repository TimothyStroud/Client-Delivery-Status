"""
Ad-hoc AetnaRx Claim status poster (headless, zero Claude tokens).

Runs ramp_aetnarx_status_digest.py --force (bypasses the 25-min dedupe and the
Tue-Fri/8-12-16 slot gate that aetnarx_tick.py applies) and posts the resulting
status line to #data-operations-aetna-updates via the Workflow Builder webhook,
using the same sanitize + {"Text": ...} mechanism as aetnarx_webhook_post.py.

For one-off "post AetnaRx status now" requests. The digest's own skip still
applies (emits nothing once the Claim pipeline + SQL have all succeeded today).
"""
import sys, os, re, json, subprocess, urllib.request
from datetime import datetime

BASE = r'C:\Users\tls2\.claude\projects\H--'
DIGEST = os.path.join(BASE, 'ramp_aetnarx_status_digest.py')
AETNA_URL_FILE = r'H:\slack_wf_aetna_updates.txt'
LOG_FILE = r'H:\aetnarx_force_post.log'


def log(msg):
    line = f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}"
    try:
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(line + "\n")
    except Exception:
        pass
    print(line)


def sanitize(text):
    text = text.replace('<!here> ', '').replace('<!here>', '')
    text = re.sub(r'(?m)^> ?', '', text)
    text = text.replace('*', '').replace('`', '')
    return text


def post(url, text):
    data = json.dumps({'Text': text}).encode('utf-8')
    req = urllib.request.Request(url, data=data, headers={'Content-Type': 'application/json'})
    with urllib.request.urlopen(req, timeout=30) as r:
        return r.status, r.read().decode('utf-8', 'replace').strip()


def read_url(path):
    try:
        return open(path, encoding='utf-8').read().strip() or None
    except Exception:
        return None


def main():
    r = subprocess.run([sys.executable, DIGEST, '--force'],
                       capture_output=True, text=True, timeout=300)
    line = next((l for l in r.stdout.splitlines() if l.startswith('SLACK|')), None)
    if not line:
        log("no SLACK line (digest skipped: " + (r.stdout.strip()[:120] or "empty") + ")")
        return 0
    txt = sanitize(line[len('SLACK|'):].replace('\\n', '\n'))
    url = read_url(AETNA_URL_FILE)
    if not url:
        log(f"SKIP: no aetna-updates webhook URL at {AETNA_URL_FILE}")
        return 0
    try:
        st, body = post(url, txt)
        log(f"posted AetnaRx status -> aetna-updates: HTTP {st} {body[:50]}")
    except Exception as e:
        log(f"ERROR -> aetna-updates: {e}")
    return 0


if __name__ == '__main__':
    sys.exit(main())
