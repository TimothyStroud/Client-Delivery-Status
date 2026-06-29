"""
Ad-hoc Aetna RCE status poster (headless, zero Claude tokens).

Runs ramp_aetna_status_digest.py --force (bypasses the 25-min dedupe and the
Tue-Fri/8-12-16 gate that aetna_tick.py applies) and posts the resulting status
line to BOTH Slack channels via the Workflow Builder webhooks, using the same
sanitize + {"Text": ...} mechanism as aetna_webhook_post.py.

Used for one-off "post RCE status all day" requests on days the regular digest
schedule skips (e.g. Mondays). Driven by a temporary scheduled task; the digest's
own both-succeeded-today skip still applies (emits nothing if RCE + NCStateAetna
have both succeeded today).
"""
import sys, os, re, json, subprocess, urllib.request
from datetime import datetime

BASE = r'C:\Users\tls2\.claude\projects\H--'
DIGEST = os.path.join(BASE, 'ramp_aetna_status_digest.py')
SUPPORT_URL_FILE = r'H:\slack_wf_support.txt'
MINING_URL_FILE = r'H:\slack_wf_mining.txt'
LOG_FILE = r'H:\aetna_rce_adhoc.log'


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
    for label, path in (('support', SUPPORT_URL_FILE), ('mining', MINING_URL_FILE)):
        url = read_url(path)
        if not url:
            log(f"SKIP {label}: no webhook URL at {path}")
            continue
        try:
            st, body = post(url, txt)
            log(f"posted RCE status -> {label}: HTTP {st} {body[:50]}")
        except Exception as e:
            log(f"ERROR -> {label}: {e}")
    return 0


if __name__ == '__main__':
    sys.exit(main())
