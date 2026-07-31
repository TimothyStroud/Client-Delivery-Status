"""
Headless poster for the weekly AetnaRx Eligibility file report (zero Claude
tokens). Mirrors aetnarx_webhook_post.py.

Runs ramp_aetnarx_elig_report.py, posts its `SLACK|<text>` line to
#data-operations-aetna-updates through the shared aetna-updates Workflow Builder
webhook, then (two-phase) re-invokes the report with --commit so a failed post is
retried on the next run instead of being silently marked done.

The webhook renders :emoji: shortcodes ONLY -- *bold*/`code`/<!here> show up
literally -- so sanitize() strips them (underscores are kept: they live inside
emoji shortcodes like :red_circle:).

Webhook URL lives OFF the git repo: H:\slack_wf_aetna_updates.txt.
The POST body key is "Text" (capital T -- the Workflow Builder variable name).

  --force   pass --force through to the report (bypasses the once-a-day dedupe)
"""
import sys, os, re, json, subprocess, urllib.request
from datetime import datetime

BASE = r'C:\Users\tls2\.claude\projects\H--'
PY = sys.executable
REPORT = os.path.join(BASE, 'ramp_aetnarx_elig_report.py')
AETNA_URL_FILE = r'H:\slack_wf_aetna_updates.txt'
LOG_FILE = r'H:\aetnarx_elig_webhook_post.log'


def log(msg):
    line = f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}"
    try:
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(line + "\n")
    except Exception:
        pass
    print(line)


def read_url(path):
    try:
        return open(path, encoding='utf-8').read().strip() or None
    except Exception:
        return None


def sanitize(text):
    text = text.replace('<!here> ', '').replace('<!here>', '')
    text = re.sub(r'(?m)^> ?', '', text)
    # NB: do NOT strip '_' -> it lives inside emoji shortcodes (:red_circle:).
    return text.replace('*', '').replace('`', '')


def post(url, text):
    data = json.dumps({'Text': text}).encode('utf-8')
    req = urllib.request.Request(url, data=data, headers={'Content-Type': 'application/json'})
    with urllib.request.urlopen(req, timeout=30) as r:
        body = r.read().decode('utf-8', 'replace').strip()
        if r.status != 200:
            raise RuntimeError(f"HTTP {r.status}: {body[:200]}")
        return body


def main():
    url = read_url(AETNA_URL_FILE)
    if not url:
        log(f"INERT: no aetna-updates webhook URL in {AETNA_URL_FILE}")
        return 0

    args = [PY, REPORT] + (['--force'] if '--force' in sys.argv else [])
    r = subprocess.run(args, capture_output=True, text=True, timeout=600)
    if r.returncode != 0:
        log(f"report failed: {(r.stderr or r.stdout).strip()[:300]}")
        return 1

    posted = False
    for line in r.stdout.splitlines():
        if line.startswith('SLACK|'):
            txt = sanitize(line[len('SLACK|'):].replace('\\n', '\n'))
            try:
                post(url, txt)
                posted = True
                log("posted AetnaRx Eligibility file report -> aetna-updates")
            except Exception as e:
                log(f"post error (will retry next run): {e}")

    if posted:
        c = subprocess.run([PY, REPORT, '--commit'], capture_output=True, text=True, timeout=60)
        log(f"committed: {c.stdout.strip()[:120]}")
    elif 'SKIP' in r.stdout:
        log("skipped (already posted today)")
    return 0


if __name__ == '__main__':
    sys.exit(main())
