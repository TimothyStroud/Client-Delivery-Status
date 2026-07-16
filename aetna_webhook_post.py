"""
Headless Aetna RCE poster (zero Claude tokens). Replaces the Claude-cron path.

Runs aetna_tick.py (fail-only monitor every tick + status digest at slots) and
posts its tagged output to Slack via a Workflow Builder webhook. As of
2026-07-14 ALL Aetna RCE posts route to #data-operations-aetna-updates ONLY
(moved off support + mining per user):
  POST_SUPPORT|<text>  -> aetna-updates channel   (RCE failure alert)
  POST_BOTH|<text>     -> aetna-updates channel   (status digest)
After a POST_SUPPORT posts OK, runs `aetna_tick.py --commit` (two-phase, so a
failed post retries next tick). Digest needs no commit (it self-dedupes).

Workflow Builder rendering quirk (verified 2026-06-26): the workflow renders
:emoji: shortcodes but NOT *bold*/`code`/<!here> (those show literally and
@here does NOT ping). So sanitize() strips *, `, blockquote '> ', and <!here>;
emoji shortcodes are kept. Per user, failure alerts lose the @here ping (option
B) — they still post a red :x: message.

Webhook URL lives OFF the git repo: H:\slack_wf_aetna_updates.txt.
The POST body key is "Text" (capital T — the Workflow Builder variable name).
"""
import sys, os, re, json, subprocess, urllib.request
from datetime import datetime

BASE = r'C:\Users\tls2\.claude\projects\H--'
PY = sys.executable
TICK = os.path.join(BASE, 'aetna_tick.py')
AETNA_URL_FILE = r'H:\slack_wf_aetna_updates.txt'
LOG_FILE = r'H:\aetna_webhook_post.log'


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
    """Workflow Builder renders mrkdwn (*bold* / _italic_) + :emoji: (confirmed
    2026-07-16), but <!here> shows literally and does not ping -> drop only that."""
    return text.replace('<!here> ', '').replace('<!here>', '')


def post(url, text):
    data = json.dumps({'Text': text}).encode('utf-8')
    req = urllib.request.Request(url, data=data, headers={'Content-Type': 'application/json'})
    with urllib.request.urlopen(req, timeout=30) as r:
        body = r.read().decode('utf-8', 'replace').strip()
        if r.status != 200:
            raise RuntimeError(f"HTTP {r.status}: {body[:200]}")
        return body


def main():
    updates = read_url(AETNA_URL_FILE)
    if not updates:
        log(f"INERT: no aetna-updates webhook URL in {AETNA_URL_FILE}")
        return 0

    r = subprocess.run([PY, TICK], capture_output=True, text=True, timeout=300)
    posted_support = False
    for line in r.stdout.splitlines():
        if line.startswith('POST_SUPPORT|'):
            txt = sanitize(line[len('POST_SUPPORT|'):].replace('\\n', '\n'))
            try:
                post(updates, txt)
                posted_support = True
                log("posted RCE FAILURE -> aetna-updates")
            except Exception as e:
                log(f"FAIL alert post error (will retry next tick): {e}")
        elif line.startswith('POST_BOTH|'):
            txt = sanitize(line[len('POST_BOTH|'):].replace('\\n', '\n'))
            try:
                post(updates, txt)
                log("posted digest -> aetna-updates")
            except Exception as e:
                log(f"digest post error: {e}")

    # Commit monitor state ONLY after a failure alert actually posted.
    if posted_support:
        c = subprocess.run([PY, TICK, '--commit'], capture_output=True, text=True, timeout=120)
        log(f"committed monitor state: {c.stdout.strip()[:120]}")
    return 0


if __name__ == '__main__':
    sys.exit(main())
