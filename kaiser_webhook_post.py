"""
Headless Kaiser Claims poster (zero Claude tokens). Replaces the Claude-cron path.

Runs kaiser_claims_slack.py (prints one 'SLACK|<text>' line with the 7-feed
most-recent-Claims summary) and posts it to the SUPPORT channel via the Slack
Workflow Builder webhook. Emoji render; *bold*/backticks/<!here> are stripped
(workflow shows them literally). POST body key = "Text". URL off the git repo at
H:\slack_wf_support.txt.
"""
import sys, os, re, json, subprocess, urllib.request
from datetime import datetime

BASE = r'C:\Users\tls2\.claude\projects\H--'
PY = sys.executable
KAISER = r'H:\kaiser_claims_slack.py'
SUPPORT_URL_FILE = r'H:\slack_wf_support.txt'
LOG_FILE = r'H:\kaiser_webhook_post.log'


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
        body = r.read().decode('utf-8', 'replace').strip()
        if r.status != 200:
            raise RuntimeError(f"HTTP {r.status}: {body[:200]}")
        return body


def main():
    try:
        support = open(SUPPORT_URL_FILE, encoding='utf-8').read().strip()
    except Exception:
        support = ''
    if not support:
        log(f"INERT: no support webhook URL in {SUPPORT_URL_FILE}")
        return 0
    r = subprocess.run([PY, KAISER], capture_output=True, text=True, timeout=300)
    for line in r.stdout.splitlines():
        if line.startswith('SLACK|'):
            txt = sanitize(line[len('SLACK|'):].replace('\\n', '\n'))
            try:
                post(support, txt)
                log("posted Kaiser claims digest -> support")
            except Exception as e:
                log(f"post error: {e}")
            return 0
    log("no SLACK line produced; nothing posted")
    return 0


if __name__ == '__main__':
    sys.exit(main())
