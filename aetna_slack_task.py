"""
Standalone HEADLESS runner for the Aetna RCE 310 Slack monitor + digest.

Runs from a Windows Scheduled Task (no MCP, no Claude session). It calls the
existing detection scripts, then posts any resulting 'SLACK|' lines to Slack via
an Incoming Webhook URL.

Usage:
  python aetna_slack_task.py monitor   # event monitor: detect -> post -> commit
  python aetna_slack_task.py digest    # 3-hour status digest: build -> post (weekdays only)

The Incoming Webhook URL is read from aetna_slack_webhook.txt (same directory).
If that file is missing or empty the script logs and exits 0 (INERT) so the
scheduled task never errors before the URL has been provided.

Design notes:
  - Reuses ramp_aetna_slack_monitor.py / ramp_aetna_status_digest.py verbatim so
    detection logic lives in one place. This wrapper only handles posting.
  - Monitor is two-phase: only run --commit AFTER every event posted OK, so a
    failed Slack post is retried on the next tick instead of being silently lost.
  - Active-window gate (see active()): no Sunday file is loaded, so BOTH modes
    skip Sunday AND Monday entirely; the first run after Saturday is at/after 3am
    Tuesday. Monitor runs Tue-Sat; digest runs Tue-Fri (keeps its weekday-only
    footprint). The scheduled-task triggers still fire on their interval every
    day -- this gate decides which firings actually do work.
"""
import sys, os, json, subprocess, urllib.request
from datetime import datetime

BASE = os.path.dirname(os.path.abspath(__file__))
WEBHOOK_FILE = os.path.join(BASE, 'aetna_slack_webhook.txt')
LOG_FILE = os.path.join(BASE, 'aetna_slack_task.log')
MONITOR = os.path.join(BASE, 'ramp_aetna_slack_monitor.py')
DIGEST = os.path.join(BASE, 'ramp_aetna_status_digest.py')


def log(msg):
    ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    line = f"[{ts}] {msg}"
    try:
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(line + "\n")
    except Exception:
        pass
    print(line)


def get_webhook():
    if not os.path.exists(WEBHOOK_FILE):
        return None
    try:
        url = open(WEBHOOK_FILE, encoding='utf-8').read().strip()
    except Exception:
        return None
    return url or None


def post(url, text):
    payload = json.dumps({"text": text}).encode('utf-8')
    req = urllib.request.Request(url, data=payload,
                                 headers={'Content-Type': 'application/json'})
    with urllib.request.urlopen(req, timeout=30) as resp:
        body = resp.read().decode('utf-8', 'replace').strip()
        if resp.status != 200 or body != 'ok':
            raise RuntimeError(f"HTTP {resp.status}: {body[:200]}")


def active(mode, now):
    """Should this firing do work? Returns (bool, reason_if_skipped).

    No Sunday file is loaded, so both modes skip Sunday + Monday entirely and
    resume at 3am Tuesday. Monitor is active Tue-Sat; digest stays weekday-only
    (Tue-Fri). weekday(): Mon=0 .. Sun=6.
    """
    wd = now.weekday()
    if wd in (6, 0):                      # Sunday, Monday
        return False, "Sun/Mon skip (no Sunday file)"
    if wd == 1 and now.hour < 3:          # Tuesday before 3am
        return False, "Tue pre-3am skip (resume 3am Tue)"
    if wd == 5 and mode == 'digest':      # Saturday: digest is weekday-only
        return False, "Sat skip (digest weekday-only)"
    return True, None


def run_script(args):
    return subprocess.run([sys.executable] + args,
                          capture_output=True, text=True, timeout=300)


def slack_lines(stdout):
    out = []
    for line in stdout.splitlines():
        if line.startswith('SLACK|'):
            out.append(line[len('SLACK|'):].replace('\\n', '\n'))
    return out


def do_monitor(url):
    r = run_script([MONITOR])
    if r.returncode != 0:
        log(f"monitor detect FAILED rc={r.returncode}: {r.stderr.strip()[:300]}")
        return 1
    lines = slack_lines(r.stdout)
    if not lines:
        log("monitor: no events")
        return 0
    for txt in lines:
        try:
            post(url, txt)
            log(f"posted event ({len(txt)} chars)")
        except Exception as e:
            log(f"POST failed, NOT committing (retries next tick): {e}")
            return 1
    c = run_script([MONITOR, '--commit'])
    if c.returncode != 0:
        log(f"commit FAILED rc={c.returncode}: {c.stderr.strip()[:300]}")
        return 1
    log(f"committed: {c.stdout.strip()}")
    return 0


def do_digest(url):
    r = run_script([DIGEST])
    if r.returncode != 0:
        log(f"digest build FAILED rc={r.returncode}: {r.stderr.strip()[:300]}")
        return 1
    lines = slack_lines(r.stdout)
    if not lines:
        log("digest: no SLACK line produced")
        return 1
    try:
        post(url, lines[0])
        log("posted digest")
        return 0
    except Exception as e:
        log(f"digest POST failed: {e}")
        return 1


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else ''
    if mode not in ('monitor', 'digest'):
        log(f"usage: aetna_slack_task.py monitor|digest (got {mode!r})")
        return 2
    url = get_webhook()
    if not url:
        log(f"INERT: no webhook URL in {WEBHOOK_FILE}; skipping {mode}")
        return 0
    ok, reason = active(mode, datetime.now())
    if not ok:
        log(f"{mode}: outside active window ({reason}); skipping")
        return 0
    return do_monitor(url) if mode == 'monitor' else do_digest(url)


if __name__ == '__main__':
    sys.exit(main())
