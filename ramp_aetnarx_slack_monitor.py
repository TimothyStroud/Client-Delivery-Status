"""
RAMP -> Slack fail-only monitor for the AetnaRx Claim pipeline.

Mirrors ramp_aetnahrp_slack_monitor.py but watches EVERY RAMP job whose name
starts with 'AetnaRx Claim' (case-insensitive), discovered dynamically from
/api/Ramp/Job/List so new pipeline steps are picked up automatically.

Event posted to #data-operations-aetna-updates (via the aetna-updates webhook,
by aetnarx_webhook_post.py):
  - Any watched job whose LATEST run FINISHES *Failed* -> one alert per (JobId,
    QueueId). Successful completions are NOT posted (mirrors RCE/HRP).

Two-phase to avoid lost alerts if a Slack post fails:
  - default run  -> prints events as 'SLACK|<text>' lines; does NOT change state.
  - --commit     -> records current QueueIds to state (call only AFTER posting).
  - --baseline   -> seeds state to current so pre-existing runs aren't announced.
  - --status     -> prints current detection without posting/committing.
"""
import sys, os, json, subprocess
from datetime import datetime

BASE = os.path.dirname(os.path.abspath(__file__))
STATE_FILE = os.path.join(BASE, 'ramp_aetnarx_slack_state.json')
CHANNEL = 'data-operations-aetna-updates'  # posted via webhook, id not needed

JOB_PREFIX = 'aetnarx claim'   # case-insensitive name prefix of watched jobs
DONE_STATUSES = ('Successful', 'Resolved', 'Failed')


def all_jobs():
    out = subprocess.run(
        ['curl', '-s', '--ntlm', '-u', ':', 'http://ramp/api/Ramp/Job/List'],
        capture_output=True, text=True, timeout=180)
    data = json.loads(out.stdout)
    d = data['Data']
    return d[0] if (isinstance(d, list) and d and isinstance(d[0], list)) else d


def watched_runs():
    """{JobId: (JobName, LatestJobRun)} for jobs named 'AetnaRx Claim*'.
    RTA jobs are excluded per user (2026-07-16)."""
    runs = {}
    for j in all_jobs():
        name = (j.get('JobName') or j.get('Name') or '')
        if name.lower().startswith(JOB_PREFIX) and 'rta' not in name.lower():
            runs[j['JobId']] = (name, j.get('LatestJobRun') or {})
    return runs


def load_state():
    if os.path.exists(STATE_FILE):
        try:
            return json.load(open(STATE_FILE))
        except Exception:
            pass
    return {}


def save_state(s):
    json.dump(s, open(STATE_FILE, 'w'), indent=2)


def fmt(iso):
    try:
        return datetime.fromisoformat(str(iso).split('.')[0]).strftime('%m/%d/%Y %I:%M %p')
    except Exception:
        return iso or '?'


def detect(runs, state):
    """Return list of (jobid, text) FAILED events vs current state."""
    last = state.get('last_completed_qid', {})
    events = []
    for jobid, (name, lr) in sorted(runs.items()):
        if lr.get('EndDate') and lr.get('Status') == 'Failed' \
                and lr.get('QueueId') != last.get(str(jobid)):
            txt = (f"<!here> :x: {name} - FAILED in RAMP\n"
                   f"QueueId {lr['QueueId']} | started {fmt(lr.get('StartDate'))} | "
                   f"ended {fmt(lr.get('EndDate'))} - please investigate")
            events.append((jobid, txt))
    return events


def commit(runs, state):
    last = state.setdefault('last_completed_qid', {})
    for jobid, (name, lr) in runs.items():
        if lr.get('EndDate') and lr.get('Status') in DONE_STATUSES:
            last[str(jobid)] = lr.get('QueueId')
    save_state(state)


def baseline(runs, state):
    last = {}
    for jobid, (name, lr) in runs.items():
        if lr.get('EndDate') and lr.get('Status') in DONE_STATUSES:
            last[str(jobid)] = lr.get('QueueId')
    state['last_completed_qid'] = last
    save_state(state)


def main():
    runs = watched_runs()
    state = load_state()

    if '--baseline' in sys.argv:
        baseline(runs, state)
        print('Baselined:', json.dumps(state))
        return

    if '--commit' in sys.argv:
        commit(runs, state)
        print('Committed:', json.dumps(state))
        return

    events = detect(runs, state)
    for _, txt in events:
        print('SLACK|' + txt.replace('\n', '\\n'))
    if '--status' in sys.argv:
        print('STATE|' + json.dumps(state))
        for jobid, (name, lr) in sorted(runs.items()):
            print(f"JOB|{jobid}|{name}|{lr.get('Status')}|{lr.get('EndDate')}")
    if not events:
        print('NO_EVENTS')


if __name__ == '__main__':
    main()
