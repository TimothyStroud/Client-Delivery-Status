"""
RAMP -> Slack monitor for the Aetna 0110 Subro Load job.

Event posted to #data-operations-aetna-updates (via the aetna-updates Workflow
webhook, same destination as the RCE/HRP/Rx monitors):
  - 'Aetna 0110 Subro Load' (JobId 2242): when a run FINISHES *Failed* only
    (success completions are not posted, mirroring the other Aetna monitors).

Data source: RAMP /api/Ramp/Job/List (LatestJobRun per job).

Two-phase to avoid lost alerts if a Slack post fails:
  - default run  -> prints events as 'SLACK|<text>' lines; does NOT change state.
  - --commit     -> records the current QueueId to state (call only AFTER posting).
  - --baseline   -> seeds state to current so pre-existing runs aren't announced.
  - --status     -> prints current detection without posting/committing.
"""
import sys, os, json, subprocess
from datetime import datetime

BASE = os.path.dirname(os.path.abspath(__file__))
STATE_FILE = os.path.join(BASE, 'ramp_aetnasubro_slack_state.json')

SUBRO_JOBID = 2242    # Aetna 0110 Subro Load -> alert on a FAILED completion

# 'Resolved' shows up on this job's queue history when an operator manually
# resolves a run (RAMP then re-queues a fresh QueueId). It is terminal for that
# QueueId, so it counts as done for commit purposes even though we never alert
# on it.
DONE_STATUSES = ('Successful', 'Failed', 'Resolved')


def jobruns():
    out = subprocess.run(
        ['curl', '-s', '--negotiate', '-u', ':', 'http://ramp/api/Ramp/Job/List'],
        capture_output=True, text=True, timeout=180)
    data = json.loads(out.stdout)
    d = data['Data']
    jobs = d[0] if (isinstance(d, list) and d and isinstance(d[0], list)) else d
    runs = {}
    for j in jobs:
        if j.get('JobId') == SUBRO_JOBID:
            runs[j['JobId']] = j.get('LatestJobRun') or {}
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
        return datetime.fromisoformat(iso).strftime('%m/%d/%Y %I:%M %p')
    except Exception:
        return iso or '?'


def detect(runs, state):
    """Return list of (key, text) events vs current state (no state change)."""
    events = []
    sub = runs.get(SUBRO_JOBID, {})

    # Announce ONLY a FAILED completion, once per QueueId (mirrors the HRP/RCE
    # monitors). Successful runs emit no SLACK line, so the poster never --commits
    # them; that's fine since a run's final status never flips and a later failure
    # carries a new QueueId.
    if sub.get('EndDate') and sub.get('Status') == 'Failed' \
            and sub.get('QueueId') != state.get('subro_last_completed_qid'):
        txt = ("<!here> :x: Aetna 0110 Subro Load - FAILED in RAMP\n"
               f"QueueId {sub['QueueId']} | started {fmt(sub.get('StartDate'))} | "
               f"ended {fmt(sub.get('EndDate'))} - please investigate")
        events.append(('subro', txt))
    return events


def commit(runs, state):
    sub = runs.get(SUBRO_JOBID, {})
    if sub.get('EndDate') and sub.get('Status') in DONE_STATUSES:
        state['subro_last_completed_qid'] = sub.get('QueueId')
    save_state(state)


def main():
    runs = jobruns()
    state = load_state()

    if '--baseline' in sys.argv:
        sub = runs.get(SUBRO_JOBID, {})
        # Only suppress a completion if the current latest run is ALREADY done.
        state['subro_last_completed_qid'] = (sub.get('QueueId')
            if (sub.get('EndDate') and sub.get('Status') in DONE_STATUSES) else None)
        save_state(state)
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
        print('SUBRO|' + json.dumps(runs.get(SUBRO_JOBID, {})))
    if not events:
        print('NO_EVENTS')


if __name__ == '__main__':
    main()
