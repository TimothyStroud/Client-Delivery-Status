"""
RAMP -> Slack monitor for the Aetna 0110 HRP Load job.

Event posted to #team-rdp-operations-support (C09EPLQL2D9):
  - 'Aetna 0110 HRP Load' (JobId 1246): when a run FINISHES *Failed* only
    (success completions are not posted, mirroring the Aetna RCE monitor).

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
STATE_FILE = os.path.join(BASE, 'ramp_aetnahrp_slack_state.json')
CHANNEL = 'C09EPLQL2D9'  # #team-rdp-operations-support

HRP_JOBID = 1246      # Aetna 0110 HRP Load -> on completion (Successful/Failed)

DONE_STATUSES = ('Successful', 'Failed')


def jobruns():
    out = subprocess.run(
        ['curl', '-s', '--negotiate', '-u', ':', 'http://ramp/api/Ramp/Job/List'],
        capture_output=True, text=True, timeout=180)
    data = json.loads(out.stdout)
    d = data['Data']
    jobs = d[0] if (isinstance(d, list) and d and isinstance(d[0], list)) else d
    runs = {}
    for j in jobs:
        if j.get('JobId') == HRP_JOBID:
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
    hrp = runs.get(HRP_JOBID, {})

    # HRP: announce ONLY a FAILED completion, once per QueueId (mirrors the RCE
    # monitor). Successful runs emit no SLACK line, so the poster never --commits
    # them; that's fine since a run's final status never flips and a later
    # failure carries a new QueueId.
    if hrp.get('EndDate') and hrp.get('Status') == 'Failed' \
            and hrp.get('QueueId') != state.get('hrp_last_completed_qid'):
        txt = ("<!here> :x: Aetna 0110 HRP Load - FAILED in RAMP\n"
               f"QueueId {hrp['QueueId']} | started {fmt(hrp.get('StartDate'))} | "
               f"ended {fmt(hrp.get('EndDate'))} - please investigate")
        events.append(('hrp', txt))
    return events


def commit(runs, state):
    hrp = runs.get(HRP_JOBID, {})
    if hrp.get('EndDate') and hrp.get('Status') in DONE_STATUSES:
        state['hrp_last_completed_qid'] = hrp.get('QueueId')
    save_state(state)


def main():
    runs = jobruns()
    state = load_state()

    if '--baseline' in sys.argv:
        hrp = runs.get(HRP_JOBID, {})
        # Only suppress an HRP completion if the current latest run is ALREADY done.
        state['hrp_last_completed_qid'] = (hrp.get('QueueId')
            if (hrp.get('EndDate') and hrp.get('Status') in DONE_STATUSES) else None)
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
        print('HRP|' + json.dumps(runs.get(HRP_JOBID, {})))
    if not events:
        print('NO_EVENTS')


if __name__ == '__main__':
    main()
