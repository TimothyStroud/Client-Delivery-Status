"""
RAMP -> Slack monitor for the Aetna RCE 310 ETL Load job.

Event posted to #team-rdp-operations-support (C09EPLQL2D9):
  - 'Aetna RCE 310 ETL Load' (JobId 2257): when a run FINISHES *Failed* only
    (success completions are not posted, per user 2026-06-26).

(NCStateAetna 0100 Delivery Ticket start check removed per user 2026-06-20.)

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
STATE_FILE = os.path.join(BASE, 'ramp_aetna_slack_state.json')
CHANNEL = 'C09EPLQL2D9'  # #team-rdp-operations-support

RCE_JOBID = 2257      # Aetna RCE 310 ETL Load -> on completion (Successful/Failed)

DONE_STATUSES = ('Successful', 'Failed')


def jobruns():
    out = subprocess.run(
        ['curl', '-s', '--ntlm', '-u', ':', 'http://ramp/api/Ramp/Job/List'],
        capture_output=True, text=True, timeout=180)
    data = json.loads(out.stdout)
    d = data['Data']
    jobs = d[0] if (isinstance(d, list) and d and isinstance(d[0], list)) else d
    runs = {}
    for j in jobs:
        if j.get('JobId') == RCE_JOBID:
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
    rce = runs.get(RCE_JOBID, {})

    # RCE: announce ONLY a FAILED completion, once per QueueId (per user
    # 2026-06-26: success completions are no longer posted). Successful runs emit
    # no SLACK line, so the cron never --commits them; that's fine since a run's
    # final status never flips and a later failure carries a new QueueId.
    if rce.get('EndDate') and rce.get('Status') == 'Failed' \
            and rce.get('QueueId') != state.get('rce_last_completed_qid'):
        txt = ("<!here> :x: Aetna RCE 310 ETL Load - FAILED in RAMP\n"
               f"QueueId {rce['QueueId']} | started {fmt(rce.get('StartDate'))} | "
               f"ended {fmt(rce.get('EndDate'))} - please investigate")
        events.append(('rce', txt))
    return events


def commit(runs, state):
    rce = runs.get(RCE_JOBID, {})
    if rce.get('EndDate') and rce.get('Status') in DONE_STATUSES:
        state['rce_last_completed_qid'] = rce.get('QueueId')
    save_state(state)


def main():
    runs = jobruns()
    state = load_state()

    if '--baseline' in sys.argv:
        rce = runs.get(RCE_JOBID, {})
        # Only suppress an RCE completion if the current latest run is ALREADY done.
        state['rce_last_completed_qid'] = (rce.get('QueueId')
            if (rce.get('EndDate') and rce.get('Status') in DONE_STATUSES) else None)
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
        print('RCE|' + json.dumps(runs.get(RCE_JOBID, {})))
    if not events:
        print('NO_EVENTS')


if __name__ == '__main__':
    main()
