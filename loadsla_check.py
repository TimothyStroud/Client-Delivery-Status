"""
Load-Completion SLA monitor (RAMP) -> email from DataOperations to RDPOperations.

Reports Success / Failed SLA when a monitored RAMP load job COMPLETES.

Schedule: task runs 8am Mon-Fri. Each client is reported ONCE per period, then
goes dormant until the next period starts (weekly -> Monday, monthly -> 1st
weekday). Per user 2026-06-26.

5-day-SLA clients (SLA met if latest run Successful AND finished within 5 days of
its start; Failed status or >5 days = Failed SLA):
    Caresource 0200 Load             (JobId 5003)  -- WEEKLY
    Wellmark 0210 Claims Load        (JobId 2097)  -- WEEKLY
    HealthSpring_FWA 0110 Claims Load (JobId 1814) -- MONTHLY (reclassified
        from weekly 2026-06-26; SLA still the 5-day load duration)

Monthly client (Optum 0110 PBM Load, first Friday):
    Load  (JobId 1844) + Snap (Optum 0200 PBM Start Snap, JobId 1950)
  -> Success requires: all 3 RAW files (RAW1/RAW2/RAW3) for the cycle present,
     AND both Load and Snap Successful. RAW1/2/3 verified via OptumPBMRx.etl.Tape
     on TRGETL3.

Data source: RAMP /api/Ramp/Job/List (includes LatestJobRun per job).
State (loadsla_state.json) records per client: last_qid + last_period.

Flags: --status (print, no send), --force (ignore state + period gate).
"""
import sys, os, json, subprocess
from datetime import datetime, timedelta

BASE = r'C:\Users\tls2\.claude\projects\H--'
sys.path.insert(0, BASE)
from send_via_outlook import send

STATE_FILE = os.path.join(BASE, 'loadsla_state.json')
FROM_ADDR = 'DataOperations@machinify.com'
TO_ADDR   = 'RDPOperations@machinify.com'

# 5-day-SLA jobs (duration = EndDate-StartDate). `cadence` drives the
# report-once-per-period gating (added 2026-06-26 per user): once a job's
# completion is reported for the current period, it goes dormant until the next
# period starts -- weekly restarts Monday, monthly restarts the 1st weekday of
# the month. The task runs 8am Mon-Fri, so the first run of a new period IS the
# Monday / 1st-weekday restart.
JOBS = [
    {'key': 'caresource',       'name': 'Caresource 0200 Load',             'jobid': 5003, 'sla_days': 5, 'cadence': 'weekly'},
    {'key': 'wellmark',         'name': 'Wellmark 0210 Claims Load',         'jobid': 2097, 'sla_days': 5, 'cadence': 'weekly'},
    {'key': 'healthspring_fwa', 'name': 'HealthSpring_FWA 0110 Claims Load', 'jobid': 1814, 'sla_days': 5, 'cadence': 'monthly'},
]


def period_key(cadence, today=None):
    """Identifier for the current SLA period. Weekly = ISO year-week (restarts
    Monday); monthly = year-month (restarts the 1st)."""
    today = today or datetime.now()
    if cadence == 'weekly':
        y, w, _ = today.isocalendar()
        return f"{y}-W{w:02d}"
    return f"{today.year}-{today.month:02d}"

OPTUM_ENABLED = True

# RAW1/2/3 are confirmed/loaded via the OptumPBMRx.etl.Tape table on TRGETL3
# (the persistent load record; raw files age off disk). A row with
# ProcessStatus=50 and non-null FileLoadDate = successfully loaded.
OPTUM = {
    'key': 'optum_pbm',
    'load_name': 'Optum 0110 PBM Load',  'load_jobid': 1844,
    'snap_name': 'Optum 0200 PBM Start Snap', 'snap_jobid': 1950,
    'raw_labels': ['RAW1', 'RAW2', 'RAW3'],
    'sql_server': 'TRGETL3', 'sql_db': 'OptumPBMRx',
}

GREEN = '#1a7f37'
RED   = '#c00000'


def load_jobruns():
    """Return {jobid: LatestJobRun dict} from RAMP Job/List."""
    out = subprocess.run(
        ['curl', '-s', '--ntlm', '-u', ':', 'http://ramp/api/Ramp/Job/List'],
        capture_output=True, text=True)
    data = json.loads(out.stdout)
    d = data['Data']
    jobs = d[0] if (isinstance(d, list) and d and isinstance(d[0], list)) else d
    runs = {}
    for j in jobs:
        runs[j['JobId']] = j.get('LatestJobRun')
    return runs


def parse_dt(s):
    return datetime.fromisoformat(s) if s else None


def load_state():
    if os.path.exists(STATE_FILE):
        try:
            return json.load(open(STATE_FILE))
        except Exception:
            return {}
    return {}


def save_state(state):
    json.dump(state, open(STATE_FILE, 'w'), indent=2)


def fmt(dt):
    return dt.strftime('%m/%d/%Y %I:%M %p') if dt else '&mdash;'


def fmt_dur(start, end):
    if not (start and end):
        return '&mdash;'
    secs = (end - start).total_seconds()
    h = int(secs // 3600); m = int((secs % 3600) // 60)
    return f'{h}h {m}m'


def weekly_email(job, run):
    start = parse_dt(run.get('StartDate'))
    end   = parse_dt(run.get('EndDate'))
    status = run.get('Status', '')
    sla = timedelta(days=job['sla_days'])
    within = (end - start) <= sla if (start and end) else False
    ok = (status == 'Successful') and within
    color = GREEN if ok else RED
    verdict = 'SUCCESS - Within SLA' if ok else 'FAILED SLA'
    reason = ''
    if not ok:
        if status != 'Successful':
            reason = f' (load status: {status})'
        elif not within:
            reason = f' (exceeded {job["sla_days"]}-day SLA)'

    body = f"""<html><body style="font-family:Calibri,Arial,sans-serif;font-size:14px;color:#222">
<p><b>SLA Update &ndash; {job['name']}</b><br>
{job['cadence'].title()} client &middot; {job['sla_days']}-day SLA from load start &middot; Source: RAMP Dashboard</p>

<p style="color:{color};font-weight:bold;font-size:16px">{verdict}{reason}</p>

<table style="border-collapse:collapse;font-size:13px">
<tr><td style="padding:5px 10px;border:1px solid #ddd;background:#f0f0f0"><b>Load Job</b></td><td style="padding:5px 10px;border:1px solid #ddd">{job['name']}</td></tr>
<tr><td style="padding:5px 10px;border:1px solid #ddd;background:#f0f0f0"><b>Status</b></td><td style="padding:5px 10px;border:1px solid #ddd;color:{color};font-weight:bold">{status}</td></tr>
<tr><td style="padding:5px 10px;border:1px solid #ddd;background:#f0f0f0"><b>Started</b></td><td style="padding:5px 10px;border:1px solid #ddd">{fmt(start)}</td></tr>
<tr><td style="padding:5px 10px;border:1px solid #ddd;background:#f0f0f0"><b>Completed</b></td><td style="padding:5px 10px;border:1px solid #ddd">{fmt(end)}</td></tr>
<tr><td style="padding:5px 10px;border:1px solid #ddd;background:#f0f0f0"><b>Duration</b></td><td style="padding:5px 10px;border:1px solid #ddd">{fmt_dur(start, end)}</td></tr>
<tr><td style="padding:5px 10px;border:1px solid #ddd;background:#f0f0f0"><b>SLA</b></td><td style="padding:5px 10px;border:1px solid #ddd">{job['sla_days']} days from load start</td></tr>
</table>
<p style="color:#999;font-size:11px">Automated load-completion SLA update generated from RAMP.</p>
</body></html>"""
    subj = f"SLA Update - {job['name']} - {'SUCCESS' if ok else 'FAILED SLA'}"
    return subj, body, ok


def optum_raw_status(cycle_date):
    """Return ({label: loaded_bool}, tag) for RAW1/2/3 of the cycle (MMDDYYYY),
    read from OptumPBMRx.etl.Tape on TRGETL3 (successfully-loaded files only)."""
    tag = cycle_date.strftime('%m%d%Y')
    found = {lbl: False for lbl in OPTUM['raw_labels']}
    query = ("SET NOCOUNT ON; SELECT FileName FROM etl.Tape "
             "WHERE FileName LIKE '%RAW[123][_]%' AND FileName NOT LIKE '%RAWLINGS%' "
             "AND ProcessStatus=50 AND FileLoadDate IS NOT NULL;")
    out = subprocess.run(
        ['sqlcmd', '-S', OPTUM['sql_server'], '-d', OPTUM['sql_db'], '-E', '-W', '-h', '-1', '-Q', query],
        capture_output=True, text=True)
    for line in out.stdout.splitlines():
        line = line.strip()
        for lbl in OPTUM['raw_labels']:
            if (lbl + '_' + tag) in line:
                found[lbl] = True
    return found, tag


def optum_email(load_run, snap_run, raw_found, tag=''):
    ls, le = parse_dt(load_run.get('StartDate')), parse_dt(load_run.get('EndDate'))
    ss, se = parse_dt(snap_run.get('StartDate')), parse_dt(snap_run.get('EndDate'))
    lstatus = load_run.get('Status', ''); sstatus = snap_run.get('Status', '')
    all_raw = all(raw_found.values())
    ok = (lstatus == 'Successful') and (sstatus == 'Successful') and all_raw
    color = GREEN if ok else RED
    verdict = 'SUCCESS - Within SLA' if ok else 'FAILED SLA'

    raw_rows = ""
    for lbl in OPTUM['raw_labels']:
        present = raw_found[lbl]
        c = GREEN if present else RED
        raw_rows += (f'<tr><td style="padding:5px 10px;border:1px solid #ddd">{lbl}</td>'
                     f'<td style="padding:5px 10px;border:1px solid #ddd;color:{c};font-weight:bold">'
                     f'{"Present" if present else "MISSING"}</td></tr>')

    def jrow(label, status, s, e):
        c = GREEN if status == 'Successful' else RED
        return (f'<tr><td style="padding:5px 10px;border:1px solid #ddd">{label}</td>'
                f'<td style="padding:5px 10px;border:1px solid #ddd;color:{c};font-weight:bold">{status}</td>'
                f'<td style="padding:5px 10px;border:1px solid #ddd">{fmt(s)}</td>'
                f'<td style="padding:5px 10px;border:1px solid #ddd">{fmt(e)}</td></tr>')

    body = f"""<html><body style="font-family:Calibri,Arial,sans-serif;font-size:14px;color:#222">
<p><b>SLA Update &ndash; {OPTUM['load_name']}</b><br>
Monthly client (first Friday) &middot; Success = RAW1/2/3 loaded + Snap complete &middot; Source: RAMP Dashboard</p>

<p style="color:{color};font-weight:bold;font-size:16px">{verdict}</p>

<p style="font-weight:bold;margin-bottom:4px">RAW files (cycle {tag}, via OptumPBMRx.etl.Tape)</p>
<table style="border-collapse:collapse;font-size:13px">
<tr style="background:#f0f0f0"><th style="padding:5px 10px;border:1px solid #ddd;text-align:left">Raw</th><th style="padding:5px 10px;border:1px solid #ddd;text-align:left">Status</th></tr>
{raw_rows}
</table>

<p style="font-weight:bold;margin:12px 0 4px">Load &amp; Snap</p>
<table style="border-collapse:collapse;font-size:13px">
<tr style="background:#f0f0f0"><th style="padding:5px 10px;border:1px solid #ddd;text-align:left">Job</th><th style="padding:5px 10px;border:1px solid #ddd;text-align:left">Status</th><th style="padding:5px 10px;border:1px solid #ddd;text-align:left">Started</th><th style="padding:5px 10px;border:1px solid #ddd;text-align:left">Completed</th></tr>
{jrow(OPTUM['load_name'], lstatus, ls, le)}
{jrow(OPTUM['snap_name'], sstatus, ss, se)}
</table>
<p style="color:#999;font-size:11px">Automated load-completion SLA update generated from RAMP.</p>
</body></html>"""
    subj = f"SLA Update - {OPTUM['load_name']} ({le.strftime('%B %Y') if le else ''}) - {'SUCCESS' if ok else 'FAILED SLA'}"
    return subj, body, ok


def main():
    status_only = '--status' in sys.argv
    force = '--force' in sys.argv
    runs = load_jobruns()
    state = load_state()
    sent_any = False

    # ---- 5-day-SLA jobs (weekly + monthly), report-once-per-period ----
    for job in JOBS:
        run = runs.get(job['jobid'])
        if not run:
            print(f"[{job['key']}] no run data"); continue
        qid = run.get('QueueId'); status = run.get('Status'); end = run.get('EndDate')
        pk = period_key(job['cadence'])
        cell = state.get(job['key'], {})
        print(f"[{job['key']}] ({job['cadence']}) QueueId={qid} Status={status} End={end} period={pk}")
        if status_only:
            continue
        if not force:
            if cell.get('last_period') == pk:
                print("   already reported this period -- dormant until restart"); continue
            if cell.get('last_qid') == qid:
                print("   this completion already reported -- skip"); continue
        if not end:
            print("   still running / not complete -- skip"); continue
        subj, body, ok = weekly_email(job, run)
        res = send(to=TO_ADDR, subject=subj, body=body, from_address=FROM_ADDR)
        print(f"   send -> {res} | {subj}")
        if res == 'Sent.':
            cell['last_qid'] = qid; cell['last_period'] = pk
            state[job['key']] = cell
            sent_any = True

    # ---- Optum monthly ----
    lrun = runs.get(OPTUM['load_jobid']); srun = runs.get(OPTUM['snap_jobid'])
    if lrun and srun:
        lqid = lrun.get('QueueId'); lend = lrun.get('EndDate'); send_end = srun.get('EndDate')
        cycle = parse_dt(lrun.get('StartDate'))
        if cycle:
            raw_found, tag = optum_raw_status(cycle)
        else:
            raw_found, tag = {l: False for l in OPTUM['raw_labels']}, '?'
        print(f"[optum_pbm] LoadQ={lqid} LoadEnd={lend} SnapEnd={send_end} tag={tag} RAW={raw_found}")
        complete = bool(lend and send_end)
        pk = period_key('monthly')  # Optum is monthly: report once/month, restart 1st weekday
        cell = state.get(OPTUM['key'], {})
        if not status_only and not OPTUM_ENABLED:
            print("   OPTUM_ENABLED=False -- not sending (pending RAW1/2/3 confirmation)")
        elif not status_only:
            if not force and cell.get('last_period') == pk:
                print("   already reported this month -- dormant until restart")
            elif not complete:
                print("   load/snap not both complete -- skip")
            elif not force and cell.get('last_qid') == lqid:
                print("   this completion already reported -- skip")
            else:
                subj, body, ok = optum_email(lrun, srun, raw_found, tag)
                res = send(to=TO_ADDR, subject=subj, body=body, from_address=FROM_ADDR)
                print(f"   send -> {res} | {subj}")
                if res == 'Sent.':
                    cell['last_qid'] = lqid; cell['last_period'] = pk
                    state[OPTUM['key']] = cell
                    sent_any = True

    if not status_only and sent_any:
        save_state(state)
        print("State saved.")


if __name__ == '__main__':
    main()
