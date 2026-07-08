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
import sys, os, json, subprocess, re
from datetime import datetime, timedelta, date

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
# RAW1/2/3 trickle in over several days (RAW1 typically last), so don't report
# the moment load+snap finish. Wait until all 3 are loaded (-> SUCCESS); only if
# they're still incomplete this many days after the cycle (first Friday) do we
# report FAILED SLA. Prevents the premature-FAILED seen on 2026-07-01.
OPTUM_RAW_SLA_DAYS = 5

GREEN = '#1a7f37'
RED   = '#c00000'

# Statuses that count as a successful completion. Mirrors the Client Delivery
# Status report's acceptance: RAMP marks a run 'Resolved' (or a 'Success/*'
# variant) when it hit an issue that was manually fixed -- still a delivery.
SUCCESS_STATES = {'Successful', 'Success', 'Success/ManualFix',
                  'Success/NoWork', 'Resolved'}


def _ok_status(s):
    return (s or '').strip() in SUCCESS_STATES


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
    ok = _ok_status(status) and within
    color = GREEN if ok else RED
    verdict = 'SUCCESS - Within SLA' if ok else 'FAILED SLA'
    reason = ''
    if not ok:
        if not _ok_status(status):
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


# ProcessStatus -> human label (OptumPBMRx.etl.Tape). 50 = fully loaded.
OPTUM_PS_LABEL = {50: 'Loaded', 42: 'Loading', 32: 'Staging'}


def _ps_label(ps):
    return OPTUM_PS_LABEL.get(ps, f'ProcessStatus {ps}')


def _blank_raw():
    return {'loaded': False, 'status': 'MISSING', 'started': None, 'completed': None, 'file': None}


def optum_raw_status():
    """Return ({label: detail}, tag) for the LATEST cycle's RAW1/2/3, read from
    OptumPBMRx.etl.Tape on TRGETL3. `detail` = {'loaded': bool, 'status': str,
    'started': dt|None, 'completed': dt|None} where started=FileCreateDate,
    completed=FileLoadDate, loaded = ProcessStatus 50 with a non-null FileLoadDate.

    Rows at ANY ProcessStatus are read (not just loaded), so a RAW still in flight
    shows its current status/timestamps. The cycle tag (MMDDYYYY) is derived from
    the RAW filenames themselves, so it follows the actual data cycle rather than a
    re-running load job's StartDate (which drifts as RAW files trickle in)."""
    query = ("SET NOCOUNT ON; SELECT ProcessStatus, "
             "CONVERT(varchar(19), FileCreateDate, 120), "
             "CONVERT(varchar(19), FileLoadDate, 120), FileName FROM etl.Tape "
             "WHERE FileName LIKE '%RAW[123][_]%' AND FileName NOT LIKE '%RAWLINGS%';")
    out = subprocess.run(
        ['sqlcmd', '-S', OPTUM['sql_server'], '-d', OPTUM['sql_db'], '-E', '-W',
         '-h', '-1', '-s', '\t', '-Q', query],
        capture_output=True, text=True)
    rx = re.compile(r'(RAW[123])_(\d{8})', re.I)
    by_cycle = {}   # MMDDYYYY tag -> {label: (ps, started, completed)}
    for line in out.stdout.splitlines():
        parts = line.split('\t')
        if len(parts) < 4:
            continue
        ps_s, cre, ld, fname = (p.strip() for p in parts[:4])
        m = rx.search(fname)
        if not m:
            continue
        try:
            ps = int(ps_s)
        except ValueError:
            continue
        started = parse_dt(cre.replace(' ', 'T')) if cre and cre != 'NULL' else None
        completed = parse_dt(ld.replace(' ', 'T')) if ld and ld != 'NULL' else None
        base = fname.replace('/', '\\').split('\\')[-1]
        by_cycle.setdefault(m.group(2), {})[m.group(1).upper()] = (ps, started, completed, base)
    if not by_cycle:
        return {lbl: _blank_raw() for lbl in OPTUM['raw_labels']}, '?'

    def _cd(t):
        return date(int(t[4:8]), int(t[0:2]), int(t[2:4]))
    tag = max(by_cycle, key=_cd)          # most recent data cycle present
    rows = by_cycle[tag]
    detail = {}
    for lbl in OPTUM['raw_labels']:
        if lbl in rows:
            ps, started, completed, base = rows[lbl]
            detail[lbl] = {'loaded': (ps == 50 and completed is not None),
                           'status': _ps_label(ps), 'started': started,
                           'completed': completed, 'file': base}
        else:
            detail[lbl] = _blank_raw()
    return detail, tag


def optum_snap_runs(days=15):
    """Recent runs of the Optum snap job (JobId 1950) from the RAMP Queue SQL
    table (TRGUTIL10.RAMP.ramp.Queue) -- full history, since the REST endpoint
    rotates old runs out within hours. Returns [{'status','start','end'}] sorted
    by StartDate. Each RAW load is followed by its own snap, so we attribute a
    per-RAW snap = the first run that STARTS at/after that RAW's load completion."""
    q = ("SET NOCOUNT ON; SELECT q.Status, "
         "CONVERT(varchar(19),q.StartDate,120), CONVERT(varchar(19),q.EndDate,120) "
         "FROM [RAMP].[ramp].[Queue] q "
         f"WHERE q.JobId={OPTUM['snap_jobid']} "
         f"AND q.CreateDate >= DATEADD(day,-{days},GETDATE()) ORDER BY q.StartDate;")
    out = subprocess.run(
        ['sqlcmd', '-S', 'TRGUTIL10', '-d', 'RAMP', '-E', '-W', '-h', '-1', '-s', '\t', '-Q', q],
        capture_output=True, text=True)
    runs = []
    for line in out.stdout.splitlines():
        parts = line.split('\t')
        if len(parts) < 3:
            continue
        status, st, en = (p.strip() for p in parts[:3])
        if not st or st == 'NULL':
            continue
        runs.append({'status': status,
                     'start': parse_dt(st.replace(' ', 'T')),
                     'end': parse_dt(en.replace(' ', 'T')) if en and en != 'NULL' else None})
    runs.sort(key=lambda r: r['start'])
    return runs


def optum_email(load_run, snap_run, raw_found, tag=''):
    le = parse_dt(load_run.get('EndDate'))

    # SLA window: RAW1/2/3 due within OPTUM_RAW_SLA_DAYS of the cycle date (first Friday).
    try:
        cyc = datetime(int(tag[4:8]), int(tag[0:2]), int(tag[2:4]))
        sla_deadline = cyc + timedelta(days=OPTUM_RAW_SLA_DAYS)
        sla_range = f"{cyc:%m/%d/%Y} &ndash; {sla_deadline:%m/%d/%Y}"
    except (ValueError, IndexError):
        cyc, sla_deadline, sla_range = None, None, 'n/a'

    # Each RAW load is followed by its own snap -> attribute per-RAW snap = the
    # first snap run starting at/after that RAW's load completion.
    snap_runs = optum_snap_runs()

    def snap_for(completed):
        if not completed:
            return None
        for r in snap_runs:
            if r['start'] and r['start'] >= completed:
                return r
        return None

    # One merged row per RAW file: Job | RAW File | Load Start | Load Completion | Snap Date | SLA Status.
    td = 'padding:5px 10px;border:1px solid #ddd'
    merged_rows = ""
    row_oks = []
    for lbl in OPTUM['raw_labels']:
        d = raw_found[lbl]
        snap = snap_for(d['completed'])
        snap_end = snap['end'] if snap else None
        snap_ok = _ok_status(snap['status']) if snap else False
        within = bool(d['completed'] and sla_deadline and d['completed'] <= sla_deadline)
        row_ok = d['loaded'] and snap_ok and within
        row_oks.append(row_ok)
        sc = GREEN if row_ok else RED
        merged_rows += (
            f'<tr><td style="{td}">{OPTUM["load_name"]}</td>'
            f'<td style="{td}">{d["file"] or lbl}</td>'
            f'<td style="{td}">{fmt(d["started"])}</td>'
            f'<td style="{td}">{fmt(d["completed"])}</td>'
            f'<td style="{td}">{fmt(snap_end)}</td>'
            f'<td style="{td};color:{sc};font-weight:bold">{"Success" if row_ok else "Failed SLA"}</td></tr>')

    ok = bool(row_oks) and all(row_oks)
    color = GREEN if ok else RED
    verdict = 'SUCCESS - Within SLA' if ok else 'FAILED SLA'

    th = 'padding:5px 10px;border:1px solid #ddd;text-align:left'
    body = f"""<html><body style="font-family:Calibri,Arial,sans-serif;font-size:14px;color:#222">
<p><b>SLA Update &ndash; {OPTUM['load_name']}</b><br>
Monthly client (first Friday) &middot; Success = each RAW loaded + snapped within SLA &middot; Source: RAMP Dashboard</p>

<p style="color:{color};font-weight:bold;font-size:16px">{verdict}</p>

<p style="margin:0 0 12px"><b>SLA window (cycle {tag}):</b> {sla_range} &middot; RAW1/2/3 due within {OPTUM_RAW_SLA_DAYS} days of the cycle date (first Friday).</p>

<table style="border-collapse:collapse;font-size:13px">
<tr style="background:#f0f0f0"><th style="{th}">Job</th><th style="{th}">RAW File</th><th style="{th}">Load Start</th><th style="{th}">Load Completion</th><th style="{th}">Snap Date</th><th style="{th}">SLA Status</th></tr>
{merged_rows}
</table>
<p style="color:#999;font-size:11px">Automated load-completion SLA update generated from RAMP. RAW load dates via OptumPBMRx.etl.Tape (Load Start = FileCreateDate, Load Completion = FileLoadDate); Snap Date = the {OPTUM['snap_name']} run following each load.</p>
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
        raw_found, tag = optum_raw_status()
        # Cycle date comes from the RAW tag (fallback: the load run's StartDate).
        try:
            cycle = datetime(int(tag[4:8]), int(tag[0:2]), int(tag[2:4]))
        except (ValueError, IndexError):
            cycle = parse_dt(lrun.get('StartDate'))
        raw_loaded = {l: d['loaded'] for l, d in raw_found.items()}
        print(f"[optum_pbm] LoadQ={lqid} LoadEnd={lend} SnapEnd={send_end} tag={tag} RAW={raw_loaded}")
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
            elif (not all(raw_loaded.values()) and not force and cycle
                  and datetime.now() < cycle + timedelta(days=OPTUM_RAW_SLA_DAYS)):
                _missing = [l for l, v in raw_loaded.items() if not v]
                print(f"   RAW not all loaded yet ({', '.join(_missing)}); within "
                      f"{OPTUM_RAW_SLA_DAYS}-day window of cycle {tag} -- waiting, no report")
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
