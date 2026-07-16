"""
Every-few-hours (weekday) status digest -> Slack #data-operations-aetna-updates.

Mirrors ramp_aetnahrp_status_digest.py. Combines:
  - RAMP: every job whose name starts with 'AetnaRx Claim' (discovered from
    /api/Ramp/Job/List), name-sorted, that has run at least once.
  - SQL Agent Job Activity Monitor (msdb), mirroring SSMS:
      * TRGETL2 'ETL AetnaRx MasterLoad Claims And Eligibility'

Prints one 'SLACK|<text>' line (newlines escaped as \\n) for the poster to send.
Always emits (periodic status report), EXCEPT when the pipeline has already gone
fully green today -> the SQL Claims-and-Eligibility job Succeeded today AND every
AetnaRx Claim RAMP job that ran today ended OK (nothing Failed / still running)
-> then emits nothing, to avoid overwhelming the channel.

Note: msdb.dbo.agent_datetime is permission-blocked here, so run_date/run_time
are converted to a datetime manually.
"""
import json, os, re, subprocess, sys
from datetime import datetime, timedelta

JOB_PREFIX = 'aetnarx claim'
CHANNEL = 'data-operations-aetna-updates'

# (server, SQL Agent job name, display label)
SQL_JOBS = [
    ("TRGETL2", "ETL AetnaRx MasterLoad Claims And Eligibility",
     "ETL AetnaRx MasterLoad Claims And Eligibility"),
]

# ---- Cross-run dedupe guard (mirrors the HRP/RCE digest) ----------------------
DEDUPE_MINUTES = 25
STATE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                          'ramp_aetnarx_digest_post_state.json')

RAMP_OK = ('Successful', 'Resolved')


def _recent_emit():
    try:
        with open(STATE_FILE) as f:
            last = datetime.fromisoformat(json.load(f)['last_emit'])
    except Exception:
        return None
    return last if datetime.now() - last < timedelta(minutes=DEDUPE_MINUTES) else None


def _claim_slot():
    tmp = STATE_FILE + '.tmp'
    with open(tmp, 'w') as f:
        json.dump({'last_emit': datetime.now().isoformat()}, f)
    os.replace(tmp, STATE_FILE)


def fmt(iso):
    try:
        return datetime.fromisoformat(str(iso).split('.')[0]).strftime('%m/%d/%Y %I:%M %p')
    except Exception:
        try:
            return datetime.strptime(iso, '%Y-%m-%d %H:%M:%S').strftime('%m/%d/%Y %I:%M %p')
        except Exception:
            return iso or '?'


_JOBS_CACHE = None


def _all_jobs():
    global _JOBS_CACHE
    if _JOBS_CACHE is not None:
        return _JOBS_CACHE
    out = subprocess.run(['curl', '-s', '--ntlm', '-u', ':',
                          'http://ramp/api/Ramp/Job/List'],
                         capture_output=True, text=True, timeout=180)
    try:
        d = json.loads(out.stdout)['Data']
        _JOBS_CACHE = d[0] if (isinstance(d, list) and d and isinstance(d[0], list)) else d
    except Exception:
        _JOBS_CACHE = []
    return _JOBS_CACHE


def claim_jobs():
    """[(JobName, LatestJobRun), ...] name-sorted, for jobs 'AetnaRx Claim*'
    that have run at least once (have a StartDate). RTA jobs are excluded per
    user (2026-07-16)."""
    out = []
    for j in _all_jobs():
        name = (j.get('JobName') or j.get('Name') or '')
        if name.lower().startswith(JOB_PREFIX) and 'rta' not in name.lower():
            lr = j.get('LatestJobRun') or {}
            if lr.get('StartDate'):
                out.append((name, lr))
    out.sort(key=lambda x: x[0].lower())
    return out


def short_name(name):
    """Drop the repetitive 'AetnaRx Claim ' prefix for a cleaner label."""
    return re.sub(r'(?i)^aetnarx\s+claim\s+', '', name).strip() or name


def phase_of(name):
    """Group a Claim job into a readable pipeline phase (Load first, then Snap)."""
    n = name.lower()
    if 'snap' in n:
        return 'Snap'
    if 'stage' in n or 'load' in n:
        return 'Load'
    return 'Other'


PHASE_ORDER = ['Load', 'Snap', 'Other']


def _to_dt(v):
    if not v:
        return None
    try:
        return datetime.fromisoformat(str(v).split('.')[0])
    except Exception:
        try:
            return datetime.strptime(str(v), '%Y-%m-%d %H:%M:%S')
        except Exception:
            return None


EXEC_ICON = ':arrows_counterclockwise:'   # in-progress marker for the main line


def ramp_line(name, lr):
    """Return (head, detail) for a RAMP job's LatestJobRun. 'head' is the
    emoji + status word for the bold main line; 'detail' is the quiet italic
    sub-line (timestamps). A job that has NOT run today is shown Idle with its
    last-run outcome (RCE SQL-monitor style) instead of a stale green (per user
    2026-07-16)."""
    status = lr.get('Status', '?')
    start = lr.get('StartDate'); end = lr.get('EndDate')
    if end and not _ran_today(lr):
        oc = 'Succeeded' if status in RAMP_OK else ('Failed' if status == 'Failed' else status)
        icon = ':x:' if status == 'Failed' else ':hourglass_flowing_sand:'
        return (f"{icon} Idle", f"last run {oc} {fmt(end)}")
    if end and status in RAMP_OK:
        return (f":white_check_mark: {status}", f"started {fmt(start)} · completed {fmt(end)}")
    if end and status == 'Failed':
        return (":x: FAILED", f"started {fmt(start)} · ended {fmt(end)} — please investigate")
    if end:
        return (status, f"started {fmt(start)} · completed {fmt(end)}")
    if not start:
        return (":hourglass_flowing_sand: Queued", "not yet started")
    return (":hourglass_flowing_sand: Running", f"started {fmt(start)} · not yet complete")


def fmt_dt(d, t):
    try:
        d, t = int(d), int(t)
        if d == 0:
            return '?'
        dt = datetime(d // 10000, (d // 100) % 100, d % 100,
                      t // 10000, (t // 100) % 100, t % 100)
        return dt.strftime('%m/%d/%Y %I:%M %p')
    except Exception:
        return '?'


EXEC_STATUS = {'1': 'Executing', '2': 'Waiting for thread', '3': 'Between retries',
               '4': 'Idle', '5': 'Suspended', '7': 'Completing'}
RUN_OUTCOME = {'0': 'Failed', '1': 'Succeeded', '3': 'Canceled', '5': 'Unknown'}


def _sp_help_job(server, name):
    """Run sp_help_job raw; return its single wide data row (>=32 fields) or None."""
    q = f"SET NOCOUNT ON; EXEC msdb.dbo.sp_help_job @job_name=N'{name}', @job_aspect=N'JOB';"
    out = subprocess.run(['sqlcmd', '-S', server, '-E', '-W', '-s', '~', '-Q', q],
                         capture_output=True, text=True, timeout=120)
    for line in out.stdout.splitlines():
        line = line.rstrip()
        if not line or line.startswith('job_id') or 'rows affected' in line:
            continue
        if set(line) <= set('-~'):
            continue
        parts = line.split('~')
        if len(parts) >= 32:
            return [p.strip() for p in parts]
    return None


def remaining_secs(server, name, cur_step):
    q = ("SET NOCOUNT ON; "
         f"DECLARE @jid uniqueidentifier=(SELECT job_id FROM msdb.dbo.sysjobs WHERE name=N'{name}'); "
         ";WITH h AS (SELECT step_id, "
         "(run_duration/10000)*3600+((run_duration/100)%100)*60+(run_duration%100) AS dur_sec, "
         "ROW_NUMBER() OVER (PARTITION BY step_id ORDER BY run_date DESC, run_time DESC) rn "
         f"FROM msdb.dbo.sysjobhistory WITH (NOLOCK) WHERE @jid=job_id AND step_id>={cur_step} "
         "AND step_id<50 AND run_status=1) "
         "SELECT ISNULL(SUM(a),0) FROM (SELECT AVG(dur_sec) a FROM h WHERE rn<=8 GROUP BY step_id) y;")
    out = subprocess.run(['sqlcmd', '-S', server, '-E', '-W', '-h', '-1', '-Q', q],
                         capture_output=True, text=True, timeout=120)
    for line in out.stdout.splitlines():
        s = line.strip()
        if s.lstrip('-').isdigit():
            return int(s)
    return None


def _clock(dt):
    return dt.strftime('%I:%M %p').lstrip('0')


def sql_job(server, name):
    """Report where the CURRENT load is, via sp_help_job (what SSMS Job Activity
    Monitor shows). Executing -> live step + ETA; else Idle + last outcome.
      [-7]=current_execution_status [-6]=current_execution_step
      [-11]=last_run_outcome [-12]=last_run_time [-13]=last_run_date
    """
    row = _sp_help_job(server, name)
    if not row:
        return ("(no data)", "")
    status, step = row[-7], row[-6]
    if status == '1':
        detail = step
        m = re.match(r'\s*(\d+)', step)
        if m:
            secs = remaining_secs(server, name, int(m.group(1)))
            if secs and secs > 0:
                detail += f" · ETA ~{_clock(datetime.now() + timedelta(seconds=secs))}"
        return (f"{EXEC_ICON} Executing", detail)
    st = EXEC_STATUS.get(status, f'State {status}')
    oc = RUN_OUTCOME.get(row[-11], row[-11])
    icon = ':white_check_mark:' if oc == 'Succeeded' else (':x:' if oc == 'Failed' else ':hourglass_flowing_sand:')
    return (f"{icon} {st}", f"last run {oc} ({fmt_dt(row[-13], row[-12])})")


def job_succeeded_today(server, name):
    """True if a SQL Agent job is Idle with last run Succeeded today."""
    row = _sp_help_job(server, name)
    if not row:
        return False
    status, outcome, lrd = row[-7], row[-11], row[-13]
    try:
        d = int(lrd)
    except (ValueError, TypeError):
        return False
    t = datetime.now()
    return status == '4' and outcome == '1' and d == t.year * 10000 + t.month * 100 + t.day


def _ran_today(lr):
    s = _to_dt(lr.get('StartDate'))
    return bool(s and s.date() == datetime.now().date())


def all_claim_green_today(jobs):
    """True if every AetnaRx Claim RAMP job that ran today ended OK (none Failed,
    none still running). Vacuously true if none ran today."""
    for name, lr in jobs:
        if _ran_today(lr):
            if lr.get('Status') not in RAMP_OK or not lr.get('EndDate'):
                return False
    return True


def main():
    force = '--force' in sys.argv
    if not force:
        recent = _recent_emit()
        if recent:
            print(f"NO_POST: deduped (a digest was already emitted at "
                  f"{recent.strftime('%I:%M %p')}, within {DEDUPE_MINUTES} min)")
            return
    _claim_slot()

    jobs = claim_jobs()

    # Fully green for today (SQL Claims+Elig Succeeded today AND no Claim RAMP job
    # failed/running today) -> redundant, emit nothing.
    if (job_succeeded_today('TRGETL2', 'ETL AetnaRx MasterLoad Claims And Eligibility')
            and all_claim_green_today(jobs)):
        print('NO_POST: AetnaRx Claim pipeline + SQL Succeeded today (RAMP + SQL)')
        return

    now = datetime.now().strftime('%m/%d/%Y %I:%M %p')
    # Bold job/section names carry the status (main line); the timestamps drop to a
    # quiet italic sub-line so the reader hones in on state (per user 2026-07-16).
    lines = [f":bar_chart: *AetnaRx Claim - Status Update*   _{now}_", ""]
    if jobs:
        # Break the pipeline into readable phase sub-sections (like HRP): each
        # phase is its own labeled block separated by a blank line.
        grouped = {}
        for name, lr in jobs:
            grouped.setdefault(phase_of(name), []).append((name, lr))
        for phase in PHASE_ORDER:
            if phase not in grouped:
                continue
            lines.append(f"*RAMP — {phase}*")
            for name, lr in grouped[phase]:
                head, detail = ramp_line(name, lr)
                lines.append(f"*{short_name(name)}*  {head}")
                if detail:
                    lines.append(f"_{detail}_")
                lines.append("")   # blank line between jobs for readability
    else:
        lines.append("*RAMP*")
        lines.append("(no AetnaRx Claim jobs found in RAMP)")
        lines.append("")
    lines.append("*SQL Job Activity Monitor*")
    for server, name, label in SQL_JOBS:
        head, detail = sql_job(server, name)
        lines.append(f"*{label}*  ({server})  {head}")
        if detail:
            lines.append(f"_{detail}_")
        lines.append("")
    while lines and lines[-1] == "":
        lines.pop()
    msg = "\n".join(lines)
    print("SLACK|" + msg.replace("\n", "\\n"))


if __name__ == '__main__':
    main()
