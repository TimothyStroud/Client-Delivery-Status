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
        return (f":white_check_mark: {status}", f"started {fmt(start)} | completed {fmt(end)}")
    if end and status == 'Failed':
        return (":x: FAILED", f"started {fmt(start)} | ended {fmt(end)} - please investigate")
    if end:
        return (status, f"started {fmt(start)} | completed {fmt(end)}")
    if not start:
        return (":hourglass_flowing_sand: Queued", "not yet started")
    return (":hourglass_flowing_sand: Running", f"started {fmt(start)} | not yet complete")


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


def last_completion(server, name):
    """Datetime the job's most recent run FINISHED = start + duration from the
    step_id=0 (job outcome) row in sysjobhistory. None if unavailable."""
    q = ("SET NOCOUNT ON; "
         f"DECLARE @jid uniqueidentifier=(SELECT job_id FROM msdb.dbo.sysjobs WHERE name=N'{name}'); "
         "SELECT TOP 1 run_date, run_time, run_duration FROM msdb.dbo.sysjobhistory WITH (NOLOCK) "
         "WHERE job_id=@jid AND step_id=0 ORDER BY run_date DESC, run_time DESC;")
    out = subprocess.run(['sqlcmd', '-S', server, '-E', '-W', '-h', '-1', '-s', '|', '-Q', q],
                         capture_output=True, text=True, timeout=120)
    for line in out.stdout.splitlines():
        parts = [p.strip() for p in line.split('|')]
        if len(parts) >= 3 and parts[0].isdigit():
            try:
                d, t, dur = int(parts[0]), int(parts[1]), int(parts[2])
                start = datetime(d // 10000, (d // 100) % 100, d % 100,
                                 t // 10000, (t // 100) % 100, t % 100)
                dur_s = (dur // 10000) * 3600 + ((dur // 100) % 100) * 60 + (dur % 100)
                return start + timedelta(seconds=dur_s)
            except Exception:
                return None
    return None


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
    if status == '1':                        # Executing -> "Executing Step N (name)" + ETA line
        detail = []
        m = re.match(r'\s*(\d+)', step)
        if m:
            secs = remaining_secs(server, name, int(m.group(1)))
            if secs and secs > 0:
                eta = _clock(datetime.now() + timedelta(seconds=secs))
                # ETA on its own line with a red-dot emoji (only standout Slack
                # renders; mrkdwn/color do not) -- per user 2026-07-16.
                detail.append(f":red_circle: ETA ~{eta}")
        return (f"Executing Step {step}", detail)
    # Idle: if it completed SUCCESSFULLY TODAY, show green circle + Successful +
    # completion time (in place of the ETA line), per user 2026-07-16.
    oc = RUN_OUTCOME.get(row[-11], row[-11])
    try:
        _d = int(row[-13]); _t = datetime.now()
        ran_today = _d == _t.year * 10000 + _t.month * 100 + _t.day
    except (ValueError, TypeError):
        ran_today = False
    if oc == 'Succeeded' and ran_today:
        comp = last_completion(server, name)
        ctext = comp.strftime('%m/%d/%Y %I:%M %p') if comp else fmt_dt(row[-13], row[-12])
        return ("", [f":large_green_circle: Successful {ctext}"])
    st = EXEC_STATUS.get(status, f'State {status}')
    return (f"- {st}", [])


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

    # Once ETL AetnaRx MasterLoad Claims And Eligibility has SUCCEEDED today, the
    # rest of the day's digests are redundant -> emit no SLACK line.
    if not force and job_succeeded_today('TRGETL2', 'ETL AetnaRx MasterLoad Claims And Eligibility'):
        print('NO_POST: ETL AetnaRx MasterLoad Claims And Eligibility Succeeded today')
        return

    now = datetime.now().strftime('%m/%d/%Y %I:%M %p')
    # Minimal PLAIN-TEXT format (per user 2026-07-16): ONLY the ETL AetnaRx
    # MasterLoad Claims And Eligibility step & ETA. Webhook renders only :emoji:
    # (no markup/color), so the only standout is the :red_circle: on the ETA line.
    lines = [f"Aetna Rx - Status Update   ({now})", ""]
    for server, name, label in SQL_JOBS:
        status_text, detail = sql_job(server, name)
        lines.append(f"{label} {status_text}".rstrip())
        lines.extend(detail)
        lines.append("")
    while lines and lines[-1] == "":
        lines.pop()
    msg = "\n".join(lines)
    print("SLACK|" + msg.replace("\n", "\\n"))


if __name__ == '__main__':
    main()
