"""
Every-few-hours (weekday) status digest -> Slack #team-rdp-operations-support ONLY.

Combines:
  - RAMP 'Aetna 0110 HRP Load' (JobId 1246): status + completion time.
  - SQL Agent Job Activity Monitor (msdb), mirroring SSMS:
      * TRGETL2 'ETL AetnaHRP MasterLoad'

Prints one 'SLACK|<text>' line (newlines escaped as \\n) for the poster to send.
Always emits (it's a periodic status report, not event-driven), except when the
HRP load has already succeeded today (both RAMP + SQL job) -> then emits nothing.

Note: msdb.dbo.agent_datetime is permission-blocked here, so run_date/run_time
are converted to a datetime manually.
"""
import json, os, re, subprocess, sys
from datetime import datetime, timedelta

HRP_JOBID = 1246
SNAP_JOBID = 1247      # RAMP 'Aetna 0120 HRP Snap' (added 2026-07-14 per user)
STAGE_JOBID = 1243     # RAMP 'Aetna 0100 HRP Stage'
CHANNEL = 'C09EPLQL2D9'

# Claim files are sourced from RAMP's [ramp].[FileLog] on TRGUTIL10 keyed by the
# last Stage's QueueId (reworked 2026-07-14 per user: show the whole batch the
# LAST 'Aetna 0100 HRP Stage' staged, not whatever happens to be on the file
# share). This drops stale stragglers (e.g. an old file still sitting in Loaded)
# and matches exactly what the stage picked up.
RAMP_SQL_SERVER = 'TRGUTIL10'

# ---- Cross-run dedupe guard (mirrors the RCE digest) --------------------------
# A near-simultaneous second run (task jitter) within DEDUPE_MINUTES prints a
# 'NO_POST: deduped ...' line and emits nothing. The slot is CLAIMED (file
# written) before the slow SQL/curl work so a second run bails almost instantly.
# Window (25 min) > max jitter, < real slot spacing (~2 h). --force bypasses.
DEDUPE_MINUTES = 25
STATE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                          'ramp_aetnahrp_digest_post_state.json')


def _recent_emit():
    """Return the last-emit datetime if within DEDUPE_MINUTES, else None."""
    try:
        with open(STATE_FILE) as f:
            last = datetime.fromisoformat(json.load(f)['last_emit'])
    except Exception:
        return None
    return last if datetime.now() - last < timedelta(minutes=DEDUPE_MINUTES) else None


def _claim_slot():
    """Stamp now as the last-emit time (atomic replace), claiming this slot."""
    tmp = STATE_FILE + '.tmp'
    with open(tmp, 'w') as f:
        json.dump({'last_emit': datetime.now().isoformat()}, f)
    os.replace(tmp, STATE_FILE)

# (server, SQL Agent job name, display label). AetnaHRP's ETL Load is the SQL
# Agent job 'ETL AetnaHRP MasterLoad' on TRGETL2.
SQL_JOBS = [
    ("TRGETL2", "ETL AetnaHRP MasterLoad", "ETL AetnaHRP Masterload"),
]

OUTCOME = {1: "Succeeded", 0: "Failed", 2: "Retry", 3: "Canceled", 4: "In Progress"}


def fmt(iso):
    try:
        return datetime.fromisoformat(iso).strftime('%m/%d/%Y %I:%M %p')
    except Exception:
        try:
            return datetime.strptime(iso, '%Y-%m-%d %H:%M:%S').strftime('%m/%d/%Y %I:%M %p')
        except Exception:
            return iso or '?'


_JOBS_CACHE = None


def _all_jobs():
    """Fetch RAMP /Job/List once per process and cache it."""
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


def job_run(jobid):
    """LatestJobRun dict for a RAMP job id (or {} if not found)."""
    for j in _all_jobs():
        if j.get('JobId') == jobid:
            return j.get('LatestJobRun') or {}
    return {}


def hrp_run():
    """LatestJobRun for RAMP 'Aetna 0110 HRP Load' (1246)."""
    return job_run(HRP_JOBID)


# RAMP terminal-good statuses: load/mining jobs report 'Successful'; snap jobs
# report 'Resolved' on a clean run. Both get a green check.
RAMP_OK = ('Successful', 'Resolved')


EXEC_ICON = ':arrows_counterclockwise:'   # in-progress marker for the main line


def ramp_line(jobid):
    """Return (head, detail) for a RAMP job's LatestJobRun. 'head' = emoji +
    status word for the bold main line; 'detail' = the quiet italic sub-line.
    A job that has NOT run today is shown Idle with its last-run outcome (per user
    2026-07-16)."""
    lr = job_run(jobid)
    status = lr.get('Status', '?')
    start = lr.get('StartDate'); end = lr.get('EndDate')
    if end and not _started_today(start):
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


def _started_today(start):
    dt = _to_dt(start)
    return bool(dt and dt.date() == datetime.now().date())


def fmt_dt(d, t):
    """Build mm/dd/yyyy h:MM AM/PM from SQL Agent run_date (yyyymmdd) + run_time (hhmmss) ints."""
    try:
        d, t = int(d), int(t)
        if d == 0:
            return '?'
        dt = datetime(d // 10000, (d // 100) % 100, d % 100,
                      t // 10000, (t // 100) % 100, t % 100)
        return dt.strftime('%m/%d/%Y %I:%M %p')
    except Exception:
        return '?'


# SQL Agent current_execution_status / last_run_outcome code maps.
EXEC_STATUS = {'1': 'Executing', '2': 'Waiting for thread', '3': 'Between retries',
               '4': 'Idle', '5': 'Suspended', '7': 'Completing'}
RUN_OUTCOME = {'0': 'Failed', '1': 'Succeeded', '3': 'Canceled', '5': 'Unknown'}


def _sp_help_job(server, name):
    """Run sp_help_job raw and return its single wide data row as a list of
    stripped fields (>=32), or None. Run raw (NOT via INSERT EXEC) so a
    non-sysadmin can read the live current step via ownership chaining; parsed
    from the END so leading text columns can't shift the fields we need."""
    q = f"SET NOCOUNT ON; EXEC msdb.dbo.sp_help_job @job_name=N'{name}', @job_aspect=N'JOB';"
    out = subprocess.run(['sqlcmd', '-S', server, '-E', '-W', '-s', '~', '-Q', q],
                         capture_output=True, text=True, timeout=120)
    for line in out.stdout.splitlines():
        line = line.rstrip()
        if not line or line.startswith('job_id') or 'rows affected' in line:
            continue
        if set(line) <= set('-~'):          # the ---- separator row
            continue
        parts = line.split('~')
        if len(parts) >= 32:
            return [p.strip() for p in parts]
    return None


def remaining_secs(server, name, cur_step):
    """Estimated seconds left = sum of avg historical durations of steps
    cur_step..end (avg over the last 8 successful runs per step)."""
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
    """Report where the CURRENT load is, via sp_help_job (the value SSMS Job
    Activity Monitor shows). When executing, returns the live step + ETA. When
    not running, returns Idle + the most recent load's outcome.

      [-7]=current_execution_status [-6]=current_execution_step
      [-11]=last_run_outcome [-12]=last_run_time [-13]=last_run_date
    """
    row = _sp_help_job(server, name)
    if not row:
        return ("(no data)", "")

    status, step = row[-7], row[-6]
    if status == '1':                        # Executing -> current step + ETA
        detail = step
        m = re.match(r'\s*(\d+)', step)      # leading step number
        if m:
            secs = remaining_secs(server, name, int(m.group(1)))
            if secs and secs > 0:
                detail += f" | ETA ~{_clock(datetime.now() + timedelta(seconds=secs))}"
        return (f"{EXEC_ICON} Executing", detail)
    st = EXEC_STATUS.get(status, f'State {status}')
    oc = RUN_OUTCOME.get(row[-11], row[-11])
    # Checkmark for a Succeeded last run, red X for Failed (per user 2026-07-14).
    icon = ':white_check_mark:' if oc == 'Succeeded' else (':x:' if oc == 'Failed' else '')
    head = f"{icon} {st}".strip()
    return (head, f"last run {oc} ({fmt_dt(row[-13], row[-12])})")


def _to_dt(v):
    """Parse a RAMP ISO-ish timestamp to datetime, or None."""
    if not v:
        return None
    try:
        return datetime.fromisoformat(str(v).split('.')[0])
    except Exception:
        try:
            return datetime.strptime(str(v), '%Y-%m-%d %H:%M:%S')
        except Exception:
            return None


def snap_is_current(load_jobid, snap_jobid):
    """True if the snap's latest run belongs to the CURRENT load — i.e. the snap
    STARTED at/after the load's latest completion. A load that hasn't completed
    yet has no current snap (returns False). Per user 2026-07-14: don't credit a
    snap that ran for a prior load cycle."""
    load_end = _to_dt(job_run(load_jobid).get('EndDate'))
    snap_start = _to_dt(job_run(snap_jobid).get('StartDate'))
    return bool(load_end and snap_start and snap_start >= load_end)


def snap_line(load_jobid, snap_jobid):
    """Status body for a Snap RAMP job. Per user (2026-07-14): do NOT show the
    snap as Successful until it has finished for the CURRENT load. A resolved snap
    whose run predates the current load's completion (or a load still running) is
    stale -> shown as waiting, not a green check. A Failed snap still surfaces as
    a red X regardless (worth investigating)."""
    lr = job_run(snap_jobid)
    status = lr.get('Status', '?')
    start = lr.get('StartDate'); end = lr.get('EndDate')
    if end and status in RAMP_OK and not snap_is_current(load_jobid, snap_jobid):
        return (":hourglass_flowing_sand: Waiting to snap current load",
                f"last snap {status} {fmt(end)}, ran before this load completed")
    if end and status in RAMP_OK:
        return (f":white_check_mark: {status}", f"started {fmt(start)} | completed {fmt(end)}")
    if end and status == 'Failed':
        return (":x: FAILED", f"started {fmt(start)} | ended {fmt(end)} - please investigate")
    if end:
        return (status, f"started {fmt(start)} | completed {fmt(end)}")
    if not start:
        return (":hourglass_flowing_sand: Queued", "not yet started")
    return (":hourglass_flowing_sand: Running", f"started {fmt(start)} | not yet complete")


def _ramp_succeeded_today(jobid, ok=('Successful',)):
    """True if a RAMP job reached an OK terminal status today."""
    lr = job_run(jobid)
    end = lr.get('EndDate')
    if lr.get('Status') in ok and end:
        try:
            return datetime.fromisoformat(str(end).split('.')[0]).date() == datetime.now().date()
        except Exception:
            return False
    return False


def hrp_succeeded_today():
    """True if RAMP 'Aetna 0110 HRP Load' (1246) completed Successful today."""
    return _ramp_succeeded_today(HRP_JOBID)


def snap_succeeded_today():
    """True if RAMP 'Aetna 0120 HRP Snap' (1247) reached an OK status today AND
    that snap ran for the CURRENT load (started after the load completed). A stale
    resolved snap from a prior cycle does NOT count (per user 2026-07-14), so the
    digest keeps posting until the snap actually runs for today's load."""
    return (_ramp_succeeded_today(SNAP_JOBID, RAMP_OK)
            and snap_is_current(HRP_JOBID, SNAP_JOBID))


def _parse_extract_dt(name):
    """Parse the embedded YYMMDDHHMMSS from VENDOR.CB-CLAIMS-EXTRACT.<ts>.csv."""
    m = re.search(r'\.(\d{12})\.csv$', name)
    if not m:
        return None
    try:
        return datetime.strptime(m.group(1), '%y%m%d%H%M%S')
    except ValueError:
        return None


def _ramp_sql(query):
    """Run a query against the RAMP db on TRGUTIL10; return rows as lists of
    stripped string fields (headers suppressed with -h -1). Returns [] on error."""
    try:
        out = subprocess.run(
            ['sqlcmd', '-S', RAMP_SQL_SERVER, '-d', 'RAMP', '-E', '-W',
             '-h', '-1', '-s', '|', '-Q', 'SET NOCOUNT ON; ' + query],
            capture_output=True, text=True, timeout=120)
    except Exception:
        return []
    rows = []
    for line in out.stdout.splitlines():
        line = line.rstrip()
        if not line or set(line) <= set('-|') or 'rows affected' in line:
            continue
        rows.append([c.strip() for c in line.split('|')])
    return rows


def last_stage_batch():
    """The claim files the LAST 'Aetna 0100 HRP Stage' staged, from FileLog.
    Returns (stage_qid, stage_end_datetime, [(filename, data_dt), ...] oldest-first),
    or (None, None, []) if unavailable. The stage QueueId is the newest one whose
    FileLog actually holds CB-CLAIMS-EXTRACT files (a stage job also logs a
    fileless 'Resolved' phase we must skip)."""
    qrows = _ramp_sql(
        "SELECT TOP 1 fl.QueueId FROM [ramp].[FileLog] fl "
        "JOIN [ramp].[Queue] q ON q.QueueId = fl.QueueId "
        f"WHERE q.JobId = {STAGE_JOBID} "
        "AND fl.FileName LIKE 'VENDOR.CB-CLAIMS-EXTRACT%' ORDER BY fl.QueueId DESC")
    if not qrows:
        return None, None, []
    qid = qrows[0][0]
    erows = _ramp_sql(
        f"SELECT CONVERT(varchar(19), EndDate, 121) FROM [ramp].[Queue] WHERE QueueId = {qid}")
    stage_end = _to_dt(erows[0][0]) if erows and erows[0] else None
    frows = _ramp_sql(
        "SELECT FileName FROM [ramp].[FileLog] "
        f"WHERE QueueId = {qid} AND FileName LIKE 'VENDOR.CB-CLAIMS-EXTRACT%' ORDER BY FileName")
    files = [(r[0], _parse_extract_dt(r[0])) for r in frows if r and r[0]]
    files.sort(key=lambda x: (x[1] or datetime.min))
    return qid, stage_end, files


def batch_state(stage_end):
    """How far the last Stage's batch has progressed through Load -> Snap, as
    (icon, label). Per user 2026-07-14 a Snap only counts once it runs for the
    CURRENT load, so 'snapped' requires snap_is_current."""
    load = job_run(HRP_JOBID)
    l_start = _to_dt(load.get('StartDate')); l_end = _to_dt(load.get('EndDate'))
    # Is the latest Load run the one for this stage batch (started after staging)?
    load_for_batch = bool(l_start and stage_end and l_start >= stage_end)
    if load_for_batch and load.get('Status') == 'Failed' and l_end:
        return ':x:', 'load FAILED'
    if load_for_batch and l_end and load.get('Status') == 'Successful':
        if snap_is_current(HRP_JOBID, SNAP_JOBID) \
                and job_run(SNAP_JOBID).get('Status') in RAMP_OK:
            return ':white_check_mark:', 'loaded + snapped'
        return ':white_check_mark:', 'loaded (snap pending)'
    if load_for_batch and not l_end:
        return ':hourglass_flowing_sand:', 'loading'
    return ':hourglass_flowing_sand:', 'staged, pending load'


def claim_file_lines(files, icon):
    """Slack lines for the claim-files section: the whole batch from the last
    Stage, each file tagged with its data date; the batch's Load/Snap progress is
    carried in the section header (see main)."""
    if not files:
        return ["- (RAMP FileLog unavailable / no CB-CLAIMS-EXTRACT files in last stage)"]
    out = []
    for name, dt in files:
        dstr = dt.strftime('%m/%d/%Y') if dt else '?'
        out.append(f"- {icon} `{name}` ({dstr})")
    return out


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


def main():
    # Cross-run dedupe: bail if another run already emitted within DEDUPE_MINUTES.
    force = '--force' in sys.argv
    if not force:
        recent = _recent_emit()
        if recent:
            print(f"NO_POST: deduped (a digest was already emitted at "
                  f"{recent.strftime('%I:%M %p')}, within {DEDUPE_MINUTES} min)")
            return
    _claim_slot()

    # Once the HRP load has already SUCCEEDED today in BOTH RAMP and the SQL
    # Agent job, the rest of the day's digests are redundant -> emit no SLACK
    # line so nothing posts. (If either is still running, failed, or hasn't run,
    # the digest still posts.)
    if (hrp_succeeded_today()
            and snap_succeeded_today()
            and job_succeeded_today('TRGETL2', 'ETL AetnaHRP MasterLoad')):
        print('NO_POST: HRP load + snap Succeeded today (RAMP + SQL)')
        return
    now = datetime.now().strftime('%m/%d/%Y %I:%M %p')
    # Bold job/section names carry the status (main line); timestamps drop to a
    # quiet italic sub-line so the reader hones in on state (per user 2026-07-16).
    lines = [f":bar_chart: *Aetna HRP - Status Update*   _{now}_", ""]
    lines.append("*RAMP*")
    for label, (head, detail) in [
            ("Aetna 0110 HRP Load", ramp_line(HRP_JOBID)),
            ("Aetna 0120 HRP Snap", snap_line(HRP_JOBID, SNAP_JOBID))]:
        lines.append(f"*{label}*  {head}")
        if detail:
            lines.append(f"_{detail}_")
        lines.append("")
    lines.append("*SQL Job Activity Monitor*")
    for server, name, label in SQL_JOBS:
        head, detail = sql_job(server, name)
        lines.append(f"*{label}*  ({server})  {head}")
        if detail:
            lines.append(f"_{detail}_")
        lines.append("")
    _stage_qid, stage_end, files = last_stage_batch()
    icon, state_label = batch_state(stage_end)
    staged_on = stage_end.strftime('%m/%d/%Y') if stage_end else '?'
    lines.append(f"*Claim Files - last Aetna 0100 HRP Stage*   _staged {staged_on}, {state_label}_")
    lines.extend(claim_file_lines(files, icon))
    msg = "\n".join(lines)
    print("SLACK|" + msg.replace("\n", "\\n"))


if __name__ == '__main__':
    main()
