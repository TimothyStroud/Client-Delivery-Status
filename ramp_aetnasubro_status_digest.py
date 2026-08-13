"""
Every-2-hours status digest for Aetna Subro -> Slack #data-operations-aetna-updates.

Combines:
  - SQL Agent job (msdb, mirroring SSMS Job Activity Monitor):
      * TRGETL2 'DTS AetnaSubro MasterLoad'  (the load itself + ETA)
  - RAMP 'Aetna 0110 Subro Load' (JobId 2242): status + completion time.
  - RAMP 'Aetna 0120 Subro Start Snap' (JobId 10121): snap for the CURRENT load.
  - The file batch the last 'Aetna 0100 Subro Stage' (JobId 2243) staged.

Prints one 'SLACK|<text>' line (newlines escaped as \\n) for the poster to send.
Content-deduped: emits only when the status text CHANGES since the last post, so
on the ~29 days a month with no Subro load it posts nothing at all.

Deliberate differences from the HRP/Rx digests (which this is otherwise a clone
of -- they are standalone by convention, no shared module):

  * No SSISDB milestone ladder. Those jobs run an SSIS package through the SSISDB
    catalog, which logs every task completion and gives a mid-run progress signal.
    'DTS AetnaSubro MasterLoad' is a SINGLE CmdExec step running a legacy DTS
    package, so there is nothing to read: whole-job survival history anchored to
    the live run's start is the only honest estimate available.
  * Idle shows the last run's outcome + date ALWAYS, not only while the completion
    falls on today. Subro is MONTHLY (runs ~the 11th-16th), so the HRP rule would
    leave a bare '- Idle' for four weeks out of five. The text stays constant
    between runs, so content dedupe still keeps it to a single post.

Note: msdb.dbo.agent_datetime is permission-blocked here, so run_date/run_time
are converted to a datetime manually.
"""
import json, os, re, subprocess, sys
from datetime import datetime, timedelta

SUBRO_JOBID = 2242     # RAMP 'Aetna 0110 Subro Load'
SNAP_JOBID = 10121     # RAMP 'Aetna 0120 Subro Start Snap'
STAGE_JOBID = 2243     # RAMP 'Aetna 0100 Subro Stage'

# The staged file batch is read from RAMP's [ramp].[FileLog] on TRGUTIL10, keyed by
# the last Stage's QueueId (same approach as the HRP digest): that shows exactly
# what the stage picked up rather than whatever happens to be on the file share.
RAMP_SQL_SERVER = 'TRGUTIL10'

# ---- Cross-run dedupe guard (mirrors the HRP/RCE digests) ---------------------
# A near-simultaneous second run (task jitter) within DEDUPE_MINUTES prints a
# 'NO_POST: deduped ...' line and emits nothing. The slot is CLAIMED (file
# written) before the slow SQL/curl work so a second run bails almost instantly.
# Window (25 min) > max jitter, < real slot spacing (~2 h). --force bypasses.
DEDUPE_MINUTES = 25
STATE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                          'ramp_aetnasubro_digest_post_state.json')


def _load_state():
    try:
        with open(STATE_FILE) as f:
            return json.load(f)
    except Exception:
        return {}


def _recent_emit():
    """Return the last-emit datetime if within DEDUPE_MINUTES, else None."""
    try:
        last = datetime.fromisoformat(_load_state()['last_emit'])
    except Exception:
        return None
    return last if datetime.now() - last < timedelta(minutes=DEDUPE_MINUTES) else None


def _last_msg():
    """Text of the digest we most recently POSTED (for content dedupe), or None."""
    return _load_state().get('last_msg')


def _claim_slot(msg=None):
    """Stamp now as the last-emit time (atomic replace), claiming this slot. If
    msg is given, also record it as the last-posted message so an identical later
    digest is skipped (content dedupe)."""
    st = _load_state()
    st['last_emit'] = datetime.now().isoformat()
    if msg is not None:
        st['last_msg'] = msg
    tmp = STATE_FILE + '.tmp'
    with open(tmp, 'w') as f:
        json.dump(st, f)
    os.replace(tmp, STATE_FILE)


# (server, SQL Agent job name, display label).
SQL_JOBS = [
    ("TRGETL2", "DTS AetnaSubro MasterLoad", "DTS AetnaSubro MasterLoad"),
]


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
    out = subprocess.run(['curl', '-s', '--negotiate', '-u', ':',
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


# RAMP terminal-good statuses: load jobs report 'Successful'; snap jobs report
# 'Resolved' on a clean run (and a load shows 'Resolved' when an operator resolved
# it by hand). Both get a green check.
RAMP_OK = ('Successful', 'Resolved')

EXEC_ICON = ':arrows_counterclockwise:'   # in-progress marker for the ETA line

# Per user 2026-08-13: a RAMP job actively Running gets :loading: (a workspace
# custom emoji), not the hourglass. Queued/Idle/Waiting keep :hourglass_flowing_sand:
# -- the distinction is "moving right now" vs "waiting its turn".
RUNNING_ICON = ':loading:'


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


def _started_today(start):
    dt = _to_dt(start)
    return bool(dt and dt.date() == datetime.now().date())


def ramp_line(jobid):
    """Return (head, detail) for a RAMP job's LatestJobRun. 'head' = emoji +
    status word for the main line; 'detail' = the quiet sub-line. A job that has
    NOT run today is shown Idle with its last-run outcome."""
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
    return (f"{RUNNING_ICON} Running", f"started {fmt(start)} | not yet complete")


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


# ---- ETA (anchored to the actual run start) ----------------------------------
# Anchoring to the live run's START (not to 'now') is what keeps the ETA fixed
# across the 2-hourly ticks; the pre-2026-07-21 'now + remaining' shape in the
# other digests marched forward ~2 h every 2 h and never converged.
#
# Subro's runtimes are wide (the last 8 successful runs span ~6 h to ~28 h) and
# there is no mid-run progress signal at all -- a single CmdExec step running a
# DTS package logs nothing until it exits. So the one honest refinement each tick
# is elapsed time itself: drop the historical outcomes already ruled out by how
# long this run has been going, then take the median of what remains.


def _recent_full_durations(server, name, n=8):
    """Durations (seconds) of the last n SUCCESSFUL full runs (step_id=0)."""
    q = ("SET NOCOUNT ON; "
         f"DECLARE @jid uniqueidentifier=(SELECT job_id FROM msdb.dbo.sysjobs WHERE name=N'{name}'); "
         f"SELECT TOP {n} run_duration FROM msdb.dbo.sysjobhistory WITH (NOLOCK) "
         "WHERE job_id=@jid AND step_id=0 AND run_status=1 ORDER BY run_date DESC, run_time DESC;")
    out = subprocess.run(['sqlcmd', '-S', server, '-E', '-W', '-h', '-1', '-Q', q],
                         capture_output=True, text=True, timeout=120)
    durs = []
    for line in out.stdout.splitlines():
        s = line.strip()
        if s.isdigit():
            v = int(s)
            durs.append((v // 10000) * 3600 + ((v // 100) % 100) * 60 + (v % 100))
    return durs


# sysjobactivity retains orphaned rows with a NULL stop_execution_date after an
# Agent restart, so a recency guard is required or an idle job reports a weeks-old
# run as live. The window must exceed the longest real run: Subro's slowest recent
# run was 28h04m, so 3 days leaves headroom without re-admitting stale rows.
STALE_ACTIVITY_DAYS = 3


def _current_run_start(server, name):
    """Start datetime of the CURRENTLY-executing run from sysjobactivity (the row
    on the latest Agent session with no stop time), or None if not executing /
    unreadable."""
    # NB: don't join msdb.dbo.syssessions to find the latest Agent session -- SELECT
    # on it is permission-denied here. Ordering sysjobactivity by session_id DESC
    # (then newest start) picks the current run's row directly.
    q = ("SET NOCOUNT ON; "
         f"DECLARE @jid uniqueidentifier=(SELECT job_id FROM msdb.dbo.sysjobs WHERE name=N'{name}'); "
         "SELECT TOP 1 CONVERT(varchar(19), ja.start_execution_date, 120) "
         "FROM msdb.dbo.sysjobactivity ja WITH (NOLOCK) "
         "WHERE ja.job_id=@jid AND ja.start_execution_date IS NOT NULL "
         "AND ja.stop_execution_date IS NULL "
         f"AND ja.start_execution_date > DATEADD(day,-{STALE_ACTIVITY_DAYS},GETDATE()) "
         "ORDER BY ja.session_id DESC, ja.start_execution_date DESC;")
    out = subprocess.run(['sqlcmd', '-S', server, '-E', '-W', '-h', '-1', '-Q', q],
                         capture_output=True, text=True, timeout=120)
    for line in out.stdout.splitlines():
        dt = _to_dt(line.strip())
        if dt:
            return dt
    return None


def _pct(sorted_vals, p):
    """Nearest-rank percentile of an ascending list (None if empty)."""
    if not sorted_vals:
        return None
    import math
    k = max(1, math.ceil(p / 100.0 * len(sorted_vals)))
    return sorted_vals[k - 1]


def _dur_h(sec):
    """Compact elapsed: 'Xh YYm' / 'Xh' / 'Ym'."""
    m = int(round(sec / 60.0))
    h, m = divmod(m, 60)
    if h and m:
        return f"{h}h {m:02d}m"
    if h:
        return f"{h}h"
    return f"{m}m"


def _eta_stamp(dt):
    """Clock time; prefixed with mm/dd when the ETA is not today."""
    if dt.date() == datetime.now().date():
        return dt.strftime('%I:%M %p').lstrip('0')
    return dt.strftime('%m/%d %I:%M %p').lstrip('0')


NOTE_ICON = ':information_source:'   # neutral on purpose -- a long run is NOT a
                                     # failure, so this never uses the red marker.


def _why_longer(durs, beyond_history, projected=False):
    """One-line, HISTORY-DERIVED explanation of why the ETA moved out. Every number
    comes from this job's OWN recent successful runs -- fastest, slowest, typical --
    so nothing is hardcoded and it can't go stale as the load's profile changes.
    Kept identical across the Aetna digests (standalone clones, no shared module)."""
    if not durs:
        return []
    lo, hi, typical, n = durs[0], durs[-1], _pct(durs, 50), len(durs)
    if beyond_history:
        tail = ("so the ETA is projected from elapsed time rather than history"
                if projected else "so recent history no longer bounds the finish")
        return [f"{NOTE_ICON} why: already past the slowest of the last {n} runs "
                f"({_dur_h(hi)}), {tail}"]
    return [f"{NOTE_ICON} why: this load's run time varies (last {n}: {_dur_h(lo)}-{_dur_h(hi)}, "
            f"typically {_dur_h(typical)}); we are past the typical, so the ETA now reflects "
            f"only the slower run times still possible"]


_STEPS_CACHE = {}


def _steps(server, name):
    """{step_id: (step_name, subsystem)} for the job, cached per process."""
    key = (server, name)
    if key in _STEPS_CACHE:
        return _STEPS_CACHE[key]
    q = ("SET NOCOUNT ON; "
         f"DECLARE @jid uniqueidentifier=(SELECT job_id FROM msdb.dbo.sysjobs WHERE name=N'{name}'); "
         "SELECT step_id, subsystem, step_name FROM msdb.dbo.sysjobsteps "
         "WHERE job_id=@jid ORDER BY step_id;")
    steps = {}
    try:
        out = subprocess.run(['sqlcmd', '-S', server, '-E', '-W', '-h', '-1', '-s', '|', '-Q', q],
                             capture_output=True, text=True, timeout=120)
        for line in out.stdout.splitlines():
            parts = [p.strip() for p in line.split('|')]
            if len(parts) >= 3 and parts[0].isdigit():
                steps[int(parts[0])] = (parts[2], parts[1])
    except Exception:
        pass
    _STEPS_CACHE[key] = steps
    return steps


def eta_detail(server, name):
    """Single expected-completion ETA line for the in-flight run, matching the
    other Aetna digests' 'ETA ~<time>' look.

    Conditional (survival) estimate: the live run's actual START + the median of
    the recent successful durations that are still POSSIBLE given how long it has
    already been running. As elapsed grows the estimate climbs and converges on the
    real finish, and it moves ONLY when a shorter outcome is genuinely eliminated.
    Once every historical outcome is ruled out we stop quoting a clock time and just
    say it is running long. Degrades to a 'now + median' estimate if the live run
    start can't be read from sysjobactivity."""
    durs = sorted(_recent_full_durations(server, name))
    if not durs:
        return [f"{EXEC_ICON} in progress"]
    start = _current_run_start(server, name)
    if not start:
        eta = datetime.now() + timedelta(seconds=_pct(durs, 50))
        return [f"{EXEC_ICON} ETA ~{_eta_stamp(eta)}"]
    elapsed = (datetime.now() - start).total_seconds()
    still_possible = [d for d in durs if d >= elapsed]
    if not still_possible:
        return ([f"{EXEC_ICON} running {_dur_h(elapsed)} - longer than usual, still processing"]
                + _why_longer(durs, True))
    eta = start + timedelta(seconds=_pct(still_possible, 50))
    line = [f"{EXEC_ICON} ETA ~{_eta_stamp(eta)}"]
    # Note only once the run is genuinely slower than usual (past p75), so normal
    # runs stay a clean one-liner.
    if elapsed > _pct(durs, 75):
        return line + _why_longer(durs, False)
    return line


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
    """Report where the CURRENT load is, via sp_help_job (the value SSMS Job
    Activity Monitor shows). When executing, returns the live step + ETA. When not
    running, returns Idle + the most recent load's outcome and date.

      [-7]=current_execution_status [-6]=current_execution_step
      [-11]=last_run_outcome [-12]=last_run_time [-13]=last_run_date
    """
    row = _sp_help_job(server, name)
    if not row:
        return ("(no data)", [])

    status, step = row[-7], row[-6]
    if status == '1':                        # Executing -> step label + ETA line
        steps = _steps(server, name)
        # sp_help_job returns current_execution_step ALREADY decorated with the step
        # name, e.g. '1 (Load)', so take the LEADING integer rather than assuming a
        # bare int (the trap that broke the HRP label on 2026-08-05).
        m = re.match(r'\s*(\d+)', str(step))
        sid = int(m.group(1)) if m else None
        if steps and len(steps) == 1:
            # Single-step job: 'Step 1/1 - Load' says nothing useful.
            head = "Executing"
        else:
            head = f"Executing Step {sid if sid is not None else step}"
            if steps and sid is not None:
                head += f"/{max(steps)}"
                label = steps.get(sid, (None,))[0]
                if label:
                    head += f" - {label}"
        return (head, eta_detail(server, name))

    # Idle. Unlike the HRP/Rx digests (daily feeds, which revert to a bare '- Idle'
    # at the start of the next day), Subro is MONTHLY, so keep showing the last
    # run's outcome + date and mark whether it landed today. The text is constant
    # between runs, so content dedupe still holds this to one post per change.
    oc = RUN_OUTCOME.get(row[-11], row[-11])
    if oc in ('Succeeded', 'Failed'):
        comp = last_completion(server, name)
        ctext = comp.strftime('%m/%d/%Y %I:%M %p') if comp else fmt_dt(row[-13], row[-12])
        if comp and comp.date() == datetime.now().date():
            icon = ':white_check_mark:' if oc == 'Succeeded' else ':x:'
            word = 'Successful' if oc == 'Succeeded' else 'Failed'
            return ("", [f"{icon} {word} {ctext}"])
        icon = ':hourglass_flowing_sand:' if oc == 'Succeeded' else ':x:'
        return ("", [f"{icon} Idle - last run {oc} {ctext}"])
    st = EXEC_STATUS.get(status, f'State {status}')
    return (f"- {st}", [])


def snap_is_current(load_jobid, snap_jobid):
    """True if the snap's latest run belongs to the CURRENT load -- i.e. the snap
    STARTED at/after the load's latest completion. A load that hasn't completed yet
    has no current snap (returns False), so a snap from a prior cycle is never
    credited to this one."""
    load_end = _to_dt(job_run(load_jobid).get('EndDate'))
    snap_start = _to_dt(job_run(snap_jobid).get('StartDate'))
    return bool(load_end and snap_start and snap_start >= load_end)


def snap_line(load_jobid, snap_jobid):
    """Status body for the Snap RAMP job. Do NOT show the snap as Successful until
    it has finished for the CURRENT load; a resolved snap predating the current
    load's completion (or a load still running) is stale -> shown as waiting, not a
    green check. A Failed snap still surfaces as a red X regardless."""
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
    return (f"{RUNNING_ICON} Running", f"started {fmt(start)} | not yet complete")


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


# Subro files are named M<YY><MM><suffix>: M2607G.TXT (groups), M2607M1/M2607M2
# (members), M2607P (providers). The embedded YYMM is the DATA MONTH, which is the
# month BEFORE the load -- worth showing so a stale batch is obvious.
SUBRO_FILE_LIKE = "M[0-9][0-9][0-9][0-9]%"


def _parse_subro_month(name):
    """Parse the embedded YYMM from M<YY><MM><suffix> -> a date (1st of month)."""
    m = re.match(r'^M(\d{2})(\d{2})', name or '')
    if not m:
        return None
    try:
        return datetime(2000 + int(m.group(1)), int(m.group(2)), 1)
    except ValueError:
        return None


def last_stage_batch():
    """The files the LAST 'Aetna 0100 Subro Stage' staged, from RAMP FileLog.
    Returns (stage_qid, stage_end_datetime, [(filename, data_month), ...]), or
    (None, None, []) if unavailable. The stage QueueId is the newest one whose
    FileLog actually holds M<YYMM> files (a stage job also logs fileless phases)."""
    qrows = _ramp_sql(
        "SELECT TOP 1 fl.QueueId FROM [ramp].[FileLog] fl "
        "JOIN [ramp].[Queue] q ON q.QueueId = fl.QueueId "
        f"WHERE q.JobId = {STAGE_JOBID} "
        f"AND fl.FileName LIKE '{SUBRO_FILE_LIKE}' ORDER BY fl.QueueId DESC")
    if not qrows:
        return None, None, []
    qid = qrows[0][0]
    erows = _ramp_sql(
        f"SELECT CONVERT(varchar(19), EndDate, 121) FROM [ramp].[Queue] WHERE QueueId = {qid}")
    stage_end = _to_dt(erows[0][0]) if erows and erows[0] else None
    frows = _ramp_sql(
        "SELECT FileName FROM [ramp].[FileLog] "
        f"WHERE QueueId = {qid} AND FileName LIKE '{SUBRO_FILE_LIKE}' ORDER BY FileName")
    files = [(r[0], _parse_subro_month(r[0])) for r in frows if r and r[0]]
    return qid, stage_end, files


def batch_state(stage_end):
    """How far the last Stage's batch has progressed through Load -> Snap, as
    (icon, label). 'snapped' requires the snap to have run for the CURRENT load."""
    load = job_run(SUBRO_JOBID)
    l_start = _to_dt(load.get('StartDate')); l_end = _to_dt(load.get('EndDate'))
    # Is the latest Load run the one for this stage batch (started after staging)?
    load_for_batch = bool(l_start and stage_end and l_start >= stage_end)
    if load_for_batch and load.get('Status') == 'Failed' and l_end:
        return ':x:', 'load FAILED'
    if load_for_batch and l_end and load.get('Status') in RAMP_OK:
        if snap_is_current(SUBRO_JOBID, SNAP_JOBID) \
                and job_run(SNAP_JOBID).get('Status') in RAMP_OK:
            return ':white_check_mark:', 'loaded + snapped'
        return ':white_check_mark:', 'loaded (snap pending)'
    if load_for_batch and not l_end:
        return ':hourglass_flowing_sand:', 'loading'
    return ':hourglass_flowing_sand:', 'staged, pending load'


def subro_file_lines(files):
    """Slack lines for the file section: the whole batch from the last Stage, each
    tagged with its data month. The batch's Load/Snap progress is carried in the
    section header (see main)."""
    if not files:
        return ["- (RAMP FileLog unavailable / no M<YYMM> files in last stage)"]
    out = []
    for name, dt in files:
        dstr = dt.strftime('%m/%Y') if dt else '?'
        out.append(f"- {name}  ({dstr})")   # plain text (no markup renders)
    return out


def _active_today():
    """True if the SQL job is currently Executing or finished (Succeeded/Failed)
    today -- i.e. a real load cycle happened today to report on. Gates the
    --evening extension so evening ticks stay silent on no-load days; the daytime
    slots aren't gated by this."""
    for server, name, _label in SQL_JOBS:
        row = _sp_help_job(server, name)
        if not row:
            continue
        if row[-7] == '1':                       # Executing right now
            return True
        comp = last_completion(server, name)
        if comp and comp.date() == datetime.now().date():
            return True
    return False


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

    # Evening extension: outside the normal slots the tick calls this with
    # --evening so a load finishing after the last daytime slot still gets its
    # Executing->Successful transition posted. Stay silent unless a load actually
    # ran today (else no-load evenings would post a stale idle line).
    if '--evening' in sys.argv and not force and not _active_today():
        print("NO_POST: evening extension, no load active/completed today")
        return

    # PLAIN-TEXT format: the webhook renders only :emoji: (no *bold*/`code`/color).
    lines = ["AETNA SUBRO - STATUS UPDATE", ""]
    for server, name, label in SQL_JOBS:
        status_text, detail = sql_job(server, name)
        lines.append(f"{label} {status_text}".rstrip())
        lines.extend(detail)
        lines.append("")

    head, detail = ramp_line(SUBRO_JOBID)
    lines.append(f"Aetna 0110 Subro Load {head}".rstrip())
    lines.append(detail)
    lines.append("")

    head, detail = snap_line(SUBRO_JOBID, SNAP_JOBID)
    lines.append(f"Aetna 0120 Subro Start Snap {head}".rstrip())
    lines.append(detail)
    lines.append("")

    _stage_qid, stage_end, files = last_stage_batch()
    _icon, state_label = batch_state(stage_end)
    staged_on = stage_end.strftime('%m/%d/%Y') if stage_end else '?'
    lines.append(f"Subro Files - last Aetna 0100 Subro Stage   (staged {staged_on}, {state_label})")
    lines.extend(subro_file_lines(files))
    msg = "\n".join(lines)

    # Content dedupe: post only when the status text CHANGES, so the Successful
    # line posts once when the load finishes and then holds until the next monthly
    # load starts (the message flips back to Executing).
    if not force and msg == _last_msg():
        print("NO_POST: status unchanged since last post")
        return
    _claim_slot(msg)
    print("SLACK|" + msg.replace("\n", "\\n"))


if __name__ == '__main__':
    main()
