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


def _load_state():
    try:
        with open(STATE_FILE) as f:
            return json.load(f)
    except Exception:
        return {}


def _recent_emit():
    try:
        last = datetime.fromisoformat(_load_state()['last_emit'])
    except Exception:
        return None
    return last if datetime.now() - last < timedelta(minutes=DEDUPE_MINUTES) else None


def _last_msg():
    """Text of the digest we most recently POSTED (for content dedupe), or None."""
    return _load_state().get('last_msg')


def _claim_slot(msg=None):
    """Stamp now as the last-emit time (claim the slot). If msg is given, also
    record it as the last-posted message so an identical later digest is skipped
    (content dedupe)."""
    st = _load_state()
    st['last_emit'] = datetime.now().isoformat()
    if msg is not None:
        st['last_msg'] = msg
    tmp = STATE_FILE + '.tmp'
    with open(tmp, 'w') as f:
        json.dump(st, f)
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
    out = subprocess.run(['curl', '-s', '--negotiate', '-u', ':',
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


# ---- ETA (anchored to the actual run start, mirrors the HRP digest) ----------
# ANCHORING FIX 2026-07-23 (per user): the OLD step-based `remaining_secs` added
# the FULL historical average of every step >= the current step onto 'now'. But
# this job is effectively a SINGLE dominant step ('Load Claims And Eligibility
# files' is ~99% of the run; the other two steps are seconds/minutes), so at any
# tick it was Executing step 1 and the ETA became now + avg(step 1 duration) --
# ignoring the hours already spent IN step 1. That's why a run already ~13 h deep
# showed an ETA ~5.5 h further out (the reported 5:39 PM). We now anchor to the
# live run's START + the MEDIAN of recent successful FULL-run durations (step_id=0)
# exactly like HRP: it accounts for elapsed time, stays fixed across ticks, bumps
# median->p75->'running longer than usual' as the run outlasts history, and is
# superseded by the live SSISDB final-stage ETA once the run is ~90% done.


HISTORY_N = 30          # runs in the ETA history window (~6 weeks of daily loads)
LONG_RUN_FACTOR = 1.15  # once past all history, project 15% more run time to go


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


def _current_run_start(server, name):
    """Start datetime of the CURRENTLY-executing run from sysjobactivity (the row
    on the latest Agent session with no stop time), or None. Don't join syssessions
    (SELECT is permission-denied); ordering by session_id DESC picks the current
    run's row directly."""
    q = ("SET NOCOUNT ON; "
         f"DECLARE @jid uniqueidentifier=(SELECT job_id FROM msdb.dbo.sysjobs WHERE name=N'{name}'); "
         "SELECT TOP 1 CONVERT(varchar(19), ja.start_execution_date, 120) "
         "FROM msdb.dbo.sysjobactivity ja WITH (NOLOCK) "
         "WHERE ja.job_id=@jid AND ja.start_execution_date IS NOT NULL "
         "AND ja.stop_execution_date IS NULL "
         # Skip orphaned rows left by Agent restarts (same guard as _live_step).
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


NOTE_ICON = ':information_source:'   # neutral on purpose -- a long run is NOT a failure,
                                     # so this never uses the red/failure marker (per user
                                     # 2026-07-27).


def _why_longer(durs, beyond_history, projected=False):
    """One-line, HISTORY-DERIVED explanation of why the ETA moved out (per user
    2026-07-27: "a small note about why the ETA is longer"). Every number in it
    comes from this job's OWN recent successful runs -- fastest, slowest, typical,
    and where the live run sits against them -- so nothing is hardcoded and it
    can't go stale as the load's profile changes.

    `projected` says whether the caller is still showing a clock ETA extrapolated
    from elapsed time (Rx) or has no ETA left to show (HRP/RCE), so the wording
    matches what's actually on the line above."""
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


def _ceil_15(dt):
    """Round up to the next 15-minute mark (extrapolated ETAs are not precise to
    the minute, so don't present them that way)."""
    dt = dt.replace(second=0, microsecond=0)
    return dt + timedelta(minutes=(-dt.minute) % 15 or 15)


# ---- Live final-stage signal from SSISDB (per user 2026-07-23) ----------------
# Mirrors the HRP digest. Input volume does NOT predict this load's duration, and
# the SSIS package is 99.0% of the job's wall-clock (the T-SQL steps 2-4 median
# ~2 min), so the package's own task log is the ONLY mid-run signal there is --
# per-step history from sysjobhistory has essentially nothing left to pace here.
SSIS_SERVER = 'TRGETLPROD2'
SSIS_PKG = 'AetnaRx_MasterLoad_Claims_And_Eligibility.dtsx'

# ---- SSIS milestone LADDER (per user 2026-08-05) ------------------------------
# Was: ONE milestone task at ~84% of the package, leaving the ETA on whole-job
# history for everything before it. The package logs EVERY task completion, so we
# now admit a ladder of rungs and use the LATEST one the live run has passed.
# Backtested walk-forward over 236 runs (_eta_backtest.py): median error 41m ->
# 33m, p90 7h09m -> 5h04m, bias -30m -> -17m, with a rung available on 36% of
# in-SSIS ticks.
#
# Admissibility is about the TAIL, not the fraction: the ETA is rung_end +
# median_tail, so the error IS the tail's spread. The old sd(frac) test would
# happily admit a task that finishes at a rock-steady 0.1% of wall-clock and
# carries a useless multi-hour tail.
LADDER_MIN_SAMPLES = 5      # a rung needs real history before we trust its tail
LADDER_MIN_FRAC = 0.55      # before this, a completion carries no information
LADDER_MAX_FRAC = 0.97      # after this there's no lead time left to be useful
LADDER_IQR_FLOOR = 20 * 60  # allow at least this much tail spread...
LADDER_IQR_REL = 0.35       # ...or this share of the tail, whichever is larger...
LADDER_IQR_CAP = 90 * 60    # ...but never more than this in absolute terms.

# Survival percentile used inside the live step (was the p50 median). Overruns are
# one-sided, so a median runs systematically early; p80 beat p50/p60/p70 on median
# error, p90 error AND bias across the Aetna jobs. Kept identical in all three
# digests even though this job's T-SQL tail is too short for it to matter much.
STEP_PCT = 80


def _ssis_sql(query):
    """Run a query against SSISDB on TRGETLPROD2; rows as lists of stripped fields."""
    try:
        out = subprocess.run(
            ['sqlcmd', '-S', SSIS_SERVER, '-d', 'SSISDB', '-E', '-W',
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


def _ssis_ladder():
    """{executable_id: median_tail_seconds} -- every package task usable as an ETA
    rung, judged over the last 8 successful executions. {} on any read error, which
    simply drops the caller back to whole-job history.

    A rung qualifies when it has LADDER_MIN_SAMPLES history, completes inside
    [LADDER_MIN_FRAC, LADDER_MAX_FRAC] of package wall-clock, and its tail
    (run_end - task_end) has a bounded interquartile spread. The tail is used in
    ABSOLUTE time rather than as a fraction of elapsed because the finalization
    work after a rung is roughly fixed no matter how long the heavy load ran
    (validated 2026-07-23: a 14 h Rx run and a 2.5 h one both had ~25 min tails).

    Chosen dynamically rather than hardcoded so the ladder self-heals across
    package redeploys -- a redeploy renumbers executable_ids, and unknown ids
    simply fail the sample-count test until they build history."""
    rows = _ssis_sql(
        f"DECLARE @pkg sysname=N'{SSIS_PKG}'; "
        ";WITH ex AS (SELECT TOP 8 execution_id, start_time, end_time, "
        "  DATEDIFF(second,start_time,end_time) AS total_sec "
        "  FROM catalog.executions WHERE package_name=@pkg AND status=7 "
        "  AND end_time IS NOT NULL ORDER BY execution_id DESC), "
        "f AS (SELECT es.executable_id, "
        "  CAST(DATEDIFF(second,ex.start_time,es.end_time) AS float)"
        "  /NULLIF(ex.total_sec,0) AS frac, "
        "  CAST(DATEDIFF(second,es.end_time,ex.end_time) AS float) AS tail "
        "  FROM catalog.executable_statistics es "
        "  JOIN ex ON ex.execution_id=es.execution_id WHERE ex.total_sec>600), "
        "s AS (SELECT DISTINCT executable_id, "
        "  COUNT(*) OVER (PARTITION BY executable_id) AS n, "
        "  AVG(frac) OVER (PARTITION BY executable_id) AS avg_frac, "
        "  PERCENTILE_CONT(0.50) WITHIN GROUP (ORDER BY tail) "
        "    OVER (PARTITION BY executable_id) AS med_tail, "
        "  PERCENTILE_CONT(0.25) WITHIN GROUP (ORDER BY tail) "
        "    OVER (PARTITION BY executable_id) AS p25, "
        "  PERCENTILE_CONT(0.75) WITHIN GROUP (ORDER BY tail) "
        "    OVER (PARTITION BY executable_id) AS p75 FROM f) "
        "SELECT executable_id, CAST(med_tail AS int) FROM s "
        f"WHERE n>={LADDER_MIN_SAMPLES} AND med_tail>=0 "
        f"  AND avg_frac>={LADDER_MIN_FRAC} AND avg_frac<={LADDER_MAX_FRAC} "
        f"  AND (p75-p25)<={LADDER_IQR_CAP} "
        f"  AND (p75-p25)<=(CASE WHEN {LADDER_IQR_FLOOR}>{LADDER_IQR_REL}*med_tail "
        f"                       THEN {LADDER_IQR_FLOOR} ELSE {LADDER_IQR_REL}*med_tail END);")
    ladder = {}
    for r in rows:
        if len(r) >= 2 and r[0].lstrip('-').isdigit() and r[1].lstrip('-').isdigit():
            ladder[int(r[0])] = int(r[1])
    return ladder


NOTE_ICON = ':information_source:'   # neutral on purpose -- a long run is NOT a failure,
                                     # so this never uses the red/failure marker (per user
                                     # 2026-07-27).


def _why_longer(durs, beyond_history, projected=False):
    """One-line, HISTORY-DERIVED explanation of why the ETA moved out (per user
    2026-07-27: "a small note about why the ETA is longer"). Every number in it
    comes from this job's OWN recent successful runs -- fastest, slowest, typical,
    and where the live run sits against them -- so nothing is hardcoded and it
    can't go stale as the load's profile changes. Kept identical across the three
    Aetna digests (they're standalone clones, no shared module).

    `projected` says whether the caller is still showing a clock ETA extrapolated
    from elapsed time (Rx) or has no ETA left to show (HRP/RCE), so the wording
    matches what's actually on the line above."""
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


def _ssis_final_stage_eta():
    """Projected package completion from the LATEST ladder rung the live execution
    has passed: rung_end + that rung's median tail. None when the package isn't
    running or hasn't reached any rung yet -- exactly the 'no finer signal than
    whole-job history' case.

    Because later rungs have shorter, tighter tails, taking the latest one passed
    means the estimate tightens progressively through the package instead of
    staying blind until a single ~84% task fires, which on 2-hourly ticks it
    frequently never did within the life of a run."""
    ladder = _ssis_ladder()
    if not ladder:
        return None
    ids = ','.join(str(int(k)) for k in ladder)
    # Pick the LATEST genuinely-current running execution (the recency guard
    # excludes orphaned executions left stuck at status=2 -- there is one from
    # 2023), THEN ask which rungs IT has passed. Doing it in that order matters: a
    # real run that has passed no rung yet must return None and fall back, not
    # silently match an older execution that happens to have rung data.
    rows = _ssis_sql(
        f"DECLARE @pkg sysname=N'{SSIS_PKG}'; DECLARE @rid bigint; "
        "SELECT TOP 1 @rid=execution_id FROM catalog.executions "
        "  WHERE package_name=@pkg AND status=2 "
        "  AND start_time>DATEADD(day,-3,GETDATE()) ORDER BY execution_id DESC; "
        "SELECT TOP 1 CONVERT(varchar(19),s.end_time,120), s.executable_id "
        "  FROM catalog.executable_statistics s "
        f"  WHERE s.execution_id=@rid AND s.executable_id IN ({ids}) "
        "  AND s.end_time IS NOT NULL ORDER BY s.end_time DESC;")
    if not rows or len(rows[0]) < 2:
        return None
    rend = _to_dt(rows[0][0])
    try:
        tail = ladder[int(rows[0][1])]
    except (ValueError, KeyError):
        return None
    if not rend:
        return None
    return rend + timedelta(seconds=tail)


# ---- Step-aware ETA (per user 2026-08-04) ------------------------------------
# Ported from the HRP digest after the 08/04 HRP miss (10:17 tick sat on step 6
# of 9 with ~16 min left and reported 'ETA ~6:56 PM' for a job that ended 10:33
# AM) -- the cause was consulting WHOLE-JOB history only, which this job did too.
# This job is 4 steps: step 1 is the SSIS package and steps 2-4 are a short fixed
# tail (median ~2 min, max 12 min over the last 61 runs). So once step 1 is done
# the answer is 'about two minutes', while the whole-job survival median could
# still be quoting hours out. Same class of error, smaller magnitude.
STALE_ACTIVITY_DAYS = 4   # > the longest run on record (18.4 h); see _live_step


def _live_step(server, name):
    """(step_id, step_start) for the step the CURRENT run is executing, read from
    msdb.dbo.sysjobactivity -- what SSMS's Job Activity Monitor shows. SQL Agent
    stamps last_executed_step_id/_date when a step STARTS, so this is the in-flight
    step and the moment it began. None when not running or unreadable.

    The recency guard skips orphaned rows: an Agent restart leaves a row with a
    NULL stop_execution_date, which would otherwise read as a 'live' step."""
    q = ("SET NOCOUNT ON; "
         f"DECLARE @jid uniqueidentifier=(SELECT job_id FROM msdb.dbo.sysjobs WHERE name=N'{name}'); "
         "SELECT TOP 1 ja.last_executed_step_id, "
         "  CONVERT(varchar(19), ja.last_executed_step_date, 120) "
         "FROM msdb.dbo.sysjobactivity ja WITH (NOLOCK) "
         "WHERE ja.job_id=@jid AND ja.start_execution_date IS NOT NULL "
         "AND ja.stop_execution_date IS NULL AND ja.last_executed_step_id IS NOT NULL "
         f"AND ja.start_execution_date > DATEADD(day,-{STALE_ACTIVITY_DAYS},GETDATE()) "
         "ORDER BY ja.session_id DESC, ja.start_execution_date DESC;")
    try:
        out = subprocess.run(['sqlcmd', '-S', server, '-E', '-W', '-h', '-1', '-s', '|', '-Q', q],
                             capture_output=True, text=True, timeout=120)
    except Exception:
        return None
    for line in out.stdout.splitlines():
        parts = [p.strip() for p in line.split('|')]
        if len(parts) >= 2 and parts[0].isdigit():
            dt = _to_dt(parts[1])
            if dt:
                return int(parts[0]), dt
    return None


def _step_remaining_secs(server, name, step_id, days=120):
    """Ascending historical wall-clock (seconds) from the START of `step_id` through
    the END of the job, over recent runs that finished successfully.

    sysjobhistory has no run key, so each step row is attached to the next
    step_id=0 'job outcome' row after it -- bucketed against ALL outcome rows and
    only then filtered to successful ones, since bucketing against successes alone
    would graft a failed run's steps onto the next good run."""
    q = ("SET NOCOUNT ON; "
         f"DECLARE @jid uniqueidentifier=(SELECT job_id FROM msdb.dbo.sysjobs WHERE name=N'{name}'); "
         f"DECLARE @sid int={int(step_id)}; "
         "WITH h AS (SELECT instance_id, step_id, run_status, "
         "  (run_duration/10000)*3600+((run_duration/100)%100)*60+(run_duration%100) AS secs "
         "  FROM msdb.dbo.sysjobhistory WITH (NOLOCK) WHERE job_id=@jid "
         f"  AND run_date>=CONVERT(int,CONVERT(varchar(8),DATEADD(day,-{int(days)},GETDATE()),112))), "
         "o AS (SELECT instance_id, run_status FROM h WHERE step_id=0), "
         "t AS (SELECT h.step_id, h.secs, h.run_status, "
         "  (SELECT MIN(o.instance_id) FROM o WHERE o.instance_id>h.instance_id) AS rk "
         "  FROM h WHERE h.step_id>=@sid) "
         "SELECT SUM(t.secs) FROM t JOIN o ON o.instance_id=t.rk AND o.run_status=1 "
         "GROUP BY t.rk "
         "HAVING MIN(t.run_status)=1 AND MAX(CASE WHEN t.step_id=@sid THEN 1 ELSE 0 END)=1 "
         "ORDER BY 1;")
    try:
        out = subprocess.run(['sqlcmd', '-S', server, '-E', '-W', '-h', '-1', '-Q', q],
                             capture_output=True, text=True, timeout=120)
    except Exception:
        return []
    return sorted(int(s) for s in (l.strip() for l in out.stdout.splitlines()) if s.isdigit())


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


def _step_names(server, name):
    """{step_id: step_name}."""
    return {k: v[0] for k, v in _steps(server, name).items()}


def _ssis_step_id(server, name):
    """The LAST step whose subsystem is SSIS -- the heavy package load, after which
    only the short finishers remain. Read from sysjobsteps rather than hardcoded so
    it self-heals if a step is inserted ahead of the package. Defaults to 1."""
    ssis = [k for k, v in _steps(server, name).items() if v[1].upper() == 'SSIS']
    return max(ssis) if ssis else 1


def _step_eta(server, name, sid, sstart):
    """ETA lines projected from the live step's start, or None to fall through.
    Survival-filtered within the step, like the whole-job estimator: drop the
    outcomes already ruled out by how long this step has been running, then take
    the median of what remains."""
    # Check for the LAST step before consulting history. The trailing 'Success
    # Step' is a 0-second no-op, so it either writes no history row at all (HRP
    # step 9 has 3 rows, Rx none) or writes rows of all zeros (RCE step 12 has 12).
    # The all-zeros case is the trap: every historical outcome is 0s, so any time
    # at all in the step empties the survival filter and the old ordering reported
    # 'running long, still processing' for a step with nothing left to do.
    steps = _steps(server, name)
    if steps and sid >= max(steps):
        return [f"{EXEC_ICON} final steps - wrapping up"]
    rem = _step_remaining_secs(server, name, sid)
    if not rem:
        return None
    in_step = (datetime.now() - sstart).total_seconds()
    possible = [d for d in rem if d >= in_step]
    if not possible:
        return [f"{EXEC_ICON} final steps - running long, still processing"]
    eta = sstart + timedelta(seconds=_pct(possible, STEP_PCT))
    if eta <= datetime.now():
        return [f"{EXEC_ICON} final steps - wrapping up"]
    return [f"{EXEC_ICON} ETA ~{_eta_stamp(eta)}"]


def eta_detail(server, name):
    """Single expected-completion ETA for the in-flight run. An ETA is ALWAYS
    given while the job is executing (per user 2026-07-27) -- see the long-run
    handling below.

    Order of preference:
      1. SSIS live 'final stage' ETA (milestone_end + median tail) once the package
         passes its stable late milestone -- the tightest signal available.
      2. CONDITIONAL (survival) median, replacing the old p50->p75 ladder: anchor to
         the live run's actual START and take the median of only those historical
         durations still POSSIBLE (>= elapsed). Runs shorter than we've already been
         running are ruled out, so the ETA climbs step-wise as each shorter outcome
         is eliminated and converges on the real finish. Same approach as HRP.
      3. Past EVERY run in the history window: extrapolate from the run itself
         (elapsed * LONG_RUN_FACTOR, rounded up to the next 15 min) instead of
         dropping the ETA. This is where the old code gave up and posted only
         'running Xh - longer than usual, still processing' with no time at all.

    The window is HISTORY_N runs rather than 8: this load is bimodal (~2h to ~18h),
    so the extra tail samples let the estimate step out gradually instead of jumping
    straight to the slowest run on record. Degrades to 'now + median' if the live
    run start can't be read.

    Ahead of all three (per user 2026-08-04): if the run is past the SSIS step, the
    only work left is the short fixed tail, so project from the live step instead
    -- see _step_eta. Whole-job history is consulted only while the package itself
    is running, the one phase with no finer signal."""
    ssis_step = _ssis_step_id(server, name)
    live = _live_step(server, name)
    if live and live[0] > ssis_step:
        stepped = _step_eta(server, name, live[0], live[1])
        if stepped:
            return stepped
    fs = _ssis_final_stage_eta()
    if fs:
        # The milestone's median tail is measured to the PACKAGE's end, but this
        # digest reports on the JOB, which still has its finisher steps to run
        # after the package exits. Without this the 'final stage' ETA lands early.
        post = _step_remaining_secs(server, name, ssis_step + 1)
        if post:
            fs += timedelta(seconds=_pct(post, 50))
        if fs > datetime.now():
            return [f"{EXEC_ICON} final stage - ETA ~{_eta_stamp(fs)}"]
        return [f"{EXEC_ICON} final stage - wrapping up"]
    now = datetime.now()
    durs = sorted(_recent_full_durations(server, name, HISTORY_N))
    if not durs:
        return [f"{EXEC_ICON} in progress"]
    start = _current_run_start(server, name)
    if not start:
        eta = now + timedelta(seconds=_pct(durs, 50))
        return [f"{EXEC_ICON} ETA ~{_eta_stamp(eta)}"]
    elapsed = (now - start).total_seconds()
    still_possible = [d for d in durs if d >= elapsed]
    if still_possible:
        eta, beyond_history = start + timedelta(seconds=_pct(still_possible, 50)), False
    else:
        eta, beyond_history = _ceil_15(start + timedelta(seconds=elapsed * LONG_RUN_FACTOR)), True
    # Never show a clock time that is already here/past -- keep a little lead time.
    if eta < now + timedelta(minutes=10):
        eta = _ceil_15(now + timedelta(minutes=10))
    line = f"{EXEC_ICON} ETA ~{_eta_stamp(eta)}"
    if beyond_history:
        return ([f"{line} - running {_dur_h(elapsed)}, past every recent run"]
                + _why_longer(durs, True, projected=True))
    if elapsed > _pct(durs, 75):
        return ([f"{line} - running {_dur_h(elapsed)}, longer than usual"]
                + _why_longer(durs, False))
    return [line]


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
    if status == '1':                        # Executing -> "Executing Step N/M - name" + ETA line
        # ETA anchored to the live run's start (see eta_detail), superseded by the
        # live SSISDB final-stage ETA once ~88% done. Icon = cycling-arrows.
        # Step count + name added 2026-08-04: 'Step 2' alone doesn't convey that
        # the heavy SSIS load is step 1 and everything after it is a short tail.
        names = _step_names(server, name)
        # sp_help_job returns current_execution_step ALREADY decorated with the
        # step name, e.g. '1 (Run AetnaHRP MasterLoad)'. The 2026-08-04 label
        # change assumed a bare integer, so '/N' was appended AFTER the paren
        # ('Executing Step 1 (Run AetnaHRP MasterLoad)/9') and str.isdigit() failed,
        # suppressing the name lookup entirely. Take the leading integer instead
        # (fixed 2026-08-05) -> 'Executing Step 1/9 - Run AetnaHRP MasterLoad'.
        m = re.match(r'\s*(\d+)', str(step))
        sid = int(m.group(1)) if m else None
        head = f"Executing Step {sid if sid is not None else step}"
        if names and sid is not None:
            head += f"/{max(names)}"
            label = names.get(sid)
            if label:
                head += f" - {label}"
        return (head, eta_detail(server, name))
    # Idle: show the last run's outcome as Successful/Failed ONLY while its
    # COMPLETION falls on today's date; at the start of the next day it reverts to
    # "- Idle" (per user 2026-07-17: "only show as Successful until the start of
    # the next day"). Gate on the completion time, NOT sp_help_job's last_run_date
    # (= the START date): an overnight run that started yesterday but finished
    # early today, like AetnaRx, must still count as today. NCStateAetna, which
    # both started and finished yesterday, shows Idle.
    oc = RUN_OUTCOME.get(row[-11], row[-11])
    if oc in ('Succeeded', 'Failed'):
        comp = last_completion(server, name)
        ctext = comp.strftime('%m/%d/%Y %I:%M %p') if comp else fmt_dt(row[-13], row[-12])
        if comp and comp.date() == datetime.now().date():
            if oc == 'Succeeded':
                return ("", [f":white_check_mark: Successful {ctext}"])
            return ("", [f":x: Failed {ctext}"])
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


def _active_today():
    """True if any primary SQL job is currently Executing or finished (Succeeded/
    Failed) today -- i.e. a real load cycle happened today to report on. Gates the
    --evening extension so evening ticks stay silent on no-load days (weekends,
    days the feed didn't run); the daytime 8/12/16 slots aren't gated by this."""
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
    force = '--force' in sys.argv
    if not force:
        recent = _recent_emit()
        if recent:
            print(f"NO_POST: deduped (a digest was already emitted at "
                  f"{recent.strftime('%I:%M %p')}, within {DEDUPE_MINUTES} min)")
            return
    _claim_slot()

    # Evening extension (per user 2026-07-17): outside the normal 8/12/16 slots the
    # tick calls this with --evening so a load finishing after the last daytime slot
    # still gets its Executing->Successful transition posted. Stay silent unless a
    # load actually ran today (else no-load evenings would post a stale idle line).
    if '--evening' in sys.argv and not force and not _active_today():
        print("NO_POST: evening extension, no load active/completed today")
        return

    # Minimal PLAIN-TEXT format (per user 2026-07-16): ONLY the ETL AetnaRx
    # MasterLoad Claims And Eligibility step & ETA. Webhook renders only :emoji:
    # (no markup/color), so the only standout is the icon on the ETA line.
    lines = ["AETNA RX - STATUS UPDATE", ""]
    for server, name, label in SQL_JOBS:
        status_text, detail = sql_job(server, name)
        lines.append(f"{label} {status_text}".rstrip())
        lines.extend(detail)
        lines.append("")
    while lines and lines[-1] == "":
        lines.pop()
    msg = "\n".join(lines)

    # Content dedupe (per user 2026-07-17): post only when the status text CHANGES.
    # This posts the Successful line ONCE when the load finishes, holds quietly
    # while it stays Successful, then posts again when the next load starts (the
    # message flips back to Executing). Replaces the old "succeeded today -> go
    # silent" skip, which left the last post stuck on a stale 'Executing' line.
    if not force and msg == _last_msg():
        print("NO_POST: status unchanged since last post")
        return
    _claim_slot(msg)
    print("SLACK|" + msg.replace("\n", "\\n"))


if __name__ == '__main__':
    main()
