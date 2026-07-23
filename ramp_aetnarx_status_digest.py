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


# ---- Live final-stage signal from SSISDB (per user 2026-07-23) ----------------
# Mirrors the HRP digest: input volume does NOT predict this load's duration and
# the Agent job is a single dominant step, so there's no reliable early/mid-run
# ETA. But the underlying SSIS package (SSISDB on TRGETLPROD2) has a late task that
# reliably COMPLETES at a stable ~88% of wall-clock; once the live run passes it we
# project ETA = run_start + elapsed / that_fraction. The milestone is chosen
# dynamically each call so it self-heals across redeploys; any error -> None -> the
# anchored median is used instead.
SSIS_SERVER = 'TRGETLPROD2'
SSIS_PKG = 'AetnaRx_MasterLoad_Claims_And_Eligibility.dtsx'
FS_MAX_FRAC = 0.95   # ignore terminal cleanup tasks pinned at ~99-100% (no lead time)
FS_MAX_SD = 0.08     # milestone must complete at a stable fraction across runs


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


def _ssis_milestone():
    """(executable_id, median_tail_seconds) for the latest STABLE late milestone
    task over the last 6 successful runs, or (None, None). Milestone = latest task
    completing at a consistent fraction of wall-clock (sd <= FS_MAX_SD, <=
    FS_MAX_FRAC for lead time); we return the MEDIAN tail (run_end - milestone_end),
    which is roughly fixed in absolute time regardless of load length (validated
    2026-07-23), so ETA = milestone_end + median_tail is robust on outlier runs."""
    rows = _ssis_sql(
        f"DECLARE @pkg sysname=N'{SSIS_PKG}'; "
        ";WITH ex AS (SELECT TOP 6 execution_id, start_time, end_time, "
        "  DATEDIFF(second,start_time,end_time) AS total_sec "
        "  FROM catalog.executions WHERE package_name=@pkg AND status=7 "
        "  AND end_time IS NOT NULL ORDER BY execution_id DESC), "
        "f AS (SELECT es.executable_id, "
        "  CAST(DATEDIFF(second,ex.start_time,es.end_time) AS float)"
        "  /NULLIF(ex.total_sec,0) AS frac, "
        "  DATEDIFF(second,es.end_time,ex.end_time) AS tail "
        "  FROM catalog.executable_statistics es "
        "  JOIN ex ON ex.execution_id=es.execution_id WHERE ex.total_sec>600), "
        "pick AS (SELECT TOP 1 executable_id FROM f GROUP BY executable_id "
        f"  HAVING COUNT(*)>=4 AND AVG(frac)<={FS_MAX_FRAC} "
        f"  AND ISNULL(STDEV(frac),1)<={FS_MAX_SD} ORDER BY AVG(frac) DESC) "
        "SELECT p.executable_id, (SELECT DISTINCT PERCENTILE_CONT(0.5) "
        "  WITHIN GROUP (ORDER BY CAST(f2.tail AS float)) OVER() "
        "  FROM f f2 WHERE f2.executable_id=p.executable_id) FROM pick p;")
    if rows and len(rows[0]) >= 2:
        try:
            return int(rows[0][0]), float(rows[0][1])
        except Exception:
            pass
    return None, None


def _ssis_final_stage_eta():
    """If the SSIS package is running AND has already completed its stable late
    milestone task, return projected completion = milestone_end + median_tail.
    Else None (the JOIN yields no row until the milestone completes -- exactly the
    'not in the final stretch yet' case)."""
    mid, tail = _ssis_milestone()
    if not mid or tail is None or tail < 0:
        return None
    # Pick the LATEST genuinely-current running execution (recency guard excludes
    # orphaned executions left stuck at status=2), THEN check whether IT has
    # completed the milestone. Doing it in that order matters: a real run that
    # hasn't reached its milestone yet must return None (fall back to the anchored
    # median), not silently match some older execution that happens to have one.
    rows = _ssis_sql(
        f"DECLARE @pkg sysname=N'{SSIS_PKG}'; DECLARE @mid bigint={mid}; DECLARE @rid bigint; "
        "SELECT TOP 1 @rid=execution_id FROM catalog.executions "
        "  WHERE package_name=@pkg AND status=2 "
        "  AND start_time>DATEADD(day,-3,GETDATE()) ORDER BY execution_id DESC; "
        "SELECT CONVERT(varchar(19),s.end_time,120) FROM catalog.executable_statistics s "
        "  WHERE s.execution_id=@rid AND s.executable_id=@mid;")
    if not rows:
        return None
    mend = _to_dt(rows[0][0])
    if not mend:
        return None
    return mend + timedelta(seconds=tail)


def eta_detail(server, name):
    """Single expected-completion ETA for the in-flight run (mirrors HRP). If the
    SSIS package has reached its stable late milestone, show the tighter live
    'final stage' ETA (milestone_end + median tail); otherwise anchor to the live
    run START + MEDIAN (p50) of recent successful full-run durations, bumping
    median->p75->'running longer than usual' as the run outlasts history. Degrades
    to 'now + median' if the live run start can't be read."""
    fs = _ssis_final_stage_eta()
    if fs:
        if fs > datetime.now():
            return [f"{EXEC_ICON} final stage - ETA ~{_eta_stamp(fs)}"]
        return [f"{EXEC_ICON} final stage - wrapping up"]
    durs = sorted(_recent_full_durations(server, name))
    if not durs:
        return [f"{EXEC_ICON} in progress"]
    est, hi = _pct(durs, 50), _pct(durs, 75)
    start = _current_run_start(server, name)
    if not start:
        eta = datetime.now() + timedelta(seconds=est)
        return [f"{EXEC_ICON} ETA ~{_eta_stamp(eta)}"]
    elapsed = (datetime.now() - start).total_seconds()
    if elapsed > hi:
        return [f"{EXEC_ICON} running {_dur_h(elapsed)} - longer than usual, still processing"]
    eta = start + timedelta(seconds=(est if elapsed < est else hi))
    return [f"{EXEC_ICON} ETA ~{_eta_stamp(eta)}"]


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
        # ETA anchored to the live run's start (see eta_detail), superseded by the
        # live SSISDB final-stage ETA once ~88% done. Icon = cycling-arrows.
        return (f"Executing Step {step}", eta_detail(server, name))
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
