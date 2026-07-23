"""
Every-3-hours (weekday) status digest -> Slack #team-rdp-operations-support.

Combines:
  - RAMP 'Aetna RCE 310 ETL Load' (JobId 2257): status + completion time.
  - SQL Agent Job Activity Monitor (msdb), mirroring SSMS:
      * TRGETL2 'ETL_AetnaSupport_MasterLoad'
      * TRGETL4 'ETL NCStateAetna MasterLoad'

Prints one 'SLACK|<text>' line (newlines escaped as \\n) for the cron to post.
Always emits (it's a periodic status report, not event-driven).

Note: msdb.dbo.agent_datetime is permission-blocked here, so run_date/run_time
are converted to a datetime manually.
"""
import json, os, re, subprocess, sys
from datetime import datetime, timedelta

RCE_JOBID = 2257       # RAMP 'Aetna RCE 310 ETL Load'
SNAP_JOBID = 10053     # RAMP 'Aetna RCE 400 Daily Snap' (added 2026-07-14 per user)
# Extra RAMP jobs added to the digest's RAMP section per user 2026-07-16.
DAILY_MINE_JOBID = 10574    # 'Aetna RCE 450 Daily Mine'
WEEKLY_SNAP_JOBID = 2258    # 'Aetna RCE 330 Weekly Snap'
WEEKLY_MINE_JOBID = 2266    # 'Aetna RCE 350 Weekly Mining'
NCSTATE_SNAP_JOBID = 10738  # 'NCStateAetna 0110 Snap'
CHANNEL = 'C09EPLQL2D9'

# ---- Cross-session dedupe guard (added 2026-06-26) ----------------------------
# The digest cron is session-only and recreated in EVERY open Claude session, so
# if two sessions are idle near the same slot they BOTH fire and BOTH post to
# both channels -> duplicate posts. This guard records the last emit time to a
# shared on-disk file; any run within DEDUPE_MINUTES of the last emit prints
# 'NO_POST: deduped ...' and emits nothing. The slot is CLAIMED (file written)
# before the slow SQL/curl work, so a second session bails almost immediately.
# Window (25 min) > max cron jitter (15 min), < real slot spacing (~2 h).
# Use --force to bypass (e.g. a deliberate manual on-demand post).
DEDUPE_MINUTES = 25
STATE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                          'ramp_aetna_digest_post_state.json')


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

# New minimal format (per user 2026-07-16): show ONLY the step & ETA of these two
# SQL Agent jobs. (server, SQL Agent job name, display label.)
SQL_JOBS = [
    ("TRGETL2", "SSIS AetnaRCE Daily Process", "Aetna RCE Daily Process"),
    ("TRGETL4", "ETL NCStateAetna MasterLoad", "ETL NCStateAetna MasterLoad"),
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
    """Fetch RAMP /Job/List once per process and cache it (the digest queries
    several job ids)."""
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
    """Return the LatestJobRun dict for a RAMP job id (or {} if not found)."""
    for j in _all_jobs():
        if j.get('JobId') == jobid:
            return j.get('LatestJobRun') or {}
    return {}


def rce_run():
    """LatestJobRun for RAMP 'Aetna RCE 310 ETL Load' (2257)."""
    return job_run(RCE_JOBID)


# RAMP terminal-good statuses: mining/load jobs report 'Successful'; snap jobs
# report 'Resolved' (a full snap run resolves its queue item â€” e.g. the RCE 330
# Weekly Snap sits at 'Resolved' after a ~1h run). Both get a green check.
RAMP_OK = ('Successful', 'Resolved')


def _started_today(start):
    dt = _to_dt(start)
    return bool(dt and dt.date() == datetime.now().date())


EXEC_ICON = ':arrows_counterclockwise:'   # in-progress marker for the main line


def ramp_line(jobid):
    """Return (head, detail) for a RAMP job's LatestJobRun. 'head' = emoji +
    status word for the bold main line; 'detail' = the quiet italic sub-line.
    Green check for Successful/Resolved (red X for Failed). A job that has NOT run
    today is shown Idle with its last-run outcome (per user 2026-07-16), which
    suits the weekly jobs (Weekly Snap/Mining) that don't run daily."""
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
    stripped fields (>=32), or None. See sql_job for why it's run raw and parsed
    from the end."""
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


# ---- ETA (step-anchored progress; multi-step analog of the HRP/Rx method) -----
# ANCHORING FIX 2026-07-23 (per user "can AetnaRCE/NCState follow the same method"):
# HRP/Rx are single SSIS packages, so they anchor to run start + median full-run
# duration and tighten via a live SSISDB milestone. RCE/NCState DON'T fit that: the
# 'SSIS AetnaRCE Daily Process' job is genuinely multi-step (12 steps -- the SSIS
# package is only step 2 of 12; the bulk is stored-proc steps 3-11 like Build
# Chimera / Loop Claims / DHTStats), and NCState is a dominant SSIS step 1 + a small
# stored-proc tail. So the "same method" here is applied to the SQL Agent STEP
# progression instead of SSISDB: anchor to the CURRENT step's start + the historical
# average of the current and remaining steps. This (a) fixes the OLD `remaining_secs`
# bug -- it summed avg(step >= current) onto a fresh 'now' EVERY tick, re-adding the
# current step's full average even when we were already hours into it -- and (b)
# tightens naturally as steps complete (the multi-step analog of HRP/Rx's SSISDB
# 'final stage'). Falls back to run_start + median full-run duration if step
# progress can't be read.


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


def _run_progress(server, name):
    """(run_start, current_step_start, current_step_id) for the currently-executing
    run from sysjobactivity, or (None, None, None). A live run has
    stop_execution_date NULL; ordering by session_id DESC picks the current Agent
    session's row (older orphaned NULL rows sort below it)."""
    q = ("SET NOCOUNT ON; "
         f"DECLARE @jid uniqueidentifier=(SELECT job_id FROM msdb.dbo.sysjobs WHERE name=N'{name}'); "
         "SELECT TOP 1 CONVERT(varchar(19),start_execution_date,120), "
         "  CONVERT(varchar(19),last_executed_step_date,120), last_executed_step_id "
         "FROM msdb.dbo.sysjobactivity WITH (NOLOCK) "
         "WHERE job_id=@jid AND start_execution_date IS NOT NULL AND stop_execution_date IS NULL "
         "ORDER BY session_id DESC, start_execution_date DESC;")
    out = subprocess.run(['sqlcmd', '-S', server, '-E', '-W', '-h', '-1', '-s', '|', '-Q', q],
                         capture_output=True, text=True, timeout=120)
    for line in out.stdout.splitlines():
        parts = [p.strip() for p in line.split('|')]
        if len(parts) >= 3 and _to_dt(parts[0]):
            try:
                sid = int(parts[2])
            except Exception:
                sid = None
            return _to_dt(parts[0]), _to_dt(parts[1]), sid
    return None, None, None


def _step_hist(server, name, n=6):
    """{step_id: avg duration seconds} over the last n successful runs per real
    step (1..49; excludes the step_id=0 job-outcome row)."""
    q = ("SET NOCOUNT ON; "
         f"DECLARE @jid uniqueidentifier=(SELECT job_id FROM msdb.dbo.sysjobs WHERE name=N'{name}'); "
         ";WITH h AS (SELECT step_id, "
         "(run_duration/10000)*3600+((run_duration/100)%100)*60+(run_duration%100) AS s, "
         "ROW_NUMBER() OVER (PARTITION BY step_id ORDER BY run_date DESC, run_time DESC) rn "
         "FROM msdb.dbo.sysjobhistory WITH (NOLOCK) WHERE job_id=@jid AND step_id BETWEEN 1 AND 49 "
         "AND run_status=1) "
         f"SELECT step_id, AVG(s) FROM h WHERE rn<={n} GROUP BY step_id ORDER BY step_id;")
    out = subprocess.run(['sqlcmd', '-S', server, '-E', '-W', '-h', '-1', '-s', '|', '-Q', q],
                         capture_output=True, text=True, timeout=120)
    hist = {}
    for line in out.stdout.splitlines():
        parts = [p.strip() for p in line.split('|')]
        if len(parts) >= 2 and parts[0].isdigit() and parts[1].lstrip('-').isdigit():
            hist[int(parts[0])] = int(parts[1])
    return hist


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


def eta_detail(server, name):
    """Progressive expected-completion ETA for the in-flight run. Primary estimate
    anchors to the CURRENT step's start + the historical average duration of the
    current and remaining steps (so it never re-adds time already spent in earlier
    steps, and tightens as the job advances). Falls back to run_start + MEDIAN
    full-run duration (p50->p75->'running longer than usual') when step progress
    can't be read."""
    now = datetime.now()
    start, step_start, sid = _run_progress(server, name)
    hist = _step_hist(server, name)
    if step_start and sid is not None and hist:
        rem_incl = sum(v for s, v in hist.items() if s >= sid)   # current + later steps
        rem_after = sum(v for s, v in hist.items() if s > sid)   # later steps only
        eta = step_start + timedelta(seconds=rem_incl)
        if eta <= now:                       # current step already overran its average
            eta = now + timedelta(seconds=rem_after)
        if eta <= now:                       # on/after the last step -> essentially done
            return [f"{EXEC_ICON} final stage - wrapping up"]
        return [f"{EXEC_ICON} ETA ~{_eta_stamp(eta)}"]
    # Fallback: anchored full-run median (mirrors HRP/Rx).
    durs = sorted(_recent_full_durations(server, name))
    if not durs:
        return [f"{EXEC_ICON} in progress"]
    est, hi = _pct(durs, 50), _pct(durs, 75)
    if not start:
        return [f"{EXEC_ICON} ETA ~{_eta_stamp(now + timedelta(seconds=est))}"]
    elapsed = (now - start).total_seconds()
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
    """Report where the CURRENT load is, via sp_help_job (the value SSMS Job
    Activity Monitor shows). When executing, returns the live step e.g.
    '*Executing*: 4 (Build Chimera)'. When not running, returns Idle + the most
    recent load's outcome.

    sp_help_job is run raw (NOT via INSERT EXEC -- it nests an INSERT EXEC of
    xp_sqlagent_enum_jobs internally, which also lets a non-sysadmin read the
    live current step through ownership chaining). The single wide data row is
    parsed positionally FROM THE END so leading text columns (description/owner)
    can't shift the fields we need:
      [-7]=current_execution_status [-6]=current_execution_step
      [-11]=last_run_outcome [-12]=last_run_time [-13]=last_run_date
    """
    row = _sp_help_job(server, name)
    if not row:
        return ("(no data)", "")

    status, step = row[-7], row[-6]
    if status == '1':                        # Executing -> "Executing Step N (name)" + ETA line
        # Step-anchored progressive ETA (see eta_detail); the label keeps the live
        # step name from sp_help_job, e.g. "Executing Step 4 (Build Chimera)".
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
    """True if the snap's latest run belongs to the CURRENT load â€” i.e. the snap
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


def rce_succeeded_today():
    """True if RAMP 'Aetna RCE 310 ETL Load' (2257) completed Successful today."""
    return _ramp_succeeded_today(RCE_JOBID)


def snap_succeeded_today():
    """True if RAMP 'Aetna RCE 400 Daily Snap' (10053) resolved OK today AND that
    snap ran for the CURRENT load (started after the RCE load completed). A stale
    resolved snap from a prior cycle does NOT count (per user 2026-07-14)."""
    return (_ramp_succeeded_today(SNAP_JOBID, RAMP_OK)
            and snap_is_current(RCE_JOBID, SNAP_JOBID))


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
    # Cross-session dedupe (added 2026-06-26): bail if another session/run already
    # emitted a digest within the last DEDUPE_MINUTES. Claim the slot up front so a
    # near-simultaneous second session bails before doing the slow SQL/curl work.
    force = '--force' in sys.argv
    if not force:
        recent = _recent_emit()
        if recent:
            print(f"NO_POST: deduped (a digest was already emitted at "
                  f"{recent.strftime('%I:%M %p')}, within {DEDUPE_MINUTES} min)")
            return
    # Claim the slot (both normal and --force) so any near-simultaneous run dedupes
    # against this one. Done before the slow SQL/curl work to shrink the race window.
    _claim_slot()

    # Evening extension (per user 2026-07-17): outside the normal 8/12/16 slots the
    # tick calls this with --evening so a load finishing after the last daytime slot
    # still gets its Executing->Successful transition posted. Stay silent unless a
    # load actually ran today (else no-load evenings would post a stale idle line).
    if '--evening' in sys.argv and not force and not _active_today():
        print("NO_POST: evening extension, no load active/completed today")
        return

    # Minimal PLAIN-TEXT format (per user 2026-07-16): the webhook renders only
    # :emoji: -- *bold*/_italic_/`code` show literally + Slack has no text color --
    # so no markup; the only standout is the icon on the ETA line.
    lines = ["AETNA RCE - STATUS UPDATE", ""]
    for server, name, label in SQL_JOBS:
        status_text, detail = sql_job(server, name)
        lines.append(f"{label} {status_text}".rstrip())
        lines.extend(detail)
        lines.append("")
    while lines and lines[-1] == "":
        lines.pop()
    msg = "\n".join(lines)

    # Content dedupe (per user 2026-07-17): post only when the status text CHANGES.
    # This posts the Successful line once when the jobs finish, holds quietly while
    # they stay Successful, then posts again when the next load starts (message
    # flips back to Executing). Replaces the old "both succeeded today -> go
    # silent" skip, which could leave the last post stuck on a stale 'Executing'.
    if not force and msg == _last_msg():
        print("NO_POST: status unchanged since last post")
        return
    _claim_slot(msg)
    print("SLACK|" + msg.replace("\n", "\\n"))


if __name__ == '__main__':
    main()
