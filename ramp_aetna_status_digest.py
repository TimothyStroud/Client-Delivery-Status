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

RCE_JOBID = 2257
SNAP_JOBID = 10053     # RAMP 'Aetna RCE 400 Daily Snap' (added 2026-07-14 per user)
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

# Unified digest list — ONE flat, ORDERED set of items (order per user 2026-07-14):
#   1. Aetna RCE Daily Process      (SQL 'SSIS AetnaRCE Daily Process', TRGETL2)
#   2. ETL NCStateAetna MasterLoad  (SQL, TRGETL4)
#   3. Aetna RCE 400 Daily Snap     (RAMP job 10053)
#   4. Aetna RCE Support MasterLoad (SQL 'ETL_AetnaSupport_MasterLoad', TRGETL2)
# The old RAMP 'Aetna RCE 310 ETL Load' (2257) line was dropped as redundant with
# the 'Aetna RCE Daily Process' SQL job (same run; matching start ~12:58am). 2257
# is still used by rce_succeeded_today() for the both-done skip.
# Each item: ('sql', server, jobname, label) or ('ramp', jobid, None, label).
DIGEST_ITEMS = [
    ('sql',  'TRGETL2', 'SSIS AetnaRCE Daily Process', 'Aetna RCE Daily Process'),
    ('sql',  'TRGETL4', 'ETL NCStateAetna MasterLoad', 'ETL NCStateAetna MasterLoad'),
    ('ramp', SNAP_JOBID, None,                          'Aetna RCE 400 Daily Snap'),
    ('sql',  'TRGETL2', 'ETL_AetnaSupport_MasterLoad', 'Aetna RCE Support MasterLoad'),
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
    """Return the LatestJobRun dict for a RAMP job id (or {} if not found)."""
    for j in _all_jobs():
        if j.get('JobId') == jobid:
            return j.get('LatestJobRun') or {}
    return {}


def rce_run():
    """LatestJobRun for RAMP 'Aetna RCE 310 ETL Load' (2257)."""
    return job_run(RCE_JOBID)


# RAMP terminal-good statuses: mining/load jobs report 'Successful'; snap jobs
# report 'Resolved' (a full snap run resolves its queue item — e.g. the RCE 330
# Weekly Snap sits at 'Resolved' after a ~1h run). Both get a green check.
RAMP_OK = ('Successful', 'Resolved')


def ramp_line(jobid):
    """Status body (no leading '- ') for a RAMP job's LatestJobRun, for the unified
    digest list. Green check for Successful/Resolved (red X for Failed)."""
    lr = job_run(jobid)
    status = lr.get('Status', '?')
    start = lr.get('StartDate'); end = lr.get('EndDate')
    if end and status in RAMP_OK:
        return f":white_check_mark: *{status}* | started {fmt(start)} | *completed {fmt(end)}*"
    if end and status == 'Failed':
        return f":x: *FAILED* | started {fmt(start)} | ended {fmt(end)} - please investigate"
    if end:
        return f"*{status}* | started {fmt(start)} | *completed {fmt(end)}*"
    if not start:
        return f"*{status}* (queued) | not yet started"
    return f"*{status}* (running) | started {fmt(start)} | not yet complete"


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


def remaining_secs(server, name, cur_step):
    """Estimated seconds left = sum of avg historical durations of steps
    cur_step..end (avg over the last 8 successful runs per step). Counts the
    current step in full (slight overestimate) and skips step 1's variable
    upstream 'Wait' (it logs ~0 duration). Returns int seconds or None."""
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
        return "(no data)"

    status, step = row[-7], row[-6]
    if status == '1':                        # Executing -> current step + ETA
        line = f"*Executing*: {step}"
        m = re.match(r'\s*(\d+)', step)      # leading step number, e.g. "4 (Build Chimera)"
        if m:
            secs = remaining_secs(server, name, int(m.group(1)))
            if secs and secs > 0:
                line += f" | ETA ~{_clock(datetime.now() + timedelta(seconds=secs))}"
        return line
    st = EXEC_STATUS.get(status, f'State {status}')
    oc = RUN_OUTCOME.get(row[-11], row[-11])
    # Did the most recent run finish TODAY?
    try:
        d = int(row[-13])
        t = datetime.now()
        ran_today = d == t.year * 10000 + t.month * 100 + t.day
    except (ValueError, TypeError):
        ran_today = False
    # Icon rules (per user 2026-07-14, 2026-07-15):
    #   - Failure always wins -> red X.
    #   - Today's run already Succeeded -> keep the green checkmark (done for the day).
    #   - Otherwise an Idle process shows the hourglass (between runs, not yet done).
    if oc == 'Failed':
        icon = ':x: '
    elif oc == 'Succeeded' and ran_today:
        icon = ':white_check_mark: '
    elif status == '4':                       # Idle, no success yet today -> hourglass
        icon = ':hourglass_flowing_sand: '
    else:
        icon = ''
    return f"{icon}*{st}* | last run {oc} ({fmt_dt(row[-13], row[-12])})"


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
        return (":hourglass_flowing_sand: *Waiting to snap current load* "
                f"(last snap {status} {fmt(end)}, ran before this load completed)")
    if end and status in RAMP_OK:
        return f":white_check_mark: *{status}* | started {fmt(start)} | *completed {fmt(end)}*"
    if end and status == 'Failed':
        return f":x: *FAILED* | started {fmt(start)} | ended {fmt(end)} - please investigate"
    if end:
        return f"*{status}* | started {fmt(start)} | *completed {fmt(end)}*"
    if not start:
        return f"*{status}* (queued) | not yet started"
    return f"*{status}* (running) | started {fmt(start)} | not yet complete"


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

    # Per user 2026-06-23: once BOTH the RCE load and NCStateAetna MasterLoad
    # have already SUCCEEDED today, the rest of the day's digests are redundant
    # -> emit no SLACK line so the cron/task posts nothing. (If either is still
    # running, failed, or hasn't run, the digest still posts.)
    if (rce_succeeded_today()
            and snap_succeeded_today()
            and job_succeeded_today('TRGETL4', 'ETL NCStateAetna MasterLoad')):
        print('NO_POST: RCE + Snap + NCStateAetna all Succeeded today')
        return
    now = datetime.now().strftime('%m/%d/%Y %I:%M %p')
    lines = [f"<!here> :bar_chart: *Aetna RCE 310 - Status Update*  ({now})", ""]
    for item in DIGEST_ITEMS:
        if item[0] == 'ramp':
            _, jobid, _, label = item
            body = snap_line(RCE_JOBID, jobid) if jobid == SNAP_JOBID else ramp_line(jobid)
            lines.append(f"- `{label}` (RAMP): " + body)
        else:
            _, server, name, label = item
            lines.append(f"- `{label}` ({server}): " + sql_job(server, name))
    msg = "\n".join(lines)
    print("SLACK|" + msg.replace("\n", "\\n"))


if __name__ == '__main__':
    main()
