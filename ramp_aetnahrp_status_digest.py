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
import base64, json, os, re, subprocess, sys
from datetime import datetime, timedelta

HRP_JOBID = 1246
SNAP_JOBID = 1247      # RAMP 'Aetna 0120 HRP Snap' (added 2026-07-14 per user)
CHANNEL = 'C09EPLQL2D9'

# Claim files staged by RAMP 'Aetna 0100 HRP Stage' (added 2026-07-14 per user).
# This Windows python can't reach \\etl2 directly (see cb_scan.py) but PowerShell
# can, so the file list is gathered via a powershell -EncodedCommand shell-out.
# Files pending load sit in the Claim root/Load/ToLoad; loaded files -> Loaded.
CLAIM_BASE = r'\\etl2\CLIENTS\AetnaHRP\Data\Claim'
PS_CLAIM = r"""
$b='\\etl2\CLIENTS\AetnaHRP\Data\Claim'
$dirs = @{ ''='pending'; 'Load'='pending'; 'ToLoad'='pending'; 'Loaded'='loaded' }
foreach($k in $dirs.Keys){
  $p = if($k){ Join-Path $b $k } else { $b }
  if(Test-Path -LiteralPath $p){
    Get-ChildItem -LiteralPath $p -Filter 'VENDOR.CB-CLAIMS-EXTRACT*.csv' -File -ErrorAction SilentlyContinue |
      ForEach-Object { "$($dirs[$k])|$($_.Name)" }
  }
}
"""

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


def ramp_line(jobid):
    """Status body (no leading '- ') for a RAMP job's LatestJobRun. Green check
    for Successful/Resolved, red X for Failed."""
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
        return "(no data)"

    status, step = row[-7], row[-6]
    if status == '1':                        # Executing -> current step + ETA
        line = f"*Executing*: {step}"
        m = re.match(r'\s*(\d+)', step)      # leading step number
        if m:
            secs = remaining_secs(server, name, int(m.group(1)))
            if secs and secs > 0:
                line += f" | ETA ~{_clock(datetime.now() + timedelta(seconds=secs))}"
        return line
    st = EXEC_STATUS.get(status, f'State {status}')
    oc = RUN_OUTCOME.get(row[-11], row[-11])
    # Checkmark for a Succeeded last run, red X for Failed (per user 2026-07-14).
    icon = ':white_check_mark: ' if oc == 'Succeeded' else (':x: ' if oc == 'Failed' else '')
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


def claim_files(limit=6):
    """CB-CLAIMS-EXTRACT files staged by 'Aetna 0100 HRP Stage', via a PowerShell
    shell-out (this python can't reach \\etl2, PowerShell can). Returns
    [(name, state, dt), ...] newest-first, capped. 'pending' = still to load,
    'loaded' = moved to the Loaded dir."""
    try:
        enc = base64.b64encode(PS_CLAIM.encode('utf-16-le')).decode('ascii')
        out = subprocess.run(['powershell', '-NoProfile', '-NonInteractive',
                              '-EncodedCommand', enc],
                             capture_output=True, text=True, timeout=90)
    except Exception:
        return None
    if out.returncode != 0 and not out.stdout.strip():
        return None
    seen = {}
    for line in out.stdout.splitlines():
        line = line.strip()
        if '|' not in line:
            continue
        state, name = line.split('|', 1)
        name = name.strip()
        if not name:
            continue
        if name not in seen or state == 'loaded':   # a name lives in one dir; 'loaded' wins
            seen[name] = state
    rows = [(n, s, _parse_extract_dt(n)) for n, s in seen.items()]
    rows.sort(key=lambda r: (r[2] or datetime.min), reverse=True)
    return rows[:limit]


def claim_file_lines():
    """Slack lines for the claim-files section (checkmark when loaded)."""
    rows = claim_files()
    if rows is None:
        return ["- (file share unavailable)"]
    if not rows:
        return ["- (no CB-CLAIMS-EXTRACT files found)"]
    out = []
    for name, state, dt in rows:
        icon = ':white_check_mark:' if state == 'loaded' else ':hourglass_flowing_sand:'
        tag = 'loaded' if state == 'loaded' else 'pending load'
        dstr = dt.strftime('%m/%d/%Y') if dt else '?'
        out.append(f"- {icon} `{name}` ({tag}, {dstr})")
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
    lines = [f"<!here> :bar_chart: *Aetna HRP - Status Update*  ({now})", ""]
    lines.append("*RAMP*")
    lines.append("- `Aetna 0110 HRP Load`: " + ramp_line(HRP_JOBID))
    lines.append("- `Aetna 0120 HRP Snap`: " + snap_line(HRP_JOBID, SNAP_JOBID))
    lines.append("")
    lines.append("*SQL Job Activity Monitor*")
    for server, name, label in SQL_JOBS:
        lines.append(f"- `{label}` ({server}): " + sql_job(server, name))
    lines.append("")
    lines.append("*Claim Files - Aetna 0100 HRP Stage*")
    lines.extend(claim_file_lines())
    msg = "\n".join(lines)
    print("SLACK|" + msg.replace("\n", "\\n"))


if __name__ == '__main__':
    main()
