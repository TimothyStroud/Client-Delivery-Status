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
import json, re, subprocess
from datetime import datetime, timedelta

RCE_JOBID = 2257
CHANNEL = 'C09EPLQL2D9'

# (server, SQL Agent job name, display label). The RCE ETL Load is the SQL job
# 'SSIS AetnaRCE Daily Process' (its steps are the real RCE monitor steps, e.g.
# step 4 = 'Build Chimera') — NOT ETL_AetnaSupport_MasterLoad (AuditSupport).
SQL_JOBS = [
    ("TRGETL2", "SSIS AetnaRCE Daily Process", "Aetna RCE ETL Load"),
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


def rce_status():
    out = subprocess.run(['curl', '-s', '--ntlm', '-u', ':',
                          'http://ramp/api/Ramp/Job/List'],
                         capture_output=True, text=True, timeout=180)
    d = json.loads(out.stdout)['Data']
    jobs = d[0] if (isinstance(d, list) and d and isinstance(d[0], list)) else d
    lr = {}
    for j in jobs:
        if j.get('JobId') == RCE_JOBID:
            lr = j.get('LatestJobRun') or {}
            break
    status = lr.get('Status', '?')
    start = lr.get('StartDate'); end = lr.get('EndDate')
    # Per user: green check for Successful (red X for Failed) instead of the
    # ```diff color trick.
    if end and status == 'Successful':
        return f"- Status: :white_check_mark: *Successful* | started {fmt(start)} | *completed {fmt(end)}*"
    if end and status == 'Failed':
        return f"- Status: :x: *FAILED* | started {fmt(start)} | ended {fmt(end)} - please investigate"
    if end:
        return f"- Status: *{status}* | started {fmt(start)} | *completed {fmt(end)}*"
    return f"- Status: *{status}* (running) | started {fmt(start)} | not yet complete"


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
    q = f"SET NOCOUNT ON; EXEC msdb.dbo.sp_help_job @job_name=N'{name}', @job_aspect=N'JOB';"
    out = subprocess.run(['sqlcmd', '-S', server, '-E', '-W', '-s', '~', '-Q', q],
                         capture_output=True, text=True, timeout=120)
    row = None
    for line in out.stdout.splitlines():
        line = line.rstrip()
        if not line or line.startswith('job_id') or 'rows affected' in line:
            continue
        if set(line) <= set('-~'):          # the ---- separator row
            continue
        parts = line.split('~')
        if len(parts) >= 32:
            row = [p.strip() for p in parts]
            break
    if not row:
        err = (out.stderr or out.stdout).strip().replace('\n', ' ')[:120]
        return f"(no data{' - ' + err if err else ''})"

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
    return f"*{st}* | last run {oc} ({fmt_dt(row[-13], row[-12])})"


def main():
    now = datetime.now().strftime('%m/%d/%Y %I:%M %p')
    lines = [f"<!here> :bar_chart: *Aetna RCE 310 - Status Update*  ({now})", ""]
    lines.append("*RAMP - Aetna RCE 310 ETL Load*")
    lines.append(rce_status())
    lines.append("")
    lines.append("*SQL Job Activity Monitor*")
    for server, name, label in SQL_JOBS:
        lines.append(f"- `{label}` ({server}): " + sql_job(server, name))
    msg = "\n".join(lines)
    print("SLACK|" + msg.replace("\n", "\\n"))


if __name__ == '__main__':
    main()
