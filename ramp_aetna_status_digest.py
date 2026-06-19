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
import json, subprocess
from datetime import datetime

RCE_JOBID = 2257
CHANNEL = 'C09EPLQL2D9'

SQL_JOBS = [
    ("TRGETL2", "ETL_AetnaSupport_MasterLoad"),
    ("TRGETL4", "ETL NCStateAetna MasterLoad"),
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
    if end:
        line = f"Status: *{status}* | started {fmt(start)} | *completed {fmt(end)}*"
    else:
        line = f"Status: *{status}* (running) | started {fmt(start)} | not yet complete"
    return line


def sql_job(server, name):
    dt = ("DATEADD(SECOND,(h.run_time/10000)*3600+((h.run_time%10000)/100)*60+(h.run_time%100),"
          "CONVERT(datetime,CONVERT(char(8),h.run_date)))")
    q = (
        "SET NOCOUNT ON; "
        "SELECT CASE WHEN act.start_execution_date IS NOT NULL AND act.stop_execution_date IS NULL "
        "THEN 'Executing' ELSE 'Idle' END, "
        "CONVERT(varchar(19),act.start_execution_date,120), "
        "h.run_status, CONVERT(varchar(19)," + dt + ",120), "
        "STUFF(STUFF(RIGHT('000000'+CAST(h.run_duration AS varchar),6),5,0,':'),3,0,':') "
        "FROM msdb.dbo.sysjobs j "
        "OUTER APPLY (SELECT TOP 1 start_execution_date, stop_execution_date FROM msdb.dbo.sysjobactivity a "
        "WHERE a.job_id=j.job_id ORDER BY a.session_id DESC) act "
        "OUTER APPLY (SELECT TOP 1 run_status, run_date, run_time, run_duration FROM msdb.dbo.sysjobhistory hh "
        "WHERE hh.job_id=j.job_id AND hh.step_id=0 ORDER BY hh.run_date DESC, hh.run_time DESC) h "
        f"WHERE j.name=N'{name}';"
    )
    out = subprocess.run(['sqlcmd', '-S', server, '-E', '-W', '-s', '|', '-h', '-1', '-Q', q],
                         capture_output=True, text=True, timeout=120)
    for line in out.stdout.splitlines():
        line = line.strip()
        if '|' in line and 'rows affected' not in line:
            parts = line.split('|')
            if len(parts) == 5:
                state, cur_start, run_status, last_run, dur = parts
                try:
                    oc = OUTCOME.get(int(run_status), run_status)
                except ValueError:
                    oc = run_status
                if state == 'Executing':
                    return (f"*{state}* since {fmt(cur_start)} | last outcome {oc} "
                            f"({fmt(last_run)}, {dur})")
                return f"*{state}* | last outcome *{oc}* | {fmt(last_run)} | duration {dur}"
    err = (out.stderr or out.stdout).strip().replace('\n', ' ')[:120]
    return f"(no data{' — ' + err if err else ''})"


def main():
    now = datetime.now().strftime('%m/%d/%Y %I:%M %p')
    lines = [f":bar_chart: *Aetna RCE 310 - Status Update*  ({now})", ""]
    lines.append("*RAMP - Aetna RCE 310 ETL Load*")
    lines.append("- " + rce_status())
    lines.append("")
    lines.append("*SQL Job Activity Monitor*")
    for server, name in SQL_JOBS:
        lines.append(f"- `{name}` ({server}): " + sql_job(server, name))
    msg = "\n".join(lines)
    print("SLACK|" + msg.replace("\n", "\\n"))


if __name__ == '__main__':
    main()
