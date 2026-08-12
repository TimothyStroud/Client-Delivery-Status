"""
EverNorthRx Deployment & Loading monitor (headless, zero Claude tokens).

Watches RAMP for the three jobs the delivery/deployment crew cares about and
posts a Slack line the moment each one STARTS and the moment each one FINISHES:

  EvernorthRx Masterload 0110 Load   JobId 10711   (NOT the 0100 Stage)
  EvernorthRx Masterload 0120 Snap   JobId 10717
  EvernorthRx COBC 0110 Load         JobId 10709   (NOT the 0100 Stage)

Cadence (per user 2026-08-12):
  - every 15 min while waiting for / running the Masterload 0110 Load
  - every  5 min once the Masterload Load has finished, i.e. while the Snap and
    the COBC Load are queued/running behind it
  - checks start at 7:00am on WEEKDAYS and stop for the day once all three jobs
    have finished (and nothing is still in flight); then idle until the next
    weekday 7am.

The Windows Scheduled Task fires this every 5 minutes; the 15-minute cadence and
the "day is done" stop are enforced HERE, so the RAMP query only runs when it is
actually due. Slack posts happen only on start/finish transitions -- the polling
itself is silent.

Delivery: the destination is the "EverNorthRx Deployment & Loading" GROUP DM
(C0BNZM47T9V), which a Slack Workflow Builder webhook CANNOT reach -- webhooks
only post to channels. So this task does all the RAMP polling headlessly and
writes each line to a delivery queue; a Claude cron drains the queue and posts
it via the Slack plugin (--drain / --ack, two-phase so a failed post retries).
If a channel + webhook is ever set up, drop the URL into H:\slack_wf_evernorthrx_deploy.txt
and delivery switches to the webhook automatically -- no other change needed.

State: H:\evernorthrx_deploy_monitor_state.json
Queue: H:\evernorthrx_deploy_pending.json
Log:   H:\evernorthrx_deploy_monitor.log

Modes:
  (default)     run a due check, deliver transitions, update state
  --dry-run     run the check regardless of cadence; print, never deliver/save
  --force       ignore the cadence gate / day-complete flag (still delivers)
  --status      print current RAMP state for the three jobs and exit
  --drain       print pending lines as POST|<id>|<text> (or NONE) for the cron
  --ack <ids>   remove lines that have actually posted
"""
import sys, os, re, json, subprocess, urllib.request
from datetime import datetime, timedelta

SQL_SERVER = "TRGUTIL10"

JOB_MASTER_LOAD = 10711   # EvernorthRx Masterload 0110 Load
JOB_MASTER_SNAP = 10717   # EvernorthRx Masterload 0120 Snap
JOB_COBC_LOAD   = 10709   # EvernorthRx COBC 0110 Load

# Display order = execution order.
JOBS = [
    (JOB_MASTER_LOAD, "EverNorthRx Masterload 0110 Load"),
    (JOB_MASTER_SNAP, "EverNorthRx Masterload 0120 Snap"),
    (JOB_COBC_LOAD,   "EverNorthRx COBC 0110 Load"),
]
JOB_NAME = dict(JOBS)
JOB_IDS = [j for j, _ in JOBS]

URL_FILE   = r"H:\slack_wf_evernorthrx_deploy.txt"
STATE_FILE = r"H:\evernorthrx_deploy_monitor_state.json"
QUEUE_FILE = r"H:\evernorthrx_deploy_pending.json"
LOG_FILE   = r"H:\evernorthrx_deploy_monitor.log"
SEP = "\x1f"

START_HOUR = 7            # first check of the day, weekdays only
SLOW_MINUTES = 15         # waiting for / running the Masterload Load
FAST_MINUTES = 5          # Snap + COBC queued behind a finished Masterload Load
CADENCE_SLACK_SECONDS = 60  # tolerate the task firing a few seconds early

DRY   = "--dry-run" in sys.argv
FORCE = "--force" in sys.argv


def log(msg):
    line = f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}"
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass
    print(line)


def run_sql(query, db="RAMP", server=None):
    r = subprocess.run(
        ["sqlcmd", "-S", server or SQL_SERVER, "-d", db, "-E", "-Q", query,
         "-W", "-s", SEP, "-h", "-1"],
        capture_output=True, text=True, check=False, timeout=120,
    )
    out = []
    for line in r.stdout.splitlines():
        s = line.rstrip("\n")
        if not s or s.startswith("---") or "rows affected" in s:
            continue
        out.append(s.split(SEP))
    return out


def parse_dt(s):
    s = (s or "").strip()
    if not s or s.upper() == "NULL":
        return None
    try:
        return datetime.strptime(s[:19], "%Y-%m-%d %H:%M:%S")
    except ValueError:
        return None


def classify(status):
    """RAMP Queue.Status -> SUCCESS / FAILED / PENDING.

    Observed values for this feed: Ready (queued or running), Successful,
    Resolved (manually cleared / no-op), Failed.
    """
    s = (status or "").strip().lower()
    if s.startswith("success") or s == "resolved":
        return "SUCCESS"
    if s == "failed" or s.startswith("fail"):
        return "FAILED"
    return "PENDING"


def fmt_dur(start, end):
    if not (start and end):
        return ""
    secs = int((end - start).total_seconds())
    if secs < 0:
        return ""
    h, m = divmod(secs // 60, 60)
    if h and m:
        return f"{h}h {m}m"
    if h:
        return f"{h}h"
    return f"{m}m" if m else "<1m"


def load_state():
    try:
        with open(STATE_FILE, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def save_state(st):
    if DRY:
        return
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(st, f, indent=2)


def read_url():
    try:
        return open(URL_FILE, encoding="utf-8").read().strip() or None
    except Exception:
        return None


def sanitize(text):
    """The Workflow Builder webhook renders :emoji: shortcodes ONLY -- *bold*,
    `code` and <!here> come through literally. Keep underscores (they live
    inside emoji shortcodes)."""
    text = text.replace("<!here> ", "").replace("<!here>", "")
    text = re.sub(r"(?m)^> ?", "", text)
    return text.replace("*", "").replace("`", "")


def load_queue():
    try:
        with open(QUEUE_FILE, encoding="utf-8") as f:
            q = json.load(f)
        return q if isinstance(q, list) else []
    except (OSError, ValueError):
        return []


def save_queue(q):
    with open(QUEUE_FILE, "w", encoding="utf-8") as f:
        json.dump(q, f, indent=2)


def enqueue(text):
    """Hand a line off to the Slack delivery queue.

    The destination is a GROUP DM, which a Workflow Builder webhook cannot
    reach, so this task cannot post by itself. It records the line here and a
    Claude cron drains the queue (--drain / --ack). The queue file is the
    durable record: a line is only removed once it has actually posted.
    """
    q = load_queue()
    entry = {
        "id": f"{datetime.now():%Y%m%d%H%M%S}-{len(q)}",
        "queued": f"{datetime.now():%Y-%m-%d %H:%M:%S}",
        "text": text,
    }
    q.append(entry)
    save_queue(q)
    log(f"  QUEUED [{entry['id']}]: {text}")
    return True


def drain():
    """Print pending lines for the Claude cron, oldest first."""
    q = load_queue()
    if not q:
        print("NONE")
        return
    for e in q:
        print(f"POST|{e['id']}|{e['text']}")


def ack(ids):
    """Drop lines that have actually posted."""
    ids = set(ids)
    q = load_queue()
    keep = [e for e in q if e["id"] not in ids]
    save_queue(keep)
    log(f"acked {len(q) - len(keep)} queued line(s); {len(keep)} still pending")
    print(f"ACKED {len(q) - len(keep)}; PENDING {len(keep)}")


def post(text):
    """Deliver one line: webhook if one is configured, else the Slack queue."""
    text = sanitize(text)
    if DRY:
        log(f"  [dry-run] would deliver: {text}")
        return True
    url = read_url()
    if not url:
        return enqueue(text)
    data = json.dumps({"Text": text}).encode("utf-8")
    req = urllib.request.Request(url, data=data,
                                 headers={"Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            body = r.read().decode("utf-8", "replace").strip()
        log(f"  POSTED (HTTP {r.status} {body[:40]}): {text}")
        return True
    except Exception as e:
        log(f"  ERROR posting: {e} -- text: {text}")
        return False


# --------------------------------------------------------------------------- #
# RAMP
# --------------------------------------------------------------------------- #
def todays_runs(today):
    """Every run of the three jobs queued or started today, oldest first.

    Keyed off COALESCE(StartDate, CreateDate) so a run that is queued but has
    not begun yet is still visible (that is how we know the Snap/COBC have
    lined up behind the Masterload).
    """
    rows = run_sql(
        "SET NOCOUNT ON; SELECT QueueId, JobId, ISNULL(Status,''), "
        "CONVERT(varchar(19),StartDate,121), CONVERT(varchar(19),EndDate,121), "
        "CONVERT(varchar(19),CreateDate,121) "
        f"FROM [RAMP].[ramp].[Queue] WHERE JobId IN ({','.join(map(str, JOB_IDS))}) "
        f"AND CAST(COALESCE(StartDate, CreateDate) AS date) = '{today}' "
        "ORDER BY QueueId ASC"
    )
    runs = []
    for r in rows:
        if len(r) < 6:
            continue
        start, end = parse_dt(r[3]), parse_dt(r[4])
        # A 'Resolved' entry whose StartDate == EndDate is a queue row that was
        # cleared (dependency satisfied / superseded), not a run that did work.
        # e.g. 8/12 QueueId 1421742: Resolved, 11:42:22 -> 11:42:22.
        if r[2].strip().lower() == "resolved" and start and end and start == end:
            continue
        runs.append({
            "qid": int(r[0]),
            "jobid": int(r[1]),
            "status": r[2].strip(),
            "start": start,
            "end": end,
            "create": parse_dt(r[5]),
        })
    return runs


def run_by_qid(qid):
    rows = run_sql(
        "SET NOCOUNT ON; SELECT QueueId, JobId, ISNULL(Status,''), "
        "CONVERT(varchar(19),StartDate,121), CONVERT(varchar(19),EndDate,121) "
        f"FROM [RAMP].[ramp].[Queue] WHERE QueueId={int(qid)}"
    )
    if not rows or len(rows[0]) < 5:
        return None
    r = rows[0]
    return {"qid": int(r[0]), "jobid": int(r[1]), "status": r[2].strip(),
            "start": parse_dt(r[3]), "end": parse_dt(r[4])}


# --------------------------------------------------------------------------- #
# What is actually doing the work behind each RAMP job
# --------------------------------------------------------------------------- #
# Parsed out of the RAMP JobXml (config.Job.JobXML) 2026-08-12. A RAMP job is
# mostly a SqlAgentKickOff/SqlAgentMonitor wrapper -- the real time is spent in
# the SQL Agent job it kicks off and waits on, so that is where any mid-run
# signal has to come from.
#
#   0110 Load : KickOff+Monitor 'ETL EvernorthRx MasterLoad'      @ETL4
#   0120 Snap : Monitor of THREE jobs, then the snap itself
#   COBC Load : KickOff+Monitor 'ETL EvernorthRx COBC MasterLoad' @ETL4
AGENT_JOBS = {
    JOB_MASTER_LOAD: [("ETL4", "ETL EvernorthRx MasterLoad")],
    JOB_COBC_LOAD:   [("ETL4", "ETL EvernorthRx COBC MasterLoad")],
    JOB_MASTER_SNAP: [("TRGETLPROD5", "Rx EverNorthRx_Mine - Run Etl.PostSnapProcess"),
                      ("ETL4", "ETL EvernorthRx MasterLoad"),
                      ("ETL4", "ETL EvernorthRx COBC MasterLoad")],
}

# Short display names -- the full agent job names are far too long for a Slack line.
AGENT_LABEL = {
    "Rx EverNorthRx_Mine - Run Etl.PostSnapProcess": "PostSnapProcess",
    "ETL EvernorthRx MasterLoad": "MasterLoad",
    "ETL EvernorthRx COBC MasterLoad": "COBC MasterLoad",
}

# Which agent jobs are SSIS-backed, and where their SSISDB lives.
#
# NOTE (2026-08-12): the 0120 Snap has **no SSIS on its critical path**. Its
# blocking job, 'Rx EverNorthRx_Mine - Run Etl.PostSnapProcess', is a SINGLE
# T-SQL step (`exec Etl.PostSnapProcess`) on TRGETLPROD5 -- there is no package,
# no task log, and therefore no SSISDB ladder to build. The Snap is paced off
# that job's own msdb duration history instead. The SSISDB ladder below applies
# to the two LOAD jobs, which is where the SSIS packages actually are.
SSIS_PKG = {
    "ETL EvernorthRx MasterLoad":      ("ETL4", "EvernorthRx_MasterLoad.dtsx"),
    "ETL EvernorthRx COBC MasterLoad": ("ETL4", "EvernorthRx_MasterLoad_COBC.dtsx"),
}

HISTORY_N = 20      # recent successful runs used to pace a job
STEP_PCT = 80       # survival percentile (overruns are one-sided, so p50 runs early)

# SSIS milestone ladder admissibility -- same shape as the Aetna digests.
LADDER_MIN_SAMPLES = 5
LADDER_MIN_FRAC = 0.55
LADDER_MAX_FRAC = 0.97
LADDER_IQR_FLOOR = 20 * 60
LADDER_IQR_REL = 0.35
LADDER_IQR_CAP = 90 * 60


def _pct(vals, p):
    if not vals:
        return None
    v = sorted(vals)
    k = (len(v) - 1) * p / 100.0
    lo = int(k)
    hi = min(lo + 1, len(v) - 1)
    return v[lo] + (v[hi] - v[lo]) * (k - lo)


def _dur_txt(secs):
    if secs is None:
        return "?"
    secs = int(secs)
    h, m = divmod(secs // 60, 60)
    if h and m:
        return f"{h}h {m}m"
    if h:
        return f"{h}h"
    return f"{m}m" if m else "<1m"


# msdb stores run_date/run_time/run_duration as ints; dbo.agent_datetime is NOT
# executable by this login (permission denied on TRGETLPROD5), so decode by hand.
_MSDB_START = ("DATEADD(second,(h.run_time%100)+((h.run_time/100)%100)*60"
               "+(h.run_time/10000)*3600, CONVERT(datetime,CONVERT(char(8),h.run_date),112))")
_MSDB_DUR = "(h.run_duration%100)+((h.run_duration/100)%100)*60+(h.run_duration/10000)*3600"


def agent_live(server, job):
    """(start_dt, step_name) if that SQL Agent job is running right now, else None."""
    try:
        rows = run_sql(
            "SET NOCOUNT ON; SELECT TOP 1 CONVERT(varchar(19),ja.start_execution_date,121), "
            "ISNULL(js.step_name,'') FROM sysjobactivity ja "
            "JOIN sysjobs j ON j.job_id=ja.job_id "
            "LEFT JOIN sysjobsteps js ON js.job_id=ja.job_id "
            "  AND js.step_id=ja.last_executed_step_id "
            f"WHERE j.name=N'{job.replace(chr(39), chr(39) * 2)}' "
            "AND ja.start_execution_date IS NOT NULL AND ja.stop_execution_date IS NULL "
            "AND ja.session_id=(SELECT MAX(session_id) FROM syssessions) "
            "ORDER BY ja.start_execution_date DESC",
            db="msdb", server=server)
    except Exception:
        return None
    if not rows or len(rows[0]) < 2:
        return None
    st = parse_dt(rows[0][0])
    return (st, rows[0][1].strip()) if st else None


def agent_history(server, job, n=HISTORY_N):
    """Durations (seconds) of the last n SUCCESSFUL runs of that agent job."""
    try:
        rows = run_sql(
            f"SET NOCOUNT ON; SELECT TOP {int(n)} {_MSDB_DUR} FROM sysjobhistory h "
            "JOIN sysjobs j ON j.job_id=h.job_id "
            f"WHERE j.name=N'{job.replace(chr(39), chr(39) * 2)}' AND h.step_id=0 "
            "AND h.run_status=1 ORDER BY h.instance_id DESC",
            db="msdb", server=server)
    except Exception:
        return []
    out = []
    for r in rows:
        if r and r[0].strip().lstrip("-").isdigit():
            out.append(int(r[0].strip()))
    return out


def agent_last_run(server, job, since=None):
    """(start_dt, end_dt, ok) of the most recent completed run, optionally the
    most recent one that STARTED at/after `since`. None if there isn't one."""
    where = ""
    if since:
        where = f" AND {_MSDB_START} >= '{since:%Y-%m-%d %H:%M:%S}'"
    try:
        rows = run_sql(
            f"SET NOCOUNT ON; SELECT TOP 1 CONVERT(varchar(19),{_MSDB_START},121), "
            f"{_MSDB_DUR}, h.run_status FROM sysjobhistory h "
            "JOIN sysjobs j ON j.job_id=h.job_id "
            f"WHERE j.name=N'{job.replace(chr(39), chr(39) * 2)}' AND h.step_id=0{where} "
            "ORDER BY h.instance_id DESC",
            db="msdb", server=server)
    except Exception:
        return None
    if not rows or len(rows[0]) < 3:
        return None
    st = parse_dt(rows[0][0])
    if not st or not rows[0][1].strip().lstrip("-").isdigit():
        return None
    dur = int(rows[0][1].strip())
    return st, st + timedelta(seconds=dur), rows[0][2].strip() == "1"


# ---- SSIS milestone ladder (the two LOAD jobs only -- the Snap has no SSIS) --
def ssis_ladder(server, pkg):
    """{executable_id: median_tail_seconds} over the last 8 successful runs.

    A rung qualifies when it has real history, completes inside
    [LADDER_MIN_FRAC, LADDER_MAX_FRAC] of package wall-clock, and its TAIL
    (run_end - task_end) has a bounded interquartile spread -- the ETA is
    rung_end + median_tail, so the error IS the tail's spread. Chosen
    dynamically so it self-heals across package redeploys (a redeploy renumbers
    executable_ids; unknown ids just fail the sample test until they build up).
    """
    try:
        rows = run_sql(
            f"SET NOCOUNT ON; DECLARE @pkg sysname=N'{pkg}'; "
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
            f"                  THEN {LADDER_IQR_FLOOR} ELSE {LADDER_IQR_REL}*med_tail END);",
            db="SSISDB", server=server)
    except Exception:
        return {}
    ladder = {}
    for r in rows:
        if len(r) >= 2 and r[0].strip().lstrip("-").isdigit() and r[1].strip().lstrip("-").isdigit():
            ladder[int(r[0].strip())] = int(r[1].strip())
    return ladder


def ssis_eta(server, pkg):
    """(eta_dt, rung_note) projected from the LATEST ladder rung the live
    execution has passed: rung_end + that rung's median tail. (None, '') when the
    package isn't running or hasn't cleared a rung yet -- which is exactly the
    'no signal finer than whole-job history' case, so the caller falls back.

    Later rungs have shorter, tighter tails, so taking the latest one passed
    tightens the estimate progressively instead of staying blind until one
    late task fires."""
    ladder = ssis_ladder(server, pkg)
    if not ladder:
        return None, ""
    ids = ",".join(str(int(k)) for k in ladder)
    try:
        # Pick the latest genuinely-current execution FIRST (recency guard drops
        # orphans stuck at status=2), THEN ask which rungs it has passed -- doing
        # it the other way round can silently match an older execution.
        rows = run_sql(
            "SET NOCOUNT ON; "
            f"DECLARE @e bigint=(SELECT TOP 1 execution_id FROM catalog.executions "
            f"  WHERE package_name=N'{pkg}' AND status=2 "
            "   AND start_time>=DATEADD(hour,-36,SYSDATETIMEOFFSET()) "
            "  ORDER BY execution_id DESC); "
            "IF @e IS NULL SELECT '',''; ELSE "
            "SELECT TOP 1 CONVERT(varchar(19),es.end_time,121), "
            "  CAST(es.executable_id AS varchar(30)) "
            "FROM catalog.executable_statistics es "
            f" WHERE es.execution_id=@e AND es.executable_id IN ({ids}) "
            "ORDER BY es.end_time DESC",
            db="SSISDB", server=server)
    except Exception:
        return None, ""
    if not rows or len(rows[0]) < 2:
        return None, ""
    rung_end = parse_dt(rows[0][0])
    rid = rows[0][1].strip()
    if not rung_end or not rid.lstrip("-").isdigit():
        return None, ""
    tail = ladder.get(int(rid))
    if tail is None:
        return None, ""
    return rung_end + timedelta(seconds=tail), f"SSIS milestone +{_dur_txt(tail)}"


# --------------------------------------------------------------------------- #
# Mid-run detail attached to the START / FINISH lines
# --------------------------------------------------------------------------- #
# Deliberately only ever attached to a start or finish message -- per the brief,
# the monitor posts ONLY when a job starts and finishes, so an ETA rides along
# with the start line rather than becoming its own progress chatter.
def _eta_line(eta, note=""):
    if not eta:
        return ""
    return f" - ETA ~{eta:%H:%M}" + (f" ({note})" if note else "")


def start_detail(run):
    """One extra line for the STARTED message: what the job is really waiting on
    and when it is likely to finish. Returns '' if nothing useful is known --
    every lookup here is best-effort and must never break the alert."""
    try:
        jobid = run["jobid"]
        agents = AGENT_JOBS.get(jobid, [])

        if jobid == JOB_MASTER_SNAP:
            # The Snap is a SqlAgentMonitor gate over three agent jobs. Report
            # whichever one is actually running -- that is the true blocker.
            for server, job in agents:
                live = agent_live(server, job)
                if not live:
                    continue
                st, step = live
                elapsed = (datetime.now() - st).total_seconds()
                hist = agent_history(server, job)
                typical = _pct(hist, STEP_PCT)
                label = AGENT_LABEL.get(job, job)
                bits = [f":hourglass_flowing_sand: waiting on {label} "
                        f"(running {_dur_txt(elapsed)}"
                        + (f", typically {_dur_txt(typical)}" if typical else "") + ")"]
                if typical:
                    bits.append(_eta_line(st + timedelta(seconds=typical)).lstrip(" -").strip())
                return "   " + " - ".join(b for b in bits if b)
            # Nothing running: pace off the Snap's own RAMP history.
            hist = ramp_history(jobid)
            typical = _pct(hist, STEP_PCT)
            if typical:
                return (f"   :hourglass_flowing_sand: typically {_dur_txt(typical)}"
                        + _eta_line(run["start"] + timedelta(seconds=typical)))
            return ""

        # The two LOAD jobs: SSIS is the whole story, so use the milestone ladder
        # when the package has cleared a rung, else the package's own history.
        for server, job in agents:
            pkg = SSIS_PKG.get(job)
            if not pkg:
                continue
            eta, note = ssis_eta(pkg[0], pkg[1])
            if eta:
                return f"   :hourglass_flowing_sand: SSIS {pkg[1]} running{_eta_line(eta, note)}"
            hist = agent_history(server, job)
            typical = _pct(hist, STEP_PCT)
            if typical:
                live = agent_live(server, job)
                anchor = live[0] if live else run["start"]
                return (f"   :hourglass_flowing_sand: {AGENT_LABEL.get(job, job)} typically "
                        f"{_dur_txt(typical)}" + _eta_line(anchor + timedelta(seconds=typical)))
        return ""
    except Exception as e:
        log(f"  (start_detail unavailable: {e})")
        return ""


def finish_detail(run):
    """One extra line for the FINISHED message explaining where the time went.

    This matters more than it looks: on 8/12 the RAMP 0110 Load queue entry
    showed 11:44-11:49 (5m) while the SSIS MasterLoad package had actually run
    07:47-11:00 (3h 13m). Reporting only the RAMP span understates the load by
    hours -- exactly the discrepancy Nancy queried on 8/11."""
    try:
        jobid = run["jobid"]
        for server, job in AGENT_JOBS.get(jobid, []):
            if jobid == JOB_MASTER_SNAP and job not in (
                    "Rx EverNorthRx_Mine - Run Etl.PostSnapProcess",):
                continue        # only the Snap's real blocker is worth reporting
            # A run of the agent job overlapping this RAMP run's window.
            window = (run["start"] or run["end"]) - timedelta(hours=12)
            last = agent_last_run(server, job, since=window)
            if not last:
                continue
            st, en, ok = last
            if run["end"] and st > run["end"]:
                continue
            label = AGENT_LABEL.get(job, job)
            mark = "" if ok else " (FAILED)"
            return (f"   :information_source: {label} ran {st:%H:%M}-{en:%H:%M} "
                    f"({_dur_txt((en - st).total_seconds())}){mark}")
        return ""
    except Exception as e:
        log(f"  (finish_detail unavailable: {e})")
        return ""


def ramp_history(jobid, n=HISTORY_N):
    """Durations (seconds) of recent successful RAMP runs of a job."""
    try:
        rows = run_sql(
            f"SET NOCOUNT ON; SELECT TOP {int(n)} DATEDIFF(second,StartDate,EndDate) "
            f"FROM [RAMP].[ramp].[Queue] WHERE JobId={int(jobid)} AND EndDate IS NOT NULL "
            "AND Status='Successful' AND DATEDIFF(second,StartDate,EndDate) > 0 "
            "ORDER BY QueueId DESC")
    except Exception:
        return []
    return [int(r[0].strip()) for r in rows
            if r and r[0].strip().lstrip("-").isdigit()]


# --------------------------------------------------------------------------- #
# Messages
# --------------------------------------------------------------------------- #
def msg_started(run, detail=""):
    name = JOB_NAME.get(run["jobid"], f"Job {run['jobid']}")
    line = f":arrow_forward: {name} STARTED {run['start']:%m/%d %H:%M}"
    return line + (f"\n{detail}" if detail else "")


def msg_finished(run, first_seen_finished=False, detail=""):
    """Finish line. `first_seen_finished` means we never saw it running (it had
    already completed by the time the first check of the day ran), so the line
    carries the full span instead of implying we watched it."""
    name = JOB_NAME.get(run["jobid"], f"Job {run['jobid']}")
    cls = classify(run["status"])
    emoji = {"SUCCESS": ":white_check_mark:", "FAILED": ":x:"}.get(cls, ":warning:")
    word = {"SUCCESS": "FINISHED", "FAILED": "FAILED"}.get(cls, run["status"].upper() or "ENDED")
    span = ""
    if run["start"] and run["end"]:
        dur = fmt_dur(run["start"], run["end"])
        span = f" {run['start']:%H:%M}-{run['end']:%H:%M}" + (f" ({dur})" if dur else "")
    elif run["end"]:
        span = f" {run['end']:%m/%d %H:%M}"
    prefix = "already " if first_seen_finished and cls == "SUCCESS" else ""
    line = f"{emoji} {name} {prefix}{word}{span}"
    return line + (f"\n{detail}" if detail else "")


def msg_all_done(day):
    return (f":checkered_flag: All three EverNorthRx jobs are complete for "
            f"{day:%m/%d} - Masterload Load, Masterload Snap and COBC Load have "
            f"finished. No further checks today.")


# --------------------------------------------------------------------------- #
# Main check
# --------------------------------------------------------------------------- #
def day_key(now):
    return f"{now:%Y-%m-%d}"


def fresh_day(now):
    return {
        "date": day_key(now),
        "started_posted": [],   # QueueIds whose START line has posted
        "ended_posted": [],     # QueueIds whose FINISH line has posted
        "seen": [],             # QueueIds observed at all (running or not)
        "done_posted": False,   # the "all three complete" line has posted
        "complete": False,      # stop checking for the rest of the day
        "last_check": None,     # ISO timestamp of the last RAMP query
    }


def carry_over_open_runs(state, now):
    """A run that started late yesterday and finished after the window closed
    would otherwise never get its FINISH line. Anything START-posted but not
    FINISH-posted is carried into the new day and resolved on the first check."""
    prev = state.get("day") or {}
    open_qids = [q for q in prev.get("started_posted", [])
                 if q not in prev.get("ended_posted", [])]
    return open_qids


def due(day, now):
    """Is a RAMP check due? -> (bool, cadence_minutes, why)"""
    last = parse_dt(day.get("last_check") or "")
    cadence = FAST_MINUTES if day.get("phase") == "fast" else SLOW_MINUTES
    if last is None:
        return True, cadence, "first check of the day"
    elapsed = (now - last).total_seconds()
    if elapsed + CADENCE_SLACK_SECONDS >= cadence * 60:
        return True, cadence, f"{int(elapsed // 60)}m since last check (cadence {cadence}m)"
    return False, cadence, (f"only {int(elapsed // 60)}m since last check "
                            f"(cadence {cadence}m) - skipping")


def check(state, now, carried):
    today = day_key(now)
    day = state["day"]
    runs = todays_runs(today)

    # Resolve anything left open from a previous day first.
    for qid in carried:
        r = run_by_qid(qid)
        if r and r["end"]:
            log(f"carry-over: QueueId {qid} finished after yesterday's window closed")
            if post(msg_finished(r)):
                day["ended_posted"].append(qid)

    by_job = {}
    for r in runs:
        by_job.setdefault(r["jobid"], []).append(r)

    for r in runs:
        qid = r["qid"]
        first_seen = qid not in day["seen"]
        if first_seen:
            day["seen"].append(qid)

        # STARTED -- ONLY while the run is still going. Never announce the start
        # of something that has already ended: the FINISH line carries the full
        # span anyway, so a late start line would be both stale and out of order.
        # (Bug seen 2026-08-12 12:55: a recovery branch re-posted "STARTED 11:44"
        # for a run that finished at 11:49.)
        if r["start"] and r["end"] is None and qid not in day["started_posted"]:
            if post(msg_started(r, detail=start_detail(r))):
                day["started_posted"].append(qid)

        # FINISHED
        if r["end"] and qid not in day["ended_posted"]:
            already = qid not in day["started_posted"]
            if post(msg_finished(r, first_seen_finished=already,
                                 detail=finish_detail(r))):
                day["ended_posted"].append(qid)

    # ---- phase (cadence) ---------------------------------------------------
    master_done = any(r["end"] and classify(r["status"]) != "FAILED"
                      for r in by_job.get(JOB_MASTER_LOAD, []))
    follower_active = any(r["end"] is None
                          for r in by_job.get(JOB_MASTER_SNAP, []) + by_job.get(JOB_COBC_LOAD, []))
    day["phase"] = "fast" if (master_done or follower_active) else "slow"

    # ---- day complete? -----------------------------------------------------
    in_flight = [r for r in runs if r["end"] is None]
    finished_jobs = {r["jobid"] for r in runs if r["end"] is not None}
    all_three = all(j in finished_jobs for j in JOB_IDS)
    if all_three and not in_flight:
        if not day["done_posted"]:
            if post(msg_all_done(now)):
                day["done_posted"] = True
        day["complete"] = True
        log("all three jobs finished and nothing in flight - done for the day")
    else:
        missing = [JOB_NAME[j] for j in JOB_IDS if j not in finished_jobs]
        log(f"phase={day['phase']}; still waiting on: "
            + (", ".join(missing) if missing else "(in-flight run)"))

    day["last_check"] = f"{now:%Y-%m-%d %H:%M:%S}"


def print_status():
    now = datetime.now()
    runs = todays_runs(day_key(now))
    if not runs:
        print(f"no EverNorthRx Masterload/COBC runs queued or started on {now:%Y-%m-%d}")
        return
    for r in runs:
        name = JOB_NAME.get(r["jobid"], str(r["jobid"]))
        st = f"{r['start']:%H:%M}" if r["start"] else "(queued)"
        en = f"{r['end']:%H:%M}" if r["end"] else "(running)"
        print(f"{r['qid']:>9}  {name:<34} {r['status']:<11} {st} -> {en}")


def main():
    if "--status" in sys.argv:
        print_status()
        return 0
    if "--drain" in sys.argv:
        drain()
        return 0
    if "--ack" in sys.argv:
        ack(sys.argv[sys.argv.index("--ack") + 1:])
        return 0

    now = datetime.now()
    state = load_state()

    # New day (or first ever run) -> reset, carrying any unfinished run forward.
    carried = []
    if (state.get("day") or {}).get("date") != day_key(now):
        carried = carry_over_open_runs(state, now)
        state["day"] = fresh_day(now)
        if carried:
            log(f"new day {day_key(now)}; carrying {len(carried)} unfinished run(s) forward")

    day = state["day"]

    if not FORCE and not DRY:
        if now.weekday() > 4:
            log("weekend - no checks")
            return 0
        if now.hour < START_HOUR:
            log(f"before {START_HOUR}:00 - no checks yet")
            return 0
        if day.get("complete"):
            return 0        # silent: this is the normal state after the cycle ends
        ok, cadence, why = due(day, now)
        if not ok:
            log(why)
            save_state(state)
            return 0
        log(f"checking RAMP ({why})")

    try:
        check(state, now, carried)
    except Exception as e:
        log(f"ERROR during check: {e}")
        return 1

    save_state(state)
    return 0


if __name__ == "__main__":
    sys.exit(main())
