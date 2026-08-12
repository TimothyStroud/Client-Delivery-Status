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


def run_sql(query, db="RAMP"):
    r = subprocess.run(
        ["sqlcmd", "-S", SQL_SERVER, "-d", db, "-E", "-Q", query,
         "-W", "-s", SEP, "-h", "-1"],
        capture_output=True, text=True, check=False,
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
# Messages
# --------------------------------------------------------------------------- #
def msg_started(run):
    name = JOB_NAME.get(run["jobid"], f"Job {run['jobid']}")
    return f":arrow_forward: {name} STARTED {run['start']:%m/%d %H:%M}"


def msg_finished(run, first_seen_finished=False):
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
    return f"{emoji} {name} {prefix}{word}{span}"


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

        # STARTED -- only for runs we catch while they are still going. A run
        # that was already finished the first time we looked gets one FINISH
        # line instead of a stale start+finish pair.
        if r["start"] and qid not in day["started_posted"]:
            if r["end"] is None:
                if post(msg_started(r)):
                    day["started_posted"].append(qid)
            elif not first_seen:
                # We saw it running on an earlier tick but the start post
                # failed; still record the start so ordering stays sane.
                if post(msg_started(r)):
                    day["started_posted"].append(qid)

        # FINISHED
        if r["end"] and qid not in day["ended_posted"]:
            already = qid not in day["started_posted"]
            if post(msg_finished(r, first_seen_finished=already)):
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
