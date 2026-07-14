"""
Kaiser Pareo Prepay Load & Snap monitor (headless, zero Claude tokens).

Watches RAMP for each new 'Kaiser Pareo Prepay' daily cycle:
  0100 Stage (JobId 10720)  ->  0110 Load (10721)  ->  0120 Snap (10722)
For every NEW Load (QueueId > watermark) that has finished, once its Snap is
also terminal (or the Load failed, or the Snap is clearly stuck), posts ONE
combined Load+Snap Success/Failure message to a Slack Workflow Builder webhook:
  - RDP Ops - KaiserPrePay    -> #rps_kaiserprepay_discussion  (H:\slack_wf_kaiserprepay.txt)
(The #team-rdp-operations-support post was removed 2026-07-07 per user request.)
The message confirms the MA & SC files were staged and notes any missing one.

State (last-posted Load QueueId) lives at H:\kaiserprepay_monitor_state.json so
6/28 (watermark seed) is never posted; posting "starts with the next load".

Modes:
  (default)   query, post any newly-completed cycles, advance watermark.
  --dry-run   query + print what WOULD post; no Slack post, no watermark change.
"""
import sys, os, re, json, subprocess, urllib.request
from datetime import datetime, timedelta

SQL_SERVER = "TRGUTIL10"
JOB_STAGE, JOB_LOAD, JOB_SNAP = 10720, 10721, 10722
SUPPORT_URL_FILE = r"H:\slack_wf_support.txt"
PREPAY_URL_FILE  = r"H:\slack_wf_kaiserprepay.txt"
STATE_FILE = r"H:\kaiserprepay_monitor_state.json"
LOG_FILE   = r"H:\kaiserprepay_monitor.log"
SEP = "\x1f"
# If a Load succeeded but its Snap is still not terminal this long after the
# Load finished, post anyway (flagging the snap) instead of stalling forever.
SNAP_WAIT_HOURS = 3

DRY = "--dry-run" in sys.argv


def log(msg):
    line = f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}"
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass
    print(line)


def run_sql(query):
    r = subprocess.run(
        ["sqlcmd", "-S", SQL_SERVER, "-d", "RAMP", "-E", "-Q", query,
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
    if not s or s == "NULL":
        return None
    try:
        return datetime.strptime(s[:19], "%Y-%m-%d %H:%M:%S")
    except ValueError:
        return None


def classify(status):
    s = (status or "").strip().lower()
    if s.startswith("success") or s == "resolved":
        return "SUCCESS"
    if s == "failed":
        return "FAILED"
    return "PENDING"


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


def get_url(path):
    try:
        return open(path, encoding="utf-8").read().strip()
    except Exception:
        return ""


def post(url, text):
    data = json.dumps({"Text": text}).encode("utf-8")
    req = urllib.request.Request(url, data=data, headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=30) as r:
        return r.status, r.read().decode("utf-8", "replace").strip()


def post_all(text):
    for label, path in (("kaiserprepay", PREPAY_URL_FILE),):
        url = get_url(path)
        if not url:
            log(f"  SKIP {label}: no webhook at {path}")
            continue
        if DRY:
            log(f"  [dry-run] would POST -> {label}")
            continue
        try:
            st, body = post(url, text)
            log(f"  POSTED -> {label}: HTTP {st} {body[:60]}")
        except Exception as e:
            log(f"  ERROR posting -> {label}: {e}")


def cycle_stage(load_start, load_date):
    """Stage run (10720) that fed this load: latest finishing on/before the
    load start, same calendar day."""
    rows = run_sql(
        "SET NOCOUNT ON; SELECT TOP 1 QueueId, Status, "
        "CONVERT(varchar(19),StartDate,121), CONVERT(varchar(19),EndDate,121) "
        f"FROM [RAMP].[ramp].[Queue] WHERE JobId={JOB_STAGE} "
        f"AND EndDate IS NOT NULL AND EndDate <= '{load_start}' "
        f"AND CAST(StartDate AS date)='{load_date}' ORDER BY QueueId DESC"
    )
    return rows[0] if rows else None


def cycle_snap(load_end, load_date):
    """Snap run (10722) for this cycle: earliest starting at/after the load
    end, same calendar day."""
    rows = run_sql(
        "SET NOCOUNT ON; SELECT TOP 1 QueueId, Status, "
        "CONVERT(varchar(19),StartDate,121), CONVERT(varchar(19),EndDate,121) "
        f"FROM [RAMP].[ramp].[Queue] WHERE JobId={JOB_SNAP} "
        f"AND StartDate >= '{load_end}' "
        f"AND CAST(StartDate AS date)='{load_date}' ORDER BY QueueId ASC"
    )
    return rows[0] if rows else None


def staged_files(stage_qid):
    """MA/SC prepay files marked Staged under the given stage QueueId."""
    if not stage_qid:
        return {}
    rows = run_sql(
        "SET NOCOUNT ON; SELECT FileName, Status FROM [RAMP].[ramp].[FileLog] "
        f"WHERE QueueId={stage_qid} AND Status='Staged' "
        "AND FileName LIKE '%Prepay_Claim_Lines%' ORDER BY FileName"
    )
    found = {}
    for r in rows:
        fn = r[0].strip()
        m = re.search(r"_Lines_([A-Z]{2})_", fn)
        if m:
            found[m.group(1)] = fn
    return found


def data_date_from(files):
    for fn in files.values():
        m = re.search(r"_(\d{8})\d*\.csv", fn)
        if m:
            d = m.group(1)
            return f"{d[4:6]}/{d[6:8]}/{d[0:4]}"
    return None


def fmt_span(start, end):
    s = parse_dt(start); e = parse_dt(end)
    if s and e:
        return f"{s:%m/%d %H:%M}-{e:%H:%M}"
    if s:
        return f"{s:%m/%d %H:%M}"
    return "?"


def build_message(stage, load, snap, files, load_date):
    ddate = data_date_from(files) or load_date
    load_cls = classify(load[1])
    snap_cls = classify(snap[1]) if snap else "PENDING"

    if load_cls != "SUCCESS":
        head_emoji, head_word = ":x:", "Load FAILED"
    elif snap is None or snap_cls == "PENDING":
        head_emoji, head_word = ":warning:", "Load SUCCESS / Snap DID NOT COMPLETE"
    elif snap_cls == "FAILED":
        head_emoji, head_word = ":x:", "Snap FAILED"
    else:
        head_emoji, head_word = ":white_check_mark:", "Load & Snap SUCCESS"

    lines = [f"{head_emoji} Kaiser Pareo Prepay - {head_word} (data date {ddate})"]
    if stage:
        lines.append(f"0100 Stage: {classify(stage[1])} ({fmt_span(stage[2], stage[3])})")
    else:
        lines.append("0100 Stage: (no matching stage run found)")
    lines.append(f"0110 Load: {load_cls} ({fmt_span(load[2], load[3])})")
    if snap and snap_cls != "PENDING":
        lines.append(f"0120 Snap: {snap_cls} ({fmt_span(snap[2], snap[3])})")
    elif load_cls != "SUCCESS":
        lines.append("0120 Snap: N/A (load did not succeed)")
    else:
        lines.append(f"0120 Snap: DID NOT COMPLETE (not terminal after {SNAP_WAIT_HOURS}h)")

    for state in ("MA", "SC"):
        if state in files:
            lines.append(f":white_check_mark: {state}: {files[state]}")
        else:
            lines.append(f":x: {state}: NOT LOADED - no staged {state} file found")
    return "\n".join(lines)


def main():
    state = load_state()
    watermark = int(state.get("last_load_queueid", 0))

    # Seed watermark on first run so 6/28 (and anything already done) is skipped
    # and posting "starts with the next load".
    if watermark == 0:
        rows = run_sql(
            "SET NOCOUNT ON; SELECT ISNULL(MAX(QueueId),0) FROM [RAMP].[ramp].[Queue] "
            f"WHERE JobId={JOB_LOAD} AND EndDate IS NOT NULL"
        )
        watermark = int(rows[0][0]) if rows else 0
        log(f"seeded watermark to current latest Load QueueId={watermark}; no post this run")
        save_state({"last_load_queueid": watermark})
        return 0

    new_loads = run_sql(
        "SET NOCOUNT ON; SELECT QueueId, Status, "
        "CONVERT(varchar(19),StartDate,121), CONVERT(varchar(19),EndDate,121) "
        f"FROM [RAMP].[ramp].[Queue] WHERE JobId={JOB_LOAD} "
        f"AND QueueId > {watermark} AND EndDate IS NOT NULL ORDER BY QueueId ASC"
    )
    if not new_loads:
        log(f"no new Load since QueueId {watermark}")
        return 0

    now = datetime.now()
    for row in new_loads:
        load_qid = int(row[0])
        load = row  # [QueueId, Status, StartDate, EndDate]
        load_start, load_end = load[2], load[3]
        load_date = load_start[:10]
        load_cls = classify(load[1])

        stage = cycle_stage(load_start, load_date)
        snap = cycle_snap(load_end, load_date)
        snap_cls = classify(snap[1]) if snap else "PENDING"

        # Wait for the Snap only when the Load succeeded and the snap is still
        # in flight AND we're inside the grace window. Otherwise post now.
        if load_cls == "SUCCESS" and snap_cls == "PENDING":
            le = parse_dt(load_end)
            if le and (now - le) < timedelta(hours=SNAP_WAIT_HOURS):
                log(f"Load {load_qid} ({load_date}) ok; snap pending, waiting (<{SNAP_WAIT_HOURS}h). stop.")
                break  # don't advance watermark; re-check next run

        files = staged_files(stage[0] if stage else None)
        msg = build_message(stage, load, snap, files, load_date)
        log(f"posting cycle Load QueueId={load_qid} ({load_date}):")
        for ln in msg.splitlines():
            log("    " + ln)
        post_all(msg)

        state["last_load_queueid"] = load_qid
        save_state(state)

    return 0


if __name__ == "__main__":
    sys.exit(main())
