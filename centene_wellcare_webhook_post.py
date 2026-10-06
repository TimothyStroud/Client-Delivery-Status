"""
Headless Centene / WellCare RAMP status poster (zero Claude tokens). Modeled on
the Aetna posters (aetnasubro_webhook_post.py) -> #data-operations-centene-wellcare-updates.

Usage: centene_wellcare_webhook_post.py <group-key> [--dry]
One Windows task per group ("CW Tick <key>"), every 4h, staggered 5 min apart
(all weekly loads, per user 2026-10-05). Each run posts the group's job grid:
  - every job's LatestJobRun status (Job/List)
  - Stage jobs list the files they staged (RAMP FileLog, newest batch)
  - a "for review" note for any job on the RAMP Dashboard queue that is not
    moving: in-flight > max(2x AvgRunTime, 60m), or queued unstarted > 60m
    (a disabled job's queued entry is called out -- never enable it).
Content dedupe per group: an unchanged grid is not re-posted.

Webhook URL lives OFF the git repo: H:\\slack_wf_centene_wellcare_updates.txt.
POST body key "Text"; Workflow Builder renders :emoji: only (no bold/code).
"""
import sys, os, re, json, hashlib, subprocess, urllib.request
from datetime import datetime, timedelta

BASE = r'C:\Users\tls2\.claude\projects\H--'
URL_FILE = r'H:\slack_wf_centene_wellcare_updates.txt'
LOG_FILE = r'H:\centene_wellcare_webhook_post.log'
STATE_FILE = os.path.join(BASE, 'centene_wellcare_post_state.json')
RAMP_SQL_SERVER = 'TRGUTIL10'
RAMP_OK = ('Successful', 'Resolved')
MAX_FILES = 25
CUTOFF = datetime(2026, 10, 5)  # ignore runs started before this (user 2026-10-06)

# key: (title, emoji, [(JobId, name), ...])
GROUPS = {
    'centenerx': ('Centene RX', ':pill:', [
        (1955, 'Centene RX 0100 Eligibility CNC Stage'), (1956, 'Centene RX 0110 Eligibility CNC Load'),
        (10269, 'Centene RX 0120 Eligibility CNC Snap'), (1959, 'Centene RX 0130 Claims Stage'),
        (1960, 'Centene RX 0140 Claims Load'), (1969, 'Centene RX 0150 TRR Stage'),
        (1970, 'Centene RX 0160 TRR Load'), (1963, 'Centene RX 0170 COBC Stage'),
        (1964, 'Centene RX 0180 COBC Load'), (1961, 'Centene RX 0190 Snap'),
        (1962, 'Centene RX 0200 PostSnap'), (10621, 'Centene RX 0210 MINE Snap'),
        (1965, 'Centene RX Group 0100 Stage'), (1966, 'Centene RX Group 0110 Load'),
        (1957, 'Centene RX HNT Eligibility 0100 Stage'), (1958, 'Centene RX HNT Eligibility 0110 Load'),
        (10270, 'Centene RX HNT Eligibility 0120 Snap')]),
    'fidelisrx': ('Centene Fidelis Rx', ':pill:', [
        (10272, 'Centene Fidelis Rx 0100 Eligibility Stage'), (10273, 'Centene Fidelis Rx 0110 Eligibility Load'),
        (10274, 'Centene Fidelis Rx 0120 Claims Stage'), (10275, 'Centene Fidelis Rx 0130 Claims Load'),
        (10618, 'Centene Fidelis Rx 0140 Snap'), (10619, 'Centene Fidelis Rx 0150 Post Snap'),
        (10620, 'Centene Fidelis Rx 0160 MINE Snap')]),
    'centenemed': ('Centene Medical', ':hospital:', [
        (1929, 'Centene Medical 0100 Claims Stage'), (1930, 'Centene Medical 0110 Claims Load'),
        (1931, 'Centene Medical 0120 Claims Bump Load'), (2222, 'Centene Medical 0130 Support Stage'),
        (2223, 'Centene Medical 0140 Support Load'), (11369, 'Centene Medical 0150 Snap')]),
    'fidelismed': ('Centene Fidelis Medical', ':hospital:', [
        (10019, 'Centene Fidelis Medical 0100 Eligibility Stage'), (10020, 'Centene Fidelis Medical 0110 Eligibility Load'),
        (10017, 'Centene Fidelis Medical 0120 Claims Stage'), (10018, 'Centene Fidelis Medical 0130 Claims Load')]),
    'healthnet': ('HealthNet', ':hospital:', [
        (1811, 'HealthNet 0100 Claims Stage'), (1812, 'HealthNet 0110 Claims Load'),
        (1809, 'Healthnet 0200 Eligibility Stage'), (1810, 'Healthnet 0210 Eligibility Load')]),
    'qualchoice': ('CenteneQualChoice', ':hospital:', [
        (10741, 'CenteneQualChoice 0100 Stage'), (10742, 'CenteneQualChoice 0110 Load')]),
    'wellcarerx': ('WellCareRx Masterload', ':pill:', [
        (10553, 'WellCareRx Masterload 0100 Stage'), (10554, 'WellCareRx Masterload 0110 Load'),
        (10589, 'WellCareRx Masterload 0120 Snap'), (10590, 'WellCareRx Masterload 0130 Post Snap'),
        (10633, 'WellCareRx Masterload 0140 MINE Snap')]),
    'wellcaremed': ('WellCare Medical', ':hospital:', [
        (10179, 'WellCare Medical 0100 Stage'), (10180, 'WellCare Medical 0110 Load')]),
}


def log(msg):
    line = f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}"
    try:
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(line + "\n")
    except Exception:
        pass
    print(line)


def ramp_api(path):
    out = subprocess.run(['curl', '-s', '--negotiate', '-u', ':', 'http://ramp/api/Ramp/' + path],
                         capture_output=True, text=True, timeout=300)
    d = json.loads(out.stdout)['Data']
    return d[0] if (isinstance(d, list) and d and isinstance(d[0], list)) else (d or [])


def ramp_sql(query):
    try:
        out = subprocess.run(['sqlcmd', '-S', RAMP_SQL_SERVER, '-d', 'RAMP', '-E', '-W', '-h', '-1',
                              '-s', '|', '-Q', 'SET NOCOUNT ON; ' + query],
                             capture_output=True, text=True, timeout=120)
    except Exception:
        return []
    return [[c.strip() for c in l.split('|')] for l in out.stdout.splitlines()
            if l.strip() and not set(l.strip()) <= set('-|')]


def to_dt(v):
    try:
        return datetime.fromisoformat(str(v).split('.')[0]) if v else None
    except Exception:
        return None


def fmt(v):
    d = to_dt(v)
    return d.strftime('%m/%d %I:%M%p').lower() if d else '?'


def hrs(mins):
    return f"{mins / 60:.1f}h" if mins >= 60 else f"{int(mins)}m"


def staged_files(jobid, since=None):
    """(files, expected) for the newest Queue run of this Stage job that logged
    files. expected = most common file count across its last 8 batches."""
    q = [r for r in ramp_sql(f"SELECT TOP 8 fl.QueueId, COUNT(*) FROM ramp.FileLog fl "
                             f"JOIN ramp.Queue q ON q.QueueId = fl.QueueId WHERE q.JobId = {int(jobid)} "
                             + (f"AND q.EndDate >= '{since:%Y-%m-%d %H:%M:%S}' " if since else "") +
                             f"GROUP BY fl.QueueId ORDER BY fl.QueueId DESC")
         if len(r) == 2 and r[0].isdigit() and r[1].isdigit()]
    if not q:
        return [], 0
    counts = [int(r[1]) for r in q]
    expected = max(set(counts), key=lambda c: (counts.count(c), c))
    files = [r[0] for r in ramp_sql(f"SELECT FileName FROM ramp.FileLog WHERE QueueId = {q[0][0]} ORDER BY FileName")
             if r and r[0]]
    return files, expected


def job_line(name, j):
    lr = (j or {}).get('LatestJobRun') or {}
    st, s, e = lr.get('Status') or 'Never run', lr.get('StartDate'), lr.get('EndDate')
    if e and st in RAMP_OK:
        icon, det = ':white_check_mark:', f"started {fmt(s)} | completed {fmt(e)}"
    elif e and st == 'Failed':
        icon, det = ':x:', f"started {fmt(s)} | ended {fmt(e)} - please investigate"
    elif s and not e:
        icon, det = ':loading:', f"started {fmt(s)} | not yet complete"
    elif lr and not s:
        icon, det = ':hourglass_flowing_sand:', "queued, not yet started"
    else:
        icon, det = ':grey_question:', (f"started {fmt(s)} | ended {fmt(e)}" if s else "no run on record")
    if j and j.get('Enabled') != 1:
        det += " | job DISABLED in RAMP"
    return f"{icon} {name} - {st}\n{det}"


# ---- SQL Agent step + ETA (Job Activity Monitor), modeled on ramp_aetnahrp_status_digest
# The RAMP queue's JobXML names the Agent job each task drives:
#   <task taskname="SqlAgentMonitor" status=...><jobname server="ETL4">ETL WellCare MasterLoad</jobname>
STALE_ACTIVITY_DAYS = 4
STEP_PCT = 80   # step overruns are one-sided, so p80 (not median), as in the Aetna digests


def agent_sql(server, query):
    try:
        out = subprocess.run(['sqlcmd', '-S', server, '-d', 'msdb', '-E', '-W', '-h', '-1',
                              '-s', '|', '-Q', 'SET NOCOUNT ON; ' + query],
                             capture_output=True, text=True, timeout=120)
    except Exception:
        return []
    return [[c.strip() for c in l.split('|')] for l in out.stdout.splitlines()
            if l.strip() and not set(l.strip()) <= set('-|')]


def agent_targets(jobid):
    """[(server, agent_job)] from the newest queue entry's SqlAgentMonitor tasks,
    not-yet-complete tasks first (the one the RAMP job is waiting on now)."""
    try:  # -y 0: untruncated nvarchar(max) (default display width cuts the XML at 256)
        xml = subprocess.run(['sqlcmd', '-S', RAMP_SQL_SERVER, '-d', 'RAMP', '-E', '-y', '0', '-Q',
                              f"SET NOCOUNT ON; SELECT TOP 1 CAST(JobXML AS nvarchar(max)) FROM ramp.Queue "
                              f"WHERE JobId = {int(jobid)} ORDER BY QueueId DESC"],
                             capture_output=True, text=True, timeout=120).stdout
    except Exception:
        xml = ''
    todo, done = [], []
    for st, srv, nm in re.findall(r'<task[^>]*taskname="SqlAgentMonitor"[^>]*status="([^"]*)"[^>]*>'
                                  r'.*?<jobname server="([^"]+)"[^>]*>([^<]+)</jobname>', xml):
        srv = srv.upper() if srv.upper().startswith('TRG') else 'TRG' + srv.upper()
        (done if st == 'complete' else todo).append((srv, nm.replace('&amp;', '&')))
    return todo + done


def _jid(name):
    return f"DECLARE @jid uniqueidentifier=(SELECT job_id FROM msdb.dbo.sysjobs WHERE name=N'{name}'); "


def live_step(server, name):
    """(run_start, step_id, step_start) of the executing run, or None (sysjobactivity,
    recency-guarded against orphaned rows left by Agent restarts)."""
    r = agent_sql(server, _jid(name) +
                  "SELECT TOP 1 CONVERT(varchar(19), ja.start_execution_date, 120), "
                  "ISNULL(ja.last_executed_step_id, 0), CONVERT(varchar(19), ja.last_executed_step_date, 120) "
                  "FROM msdb.dbo.sysjobactivity ja WITH (NOLOCK) WHERE ja.job_id=@jid "
                  "AND ja.start_execution_date IS NOT NULL AND ja.stop_execution_date IS NULL "
                  f"AND ja.start_execution_date > DATEADD(day,-{STALE_ACTIVITY_DAYS},GETDATE()) "
                  "ORDER BY ja.session_id DESC, ja.start_execution_date DESC;")
    for p in r:
        if len(p) >= 3 and to_dt(p[0]):
            # Agent stamps last_executed_step_id/_date when a step STARTS (verified 10/06 on
            # ETL CenteneFidelisRx MasterLoad: step 5 @ 05:59:55); both stay NULL during step 1
            sid = int(p[1]) if p[1].isdigit() and int(p[1]) else 1
            return to_dt(p[0]), sid, (to_dt(p[2]) if sid > 1 else None) or to_dt(p[0])
    return None


def step_names(server, name):
    return {int(p[0]): p[1] for p in agent_sql(server, _jid(name) +
            "SELECT step_id, step_name FROM msdb.dbo.sysjobsteps WHERE job_id=@jid ORDER BY step_id;")
            if len(p) >= 2 and p[0].isdigit()}


def remaining_from_step(server, name, sid, days=120):
    """Ascending seconds from the START of step sid to the END of the job over recent
    successful runs (each step row bucketed to the next step_id=0 outcome row)."""
    r = agent_sql(server, _jid(name) + f"DECLARE @sid int={int(sid)}; "
                  "WITH h AS (SELECT instance_id, step_id, run_status, "
                  "  (run_duration/10000)*3600+((run_duration/100)%100)*60+(run_duration%100) AS secs "
                  "  FROM msdb.dbo.sysjobhistory WITH (NOLOCK) WHERE job_id=@jid "
                  f"  AND run_date>=CONVERT(int,CONVERT(varchar(8),DATEADD(day,-{int(days)},GETDATE()),112))), "
                  "o AS (SELECT instance_id, run_status FROM h WHERE step_id=0), "
                  "t AS (SELECT h.step_id, h.secs, h.run_status, "
                  "  (SELECT MIN(o.instance_id) FROM o WHERE o.instance_id>h.instance_id) AS rk "
                  "  FROM h WHERE h.step_id>=@sid) "
                  "SELECT SUM(t.secs) FROM t JOIN o ON o.instance_id=t.rk AND o.run_status=1 GROUP BY t.rk "
                  "HAVING MIN(t.run_status)=1 AND MAX(CASE WHEN t.step_id=@sid THEN 1 ELSE 0 END)=1 ORDER BY 1;")
    return sorted(int(p[0]) for p in r if p and p[0].isdigit())


def _pct(vals, p):
    import math
    return vals[max(1, math.ceil(p / 100.0 * len(vals))) - 1] if vals else None


def _eta_stamp(dt):
    return dt.strftime('%I:%M%p' if dt.date() == datetime.now().date() else '%m/%d %I:%M%p').lower().lstrip('0')


def agent_status(jobid):
    """'SQL TRGETL4 ETL WellCare MasterLoad - Step 3/9 (name)\nETA ~4:10pm' for the
    Agent job this RAMP job is driving, or '' when none of its Agent jobs is executing."""
    for server, name in agent_targets(jobid):
        live = live_step(server, name)
        if not live:
            continue
        start, sid, sstart = live
        names = step_names(server, name)
        head = f"SQL {server} {name} - Step {sid}" + (f"/{max(names)}" if names else "")
        if names.get(sid):
            head += f" ({names[sid]})"
        now = datetime.now()
        if names and sid >= max(names):
            return head + "\nfinal step - wrapping up"
        in_step = (now - (sstart or start)).total_seconds()
        possible = [d for d in remaining_from_step(server, name, sid) if d >= in_step]
        if possible:
            eta = (sstart or start) + timedelta(seconds=_pct(possible, STEP_PCT))
            return head + ("\nETA ~" + _eta_stamp(eta) if eta > now else "\nwrapping up")
        return head + f"\nrunning {hrs((now - start).total_seconds() / 60)} - longer than usual, still processing"
    return ''


def review_notes(names, jobs, queue):
    """Dashboard (Queue/List) entries for these jobs that are open but not moving."""
    now, notes = datetime.now(), []
    for q in queue:
        jid = q.get('JobId')
        if jid not in names or q.get('Status') in RAMP_OK + ('Failed',) or q.get('EndDate'):
            continue
        avg = q.get('AvgRunTime') or 0
        start, created = to_dt(q.get('StartDate')), to_dt(q.get('CreateDate'))
        if start:
            el = (now - start).total_seconds() / 60
            if el > max(2 * avg, 60):
                notes.append(f":warning: {names[jid]} - {q.get('Status')} for {hrs(el)} (avg {hrs(avg)}), "
                             f"QueueId {q.get('QueueId')}, not moving - please review")
        elif created and (now - created).total_seconds() / 60 > 60:
            el = (now - created).total_seconds() / 60
            dis = (jobs.get(jid) or {}).get('Enabled') != 1
            notes.append(f":warning: {names[jid]} - queued {hrs(el)} without starting, QueueId {q.get('QueueId')}"
                         + (" (job is DISABLED in RAMP, so it will not pick up)" if dis else "")
                         + " - please review")
    return notes


def build(key, announced):
    """Only runs started >= CUTOFF. In-flight jobs show every post; a finished run
    shows once (tracked in announced {jid: StartDate}), then drops off until the
    job runs again. Queued/never-started jobs are not shown. Files: Claims only."""
    title, emoji, members = GROUPS[key]
    jobs = {j.get('JobId'): j for j in ramp_api('Job/List')}
    try:
        queue = ramp_api('Queue/List')
    except Exception:
        queue = []
    names = dict(members)
    body, finished = [], {}
    for jid, nm in members:
        lr = (jobs.get(jid) or {}).get('LatestJobRun') or {}
        s, e = lr.get('StartDate'), lr.get('EndDate')
        sd = to_dt(s)
        if not sd or sd < CUTOFF:
            continue
        if e:
            if announced.get(str(jid)) == s:
                continue
            finished[str(jid)] = s
        block = job_line(nm, jobs.get(jid))
        if not e:
            ag = agent_status(jid)
            if ag:
                block += '\n' + ag
        if nm.endswith('Stage') and e:  # files only once this run's stage finished
            files = [f for f in staged_files(jid, sd)[0] if 'CLAIM' in f.upper()]
            if files:
                block += f"\n\nClaims files staged ({len(files)}):\n" + "\n".join(files[:MAX_FILES])
                if len(files) > MAX_FILES:
                    block += f"\n+{len(files) - MAX_FILES} more"
        body.append(block)
    msg = "\n\n".join(body)
    notes = review_notes(names, jobs, queue)
    if notes:
        msg += ("\n\n" if msg else "") + "For review:\n" + "\n".join(notes)
    return f"{title} load status - {datetime.now():%m/%d %I:%M%p}\n{msg}", msg, finished


def main():
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    if not args or args[0] not in GROUPS:
        print("usage: centene_wellcare_webhook_post.py <" + "|".join(GROUPS) + "> [--dry]")
        return 2
    key = args[0]
    try:
        state = json.load(open(STATE_FILE))
    except Exception:
        state = {}
    ann = state.setdefault('announced', {}).setdefault(key, {})
    text, content, finished = build(key, ann)
    if '--dry' in sys.argv:
        print(text or '(nothing to post)')
        return 0
    if not content:
        log(f"{key}: nothing new, not posted")
        return 0
    h = hashlib.sha1(content.encode()).hexdigest()
    if state.get(key) == h:
        log(f"{key}: unchanged, not posted")
        return 0
    url = open(URL_FILE, encoding='utf-8').read().strip()
    req = urllib.request.Request(url, data=json.dumps({'Text': text}).encode(),
                                 headers={'Content-Type': 'application/json'})
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            r.read()
        state[key] = h
        ann.update(finished)
        json.dump(state, open(STATE_FILE, 'w'), indent=1)
        log(f"{key}: posted")
    except Exception as e:
        log(f"{key}: post error (retry next tick): {e}")
    return 0


if __name__ == '__main__':
    sys.exit(main())
