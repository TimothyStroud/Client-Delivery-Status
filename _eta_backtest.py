"""Walk-forward backtest of the Aetna digest ETA estimators.

Replays every recent successful run of the four Aetna SQL Agent jobs, simulating
a digest tick every 2 h, and scores three estimators using ONLY history that
existed before the run under test:

  whole      - anchored whole-job survival median (the pre-2026-08-04 behaviour,
               still the fallback while the SSIS step runs)
  current    - what the digests do today: step-aware once past the SSIS step,
               else the single ~90% SSIS milestone, else `whole`
  ladder     - proposed: step-aware once past the SSIS step, else the LATEST
               admissible SSIS task completed so far + its median tail, else
               `whole`

Job timings come from msdb.dbo.sysjobhistory (the table SSMS's Log File Viewer
renders); in-package task timings from SSISDB catalog.executable_statistics.
Analysis helper -- not part of any scheduled job.
"""
import statistics
import subprocess
import sys
from datetime import datetime, timedelta

HIST_DAYS = 240
TICK_HOURS = 2

# ---- proposed ladder admissibility -------------------------------------------
LADDER_MIN_SAMPLES = 5      # rungs need real history before we trust a tail
LADDER_MIN_FRAC = 0.40      # a task finishing at 2% of wall-clock tells us nothing
LADDER_MAX_FRAC = 0.97      # terminal tasks have no lead time left to be useful
LADDER_IQR_FLOOR = 20 * 60  # allow at least this much tail spread...
LADDER_IQR_REL = 0.60       # ...or this fraction of the tail, whichever is larger
LADDER_IQR_CAP = 90 * 60    # ...but never more than this in absolute terms

JOBS = [
    ("TRGETL2", "ETL AetnaHRP MasterLoad", "AetnaHRP_MasterLoad.dtsx"),
    ("TRGETL2", "ETL AetnaRx MasterLoad Claims And Eligibility",
     "AetnaRx_MasterLoad_Claims_And_Eligibility.dtsx"),
    ("TRGETL2", "SSIS AetnaRCE Daily Process", "AetnaRCE_Load.dtsx"),
    ("TRGETL4", "ETL NCStateAetna MasterLoad", None),
]


def sqlrows(server, query, db=None):
    cmd = ['sqlcmd', '-S', server, '-E', '-W', '-h', '-1', '-s', '|',
           '-Q', 'SET NOCOUNT ON; ' + query]
    if db:
        cmd[3:3] = ['-d', db]
    out = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
    if out.returncode != 0:
        print(f"  !! sqlcmd failed on {server}: {out.stderr.strip()[:300]}", file=sys.stderr)
    rows = []
    for line in out.stdout.splitlines():
        line = line.rstrip()
        if not line or set(line) <= set('-|') or 'rows affected' in line:
            continue
        rows.append([c.strip() for c in line.split('|')])
    return rows


def agent_dt(run_date, run_time):
    """SQL Agent stores date as YYYYMMDD int and time as HHMMSS int."""
    d, t = int(run_date), int(run_time)
    return datetime(d // 10000, (d // 100) % 100, d % 100,
                    t // 10000, (t // 100) % 100, t % 100)


def agent_secs(run_duration):
    v = int(run_duration)
    return (v // 10000) * 3600 + ((v // 100) % 100) * 60 + (v % 100)


def load_job_runs(server, name):
    """Ascending list of successful runs: dict(start, end, steps={sid:(start,secs)}).

    sysjobhistory has no run key and these jobs can run twice a day, so step rows
    are bucketed onto the next step_id=0 outcome row -- bucketed against ALL
    outcome rows, then filtered to the successful ones (bucketing against
    successes alone would graft a failed run's steps onto the next good run)."""
    rows = sqlrows(server,
                   f"DECLARE @jid uniqueidentifier=(SELECT job_id FROM msdb.dbo.sysjobs "
                   f"WHERE name=N'{name}'); "
                   "SELECT instance_id, step_id, run_status, run_date, run_time, run_duration "
                   "FROM msdb.dbo.sysjobhistory WITH (NOLOCK) WHERE job_id=@jid "
                   f"AND run_date>=CONVERT(int,CONVERT(varchar(8),"
                   f"DATEADD(day,-{HIST_DAYS},GETDATE()),112)) ORDER BY instance_id;")
    recs = []
    for r in rows:
        if len(r) >= 6 and r[0].isdigit():
            try:
                recs.append((int(r[0]), int(r[1]), int(r[2]),
                             agent_dt(r[3], r[4]), agent_secs(r[5])))
            except (ValueError, IndexError):
                continue
    outcomes = [x for x in recs if x[1] == 0]
    runs = []
    for inst, sid, status, start, secs in outcomes:
        if status != 1:
            continue
        steps = {}
        for i2, s2, st2, start2, secs2 in recs:
            if s2 > 0 and i2 < inst and start2 >= start:
                steps[s2] = (start2, secs2)
        if steps:
            runs.append({'start': start, 'end': start + timedelta(seconds=secs),
                         'secs': secs, 'steps': steps})
    return runs


def load_pkg_runs(pkg):
    """Ascending list of successful package executions: dict(start, end,
    tasks={executable_id: end_time})."""
    if not pkg:
        return []
    rows = sqlrows('TRGETLPROD2',
                   f"DECLARE @pkg sysname=N'{pkg}'; "
                   ";WITH ex AS (SELECT execution_id, start_time, end_time "
                   " FROM catalog.executions WHERE package_name=@pkg AND status=7 "
                   " AND end_time IS NOT NULL "
                   f" AND start_time>DATEADD(day,-{HIST_DAYS},GETDATE())) "
                   "SELECT ex.execution_id, CONVERT(varchar(19),ex.start_time,120), "
                   " CONVERT(varchar(19),ex.end_time,120), es.executable_id, "
                   " CONVERT(varchar(19),es.end_time,120) "
                   "FROM ex JOIN catalog.executable_statistics es "
                   " ON es.execution_id=ex.execution_id ORDER BY ex.execution_id;",
                   db='SSISDB')
    by_id = {}
    for r in rows:
        if len(r) < 5 or not r[0].isdigit():
            continue
        try:
            eid = int(r[0])
            rec = by_id.setdefault(eid, {
                'start': datetime.strptime(r[1], '%Y-%m-%d %H:%M:%S'),
                'end': datetime.strptime(r[2], '%Y-%m-%d %H:%M:%S'), 'tasks': {}})
            rec['tasks'][int(r[3])] = datetime.strptime(r[4], '%Y-%m-%d %H:%M:%S')
        except (ValueError, KeyError):
            continue
    return [by_id[k] for k in sorted(by_id)]


def steps_meta(server, name):
    rows = sqlrows(server,
                   f"DECLARE @jid uniqueidentifier=(SELECT job_id FROM msdb.dbo.sysjobs "
                   f"WHERE name=N'{name}'); SELECT step_id, subsystem FROM msdb.dbo.sysjobsteps "
                   "WHERE job_id=@jid ORDER BY step_id;")
    meta = {}
    for r in rows:
        if len(r) >= 2 and r[0].isdigit():
            meta[int(r[0])] = r[1].upper()
    return meta


def pct(vals, p):
    """Nearest-rank percentile of an ascending list (matches the digests' _pct)."""
    if not vals:
        return None
    import math
    k = max(1, math.ceil(p / 100.0 * len(vals)))
    return sorted(vals)[k - 1]


# ---- estimators ---------------------------------------------------------------

def est_whole(prior_runs, run_start, now):
    """Anchored whole-job survival median over the last 8 successful runs."""
    durs = sorted(r['secs'] for r in prior_runs[-8:])
    if not durs:
        return None
    elapsed = (now - run_start).total_seconds()
    possible = [d for d in durs if d >= elapsed]
    if not possible:
        return None
    return run_start + timedelta(seconds=pct(possible, 50))


def step_remaining(prior_runs, sid):
    """Ascending step-start-to-job-end wall-clock for `sid` over prior runs."""
    out = []
    for r in prior_runs:
        if sid in r['steps']:
            out.append((r['end'] - r['steps'][sid][0]).total_seconds())
    return sorted(out)


def est_step_mean(prior_runs, sid, step_start, now, n=6):
    """RCE's CURRENTLY DEPLOYED per-step estimator: step_start + the sum of the
    arithmetic MEAN duration of the current and all later steps over the last n
    successful runs; if the current step has already outlived that, fall back to
    now + the mean of the steps after it. Reproduced here to measure the skew the
    mean carries (step 7 averages 1h20m against a 47m median, max 10h21m)."""
    hist = {}
    for r in prior_runs:
        for s, (_st, secs) in r['steps'].items():
            hist.setdefault(s, []).append(secs)
    means = {s: sum(v[-n:]) / len(v[-n:]) for s, v in hist.items() if v}
    if not means:
        return None
    eta = step_start + timedelta(seconds=sum(v for s, v in means.items() if s >= sid))
    if eta <= now:
        eta = now + timedelta(seconds=sum(v for s, v in means.items() if s > sid))
    return eta if eta > now else now


def est_step(prior_runs, sid, step_start, now, max_sid, p=50):
    """Survival percentile within the live step (the 2026-08-04 step-aware path).
    `p` is swept: the median is best on typical runs, but a higher percentile buys
    back accuracy on the long tail, which is where these loads actually hurt."""
    rem = step_remaining(prior_runs, sid)
    if not rem:
        return step_start if sid >= max_sid else None
    in_step = (now - step_start).total_seconds()
    possible = [d for d in rem if d >= in_step]
    if not possible:
        return None
    return step_start + timedelta(seconds=pct(possible, p))


def build_milestone(prior_pkgs, min_frac, max_frac, single,
                    iqr_rel=None, iqr_cap=None):
    """Admissible SSIS rungs from prior executions.

    `single` reproduces today's chooser: the latest task with a STABLE end
    FRACTION (sd <= 0.08). Otherwise apply the proposed rule: enough samples, a
    late-enough fraction to carry information, and a tail whose spread is bounded
    in ABSOLUTE time -- which is what the ETA error actually is."""
    stats = {}
    for p in prior_pkgs[-8:]:
        total = (p['end'] - p['start']).total_seconds()
        if total < 600:
            continue
        for eid, tend in p['tasks'].items():
            frac = (tend - p['start']).total_seconds() / total
            tail = (p['end'] - tend).total_seconds()
            stats.setdefault(eid, []).append((frac, tail))
    out = {}
    for eid, vals in stats.items():
        if len(vals) < (4 if single else LADDER_MIN_SAMPLES):
            continue
        fracs = [v[0] for v in vals]
        tails = sorted(v[1] for v in vals)
        avg_frac = sum(fracs) / len(fracs)
        if not (min_frac <= avg_frac <= max_frac):
            continue
        if single:
            if len(fracs) > 1 and statistics.stdev(fracs) > 0.08:
                continue
        else:
            rel = LADDER_IQR_REL if iqr_rel is None else iqr_rel
            cap = LADDER_IQR_CAP if iqr_cap is None else iqr_cap
            iqr = pct(tails, 75) - pct(tails, 25)
            if iqr > min(cap, max(LADDER_IQR_FLOOR, rel * pct(tails, 50))):
                continue
        out[eid] = (avg_frac, pct(tails, 50))
    return out


def run_job(server, name, pkg, sweep=None):
    print(f"\n===== {name} =====")
    runs = load_job_runs(server, name)
    pkgs = load_pkg_runs(pkg)
    meta = steps_meta(server, name)
    ssis = max([k for k, v in meta.items() if v == 'SSIS'] or [1])
    max_sid = max(meta) if meta else 1
    print(f"  {len(runs)} successful job runs, {len(pkgs)} package executions, "
          f"SSIS step={ssis}, last step={max_sid}")
    if len(runs) < 12:
        print("  too little history to backtest")
        return

    # Match each job run to the package execution that overlaps its SSIS step.
    for r in runs:
        r['pkg'] = None
        if ssis in r['steps']:
            sstart = r['steps'][ssis][0]
            best, bestd = None, 3600 * 3
            for p in pkgs:
                d = abs((p['start'] - sstart).total_seconds())
                if d < bestd:
                    best, bestd = p, d
            r['pkg'] = best

    configs = sweep or [(LADDER_MIN_FRAC, LADDER_IQR_REL, LADDER_IQR_CAP)]
    STEP_PCTS = [50, 60, 70, 80]
    LADDER = (0.55, 0.35, 90 * 60)     # chosen by the earlier frac/IQR sweep
    variants = (['whole (pre-step-aware)', 'deployed HRP/Rx', 'deployed RCE (step MEAN)']
                + [f"step p{p}, no ladder" for p in STEP_PCTS]
                + [f"step p{p} + ladder" for p in STEP_PCTS])
    errs = {k: [] for k in variants}
    inssis = {k: [] for k in variants}
    bias = {k: [] for k in variants}
    hits = {'ladder': 0}

    for i in range(10, len(runs)):
        run, prior = runs[i], runs[:i]
        prior_pkgs = [p for p in pkgs if p['start'] < run['start']]
        single_ms = build_milestone(prior_pkgs, 0.0, 0.95, True)
        ladder_ms = build_milestone(prior_pkgs, LADDER[0], LADDER_MAX_FRAC, False,
                                    iqr_rel=LADDER[1], iqr_cap=LADDER[2])
        # post-SSIS T-SQL tail, added to any package-derived ETA (the digest
        # reports the JOB, but a package ETA ends at the package)
        post_med = pct(step_remaining(prior, ssis + 1), 50) or 0

        now = run['start'] + timedelta(hours=TICK_HOURS)
        while now < run['end']:
            live_sid = max([s for s, (st, _d) in run['steps'].items() if st <= now] or [1])
            live_start = run['steps'].get(live_sid, (run['start'], 0))[0]
            w = est_whole(prior, run['start'], now)

            def with_pkg(ms):
                """Latest admissible rung this run has completed by `now`."""
                if not run['pkg'] or not ms:
                    return None
                done = [(t, ms[e][1]) for e, t in run['pkg']['tasks'].items()
                        if e in ms and t <= now]
                if not done:
                    return None
                tend, tail = max(done, key=lambda x: x[0])
                return tend + timedelta(seconds=tail + post_med)

            vals = {'whole (pre-step-aware)': w}
            past_ssis = live_sid > ssis
            if past_ssis:
                vals['deployed RCE (step MEAN)'] = \
                    est_step_mean(prior, live_sid, live_start, now) or w
            else:
                vals['deployed RCE (step MEAN)'] = w
            for p in STEP_PCTS:
                s = (est_step(prior, live_sid, live_start, now, max_sid, p=p) or w
                     if past_ssis else None)
                vals[f"step p{p}, no ladder"] = s if past_ssis else w
                if past_ssis:
                    vals[f"step p{p} + ladder"] = s
                else:
                    nl = with_pkg(ladder_ms)
                    vals[f"step p{p} + ladder"] = nl or w
            # deployed HRP/Rx = step-aware p50 past the SSIS step, else the single
            # ~90% milestone, else whole-job survival
            vals['deployed HRP/Rx'] = (vals['step p50, no ladder'] if past_ssis
                                       else (with_pkg(single_ms) or w))
            if not past_ssis and with_pkg(ladder_ms):
                hits['ladder'] += 1
            for key, val in vals.items():
                if val is not None:
                    d = (val - run['end']).total_seconds()
                    errs[key].append(abs(d))
                    bias[key].append(d)
                    if not past_ssis:
                        inssis[key].append(abs(d))
            now += timedelta(hours=TICK_HOURS)

    def dur(s):
        neg = s < 0
        m = int(round(abs(s) / 60.0))
        h, m = divmod(m, 60)
        t = f"{h}h{m:02d}m" if h else f"{m}m"
        return ('-' + t) if neg else t

    n_ssis = len(inssis['whole (pre-step-aware)']) or 1
    print(f"  {'estimator':<26} {'med err':>8} {'p90 err':>8} {'bias':>8} | "
          f"{'SSIS med':>9} {'SSIS p90':>9}")
    for k in variants:
        if not errs[k]:
            continue
        a, b = errs[k], inssis[k]
        print(f"  {k:<26} {dur(pct(a,50)):>8} {dur(pct(a,90)):>8} "
              f"{dur(pct(sorted(bias[k]),50)):>8} | "
              f"{dur(pct(b,50)) if b else '-':>9} {dur(pct(b,90)) if b else '-':>9}")
    print(f"  ladder rung available on {round(100*hits['ladder']/n_ssis)}% of "
          f"{n_ssis} in-SSIS ticks")


for server, name, pkg in JOBS:
    try:
        run_job(server, name, pkg)
    except Exception as exc:
        print(f"  !! {name}: {type(exc).__name__}: {exc}")
