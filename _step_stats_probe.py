"""Ad-hoc probe: per-step duration statistics from msdb.dbo.sysjobhistory (the
table behind SSMS's Job Activity Monitor -> View History / Log File Viewer) for
the four Aetna SQL Agent jobs, to judge how much a per-step-average ETA can
tighten each digest. Throwaway analysis helper -- not part of any scheduled job."""
import subprocess
import sys

JOBS = [
    ("TRGETL2", "ETL AetnaHRP MasterLoad"),
    ("TRGETL2", "ETL AetnaRx MasterLoad Claims And Eligibility"),
    ("TRGETL2", "SSIS AetnaRCE Daily Process"),
    ("TRGETL4", "ETL NCStateAetna MasterLoad"),
]
DAYS = 120


def run(server, q):
    out = subprocess.run(['sqlcmd', '-S', server, '-E', '-W', '-h', '-1', '-s', '|', '-Q', q],
                         capture_output=True, text=True, timeout=300)
    return out.stdout


def dur(sec):
    m = int(round(sec / 60.0))
    h, m = divmod(m, 60)
    return f"{h}h{m:02d}m" if h else f"{m}m"


for server, name in JOBS:
    print(f"\n===== {name}  ({server}) =====")
    q = ("SET NOCOUNT ON; "
         f"DECLARE @jid uniqueidentifier=(SELECT job_id FROM msdb.dbo.sysjobs WHERE name=N'{name}'); "
         "SELECT s.step_id, s.subsystem, s.step_name FROM msdb.dbo.sysjobsteps s "
         "WHERE s.job_id=@jid ORDER BY s.step_id;")
    steps = {}
    for line in run(server, q).splitlines():
        p = [x.strip() for x in line.split('|')]
        if len(p) >= 3 and p[0].isdigit():
            steps[int(p[0])] = (p[1], p[2])

    q2 = ("SET NOCOUNT ON; "
          f"DECLARE @jid uniqueidentifier=(SELECT job_id FROM msdb.dbo.sysjobs WHERE name=N'{name}'); "
          "WITH h AS (SELECT instance_id, step_id, run_status, "
          " (run_duration/10000)*3600+((run_duration/100)%100)*60+(run_duration%100) AS secs "
          " FROM msdb.dbo.sysjobhistory WITH (NOLOCK) WHERE job_id=@jid "
          f" AND run_date>=CONVERT(int,CONVERT(varchar(8),DATEADD(day,-{DAYS},GETDATE()),112))) "
          "SELECT step_id, COUNT(*), MIN(secs), MAX(secs), AVG(secs*1.0), "
          " (SELECT TOP 1 secs FROM (SELECT secs, ROW_NUMBER() OVER (ORDER BY secs) rn, "
          "   COUNT(*) OVER () c FROM h h2 WHERE h2.step_id=h.step_id AND h2.run_status=1) z "
          "   WHERE rn=(c+1)/2) AS med "
          "FROM h WHERE run_status=1 GROUP BY step_id ORDER BY step_id;")
    rows = []
    for line in run(server, q2).splitlines():
        p = [x.strip() for x in line.split('|')]
        if len(p) >= 6 and p[0].lstrip('-').isdigit():
            try:
                rows.append((int(p[0]), int(p[1]), int(p[2]), int(p[3]),
                             float(p[4]), int(p[5])))
            except ValueError:
                continue
    total_med = sum(r[5] for r in rows if r[0] != 0) or 1
    print(f"{'stp':>3} {'n':>4} {'min':>7} {'med':>8} {'avg':>8} {'max':>8} {'%job':>6}  subsystem  name")
    for sid, n, mn, mx, avg, med in rows:
        label = 'JOB TOTAL' if sid == 0 else steps.get(sid, ('?', '?'))[1]
        sub = '' if sid == 0 else steps.get(sid, ('?', '?'))[0]
        pct = '' if sid == 0 else f"{100.0*med/total_med:5.1f}%"
        print(f"{sid:>3} {n:>4} {dur(mn):>7} {dur(med):>8} {dur(avg):>8} {dur(mx):>8} {pct:>6}  {sub:<9}  {label}")
