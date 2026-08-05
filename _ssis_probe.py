"""Ad-hoc probe: how much progressive signal SSISDB carries INSIDE the dominant
SSIS step of each Aetna job. The per-step ETA can only tighten the T-SQL tail;
the SSIS package is 41-99% of wall-clock, so this asks whether the package's own
task-completion log (catalog.executable_statistics) can pace that step too.
Throwaway analysis helper -- not part of any scheduled job."""
import subprocess

SRV = 'TRGETLPROD2'


def sql(q, db='SSISDB', server=SRV):
    out = subprocess.run(['sqlcmd', '-S', server, '-d', db, '-E', '-W', '-h', '-1',
                          '-s', '|', '-Q', 'SET NOCOUNT ON; ' + q],
                         capture_output=True, text=True, timeout=300)
    rows = []
    for line in out.stdout.splitlines():
        line = line.rstrip()
        if not line or set(line) <= set('-|') or 'rows affected' in line:
            continue
        rows.append([c.strip() for c in line.split('|')])
    return rows, out.stderr.strip()


print("=== packages executed in the last 120 days (name, runs, med minutes) ===")
rows, err = sql(
    "SELECT package_name, COUNT(*), "
    " AVG(DATEDIFF(second,start_time,end_time)/60) "
    "FROM catalog.executions WHERE start_time>DATEADD(day,-120,GETDATE()) "
    "AND end_time IS NOT NULL AND status=7 "
    "GROUP BY package_name HAVING COUNT(*)>=5 ORDER BY 3 DESC;")
if err:
    print("ERR", err)
for r in rows:
    print("  ", " | ".join(r))

for pkg in ('AetnaRx_MasterLoad_Claims_And_Eligibility.dtsx', 'AetnaRCE_Load.dtsx'):
    print(f"\n=== {pkg}: executables by avg end-fraction (last 8 successful runs) ===")
    rows, err = sql(
        f"DECLARE @pkg sysname=N'{pkg}'; "
        ";WITH ex AS (SELECT TOP 8 execution_id, start_time, end_time, "
        "  DATEDIFF(second,start_time,end_time) AS total_sec "
        "  FROM catalog.executions WHERE package_name=@pkg AND status=7 "
        "  AND end_time IS NOT NULL ORDER BY execution_id DESC), "
        "f AS (SELECT es.executable_id, "
        "  CAST(DATEDIFF(second,ex.start_time,es.end_time) AS float)/NULLIF(ex.total_sec,0) AS frac, "
        "  DATEDIFF(second,es.end_time,ex.end_time) AS tail "
        "  FROM catalog.executable_statistics es JOIN ex ON ex.execution_id=es.execution_id "
        "  WHERE ex.total_sec>600) "
        "SELECT f.executable_id, e.executable_name, COUNT(*) n, "
        "  CAST(AVG(f.frac) AS decimal(5,3)), CAST(ISNULL(STDEV(f.frac),9) AS decimal(5,3)), "
        "  AVG(f.tail)/60 AS tail_min, "
        "  (MAX(f.tail)-MIN(f.tail))/60 AS tail_spread_min "
        "FROM f LEFT JOIN (SELECT DISTINCT executable_id, executable_name "
        "   FROM catalog.executables WHERE package_name=@pkg) e ON e.executable_id=f.executable_id "
        "GROUP BY f.executable_id, e.executable_name HAVING COUNT(*)>=5 "
        "ORDER BY AVG(f.frac);")
    if err:
        print("ERR", err)
    print(f"  {'exec_id':>8} {'n':>3} {'frac':>6} {'sd':>6} {'tail_m':>7} {'spread_m':>9}  name")
    for r in rows:
        if len(r) >= 7:
            print(f"  {r[0]:>8} {r[2]:>3} {r[3]:>6} {r[4]:>6} {r[5]:>7} {r[6]:>9}  {r[1]}")
