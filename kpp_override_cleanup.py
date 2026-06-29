"""
One-shot headless cleanup (zero Claude tokens): removes the temporary
KaiserPrePayCOB 6/29 "L" MANUAL_OVERRIDES pin from rdp_client_delivery_status.py
once the 'Kaiser Pareo Prepay 0110 Load' (JobId 10721) has SUCCEEDED today, then
regenerates the report and deletes its own scheduled task.

Per user 2026-06-29: "Remove the KaiserPrePayCOB override once it loads."

Idempotent: if the override block (between the KPP_OVERRIDE_START/END sentinels)
is already gone, it just removes the task and exits. If the load hasn't succeeded
yet, it does nothing and waits for the next run.
"""
import os, re, subprocess, sys
from datetime import datetime

BASE = r"C:\Users\tls2\.claude\projects\H--"
REPORT = os.path.join(BASE, "rdp_client_delivery_status.py")
PY = r"C:\Program Files\Python311\python.exe"
LOG = r"H:\kpp_override_cleanup.log"
TASK = "KPP Override Cleanup"
LOAD_JOBID = 10721  # Kaiser Pareo Prepay 0110 Load
START_RE = re.compile(r"^[ \t]*# KPP_OVERRIDE_START\b.*$")
END_RE = re.compile(r"^[ \t]*# KPP_OVERRIDE_END\b.*$")


def log(msg):
    line = f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}"
    try:
        with open(LOG, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass
    print(line)


def delete_task():
    subprocess.run(["schtasks", "/Delete", "/TN", TASK, "/F"],
                   capture_output=True, text=True)
    log(f"deleted scheduled task '{TASK}'")


def block_present(lines):
    return any(START_RE.match(l) for l in lines)


def loaded_today():
    """True if Pareo Prepay 0110 Load succeeded today."""
    q = ("SET NOCOUNT ON; SELECT TOP 1 1 FROM [RAMP].[ramp].[Queue] "
         f"WHERE JobId={LOAD_JOBID} AND EndDate IS NOT NULL "
         "AND CAST(StartDate AS date)=CAST(GETDATE() AS date) "
         "AND (Status LIKE 'Success%' OR Status='Resolved')")
    r = subprocess.run(["sqlcmd", "-S", "TRGUTIL10", "-d", "RAMP", "-E",
                        "-Q", q, "-W", "-h", "-1"],
                       capture_output=True, text=True)
    return any(ln.strip() == "1" for ln in r.stdout.splitlines())


def remove_block():
    with open(REPORT, encoding="utf-8") as f:
        lines = f.readlines()
    out, skipping = [], False
    for l in lines:
        if not skipping and START_RE.match(l):
            skipping = True
            continue
        if skipping:
            if END_RE.match(l):
                skipping = False
            continue
        out.append(l)
    with open(REPORT, "w", encoding="utf-8") as f:
        f.writelines(out)


def main():
    with open(REPORT, encoding="utf-8") as f:
        lines = f.readlines()
    if not block_present(lines):
        log("override block already gone; nothing to do")
        delete_task()
        return 0
    if not loaded_today():
        log("Pareo Prepay 0110 Load not yet successful today; waiting")
        return 0
    log("load succeeded today -> removing KaiserPrePayCOB 6/29 override block")
    remove_block()
    r = subprocess.run([PY, REPORT], capture_output=True, text=True, timeout=1500)
    ok = "[done] Wrote" in r.stdout
    log(f"regenerated report: {'ok' if ok else 'CHECK OUTPUT'}")
    if not ok:
        log("report tail: " + r.stdout.strip()[-300:])
    delete_task()
    return 0


if __name__ == "__main__":
    sys.exit(main())
