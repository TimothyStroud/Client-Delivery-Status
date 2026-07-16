"""
Combined AetnaRx Claim tick — mirrors aetnahrp_tick.py.

The Windows task "AetnaRx Tick" runs this every 2 hours, 6am-4pm. Each tick:
  - ALWAYS runs the fail-only monitor (ramp_aetnarx_slack_monitor.py). Any
    AetnaRx Claim FAILED alert is re-emitted as 'POST_SUPPORT|<text>' -> post to
    #data-operations-aetna-updates, then --commit.
  - At a DIGEST SLOT (hour 8/12/16, Tue-Fri) also runs the status digest
    (ramp_aetnarx_status_digest.py). Its single line is re-emitted as
    'POST_DIGEST|<text>' -> post to the SAME channel. The digest's own dedupe +
    fully-green-skip apply; it needs no commit (self-dedupes).

--commit  -> commits the monitor state. The poster runs this ONLY after a
             POST_SUPPORT line posts OK (two-phase, so a failed post retries).
"""
import sys, subprocess
from datetime import datetime

BASE = r'C:\Users\tls2\.claude\projects\H--'
PY = sys.executable
MONITOR = BASE + r'\ramp_aetnarx_slack_monitor.py'
DIGEST = BASE + r'\ramp_aetnarx_status_digest.py'

DIGEST_HOURS = {8, 12, 16}        # 8am / 12pm / 4pm
DIGEST_DOW = {1, 2, 3, 4}         # Tue-Fri (Mon=0 .. Sun=6)


def run(args):
    return subprocess.run([PY] + args, capture_output=True, text=True)


def main():
    if '--commit' in sys.argv:
        r = run([MONITOR, '--commit'])
        sys.stdout.write(r.stdout)
        return

    posted_any = False

    # 1) Fail-only monitor — every tick.
    r = run([MONITOR])
    for line in r.stdout.splitlines():
        if line.startswith('SLACK|'):
            print('POST_SUPPORT|' + line[len('SLACK|'):])
            posted_any = True

    # 2) Status digest — only at a digest slot (hour 8/12/16, Tue-Fri).
    now = datetime.now()
    if now.hour in DIGEST_HOURS and now.weekday() in DIGEST_DOW:
        rd = run([DIGEST])
        for line in rd.stdout.splitlines():
            if line.startswith('SLACK|'):
                print('POST_DIGEST|' + line[len('SLACK|'):])
                posted_any = True

    if not posted_any:
        print('NO_EVENTS')


if __name__ == '__main__':
    main()
