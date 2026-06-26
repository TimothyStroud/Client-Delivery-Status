"""
Combined Aetna RCE tick — merges the fail-only monitor and the status digest into
ONE Claude cron to cut token use (added 2026-06-26 per user).

ONE cron `7 6-16/2 * * *` runs this every 2 hours, 6am-4pm. Each tick:
  - ALWAYS runs the fail-only monitor (ramp_aetna_slack_monitor.py). Any RCE-FAILED
    alert is re-emitted as 'POST_SUPPORT|<text>' -> post to #team-rdp-operations-support
    (C09EPLQL2D9) only.
  - At a DIGEST SLOT (hour 8/12/16, Tue-Fri) also runs the status digest
    (ramp_aetna_status_digest.py). Its single line is re-emitted as
    'POST_BOTH|<text>' -> post to BOTH channels (C09EPLQL2D9 + C09G5BQBL49).
    The digest's own dedupe + both-succeeded skip still apply.

--commit  -> commits the monitor state. The cron runs this ONLY after a POST_SUPPORT
             line has posted successfully (two-phase, so a failed post retries next
             tick). Digest needs no commit (it self-dedupes).

Channel routing + the two-phase monitor commit are driven by the cron prompt.
"""
import sys, subprocess
from datetime import datetime

BASE = r'C:\Users\tls2\.claude\projects\H--'
PY = sys.executable
MONITOR = BASE + r'\ramp_aetna_slack_monitor.py'
DIGEST = BASE + r'\ramp_aetna_status_digest.py'

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
                print('POST_BOTH|' + line[len('SLACK|'):])
                posted_any = True

    if not posted_any:
        print('NO_EVENTS')


if __name__ == '__main__':
    main()
