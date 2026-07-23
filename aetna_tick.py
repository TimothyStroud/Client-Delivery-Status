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

DIGEST_HOURS = {4, 6, 8, 10, 12, 14, 16, 18, 20, 22}  # every 2h 4am-10pm (per user 2026-07-23; content-dedupe suppresses repeats)
DIGEST_DOW = {0, 1, 2, 3, 4, 5, 6}  # every day incl. weekends (per user 2026-07-19)
EVENING_FROM = 17                 # 5pm+: evening extension (per user 2026-07-17).
                                  # Past the last normal slot, keep running the
                                  # digest --evening ANY day so a load finishing
                                  # after 4pm still gets its Successful post. The
                                  # digest self-gates on a real load being active/
                                  # done today, so no-load evenings stay silent.


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

    # 2) Status digest — at a normal slot (hour 8/12/16, Tue-Fri) or, on any day,
    #    as an evening extension (>= 5pm) so a load finishing after the last daytime
    #    slot still posts its Successful line. The digest gates --evening itself.
    now = datetime.now()
    digest_args = None
    if now.hour in DIGEST_HOURS and now.weekday() in DIGEST_DOW:
        digest_args = [DIGEST]
    elif now.hour >= EVENING_FROM:
        digest_args = [DIGEST, '--evening']
    if digest_args:
        rd = run(digest_args)
        for line in rd.stdout.splitlines():
            if line.startswith('SLACK|'):
                print('POST_BOTH|' + line[len('SLACK|'):])
                posted_any = True

    if not posted_any:
        print('NO_EVENTS')


if __name__ == '__main__':
    main()
