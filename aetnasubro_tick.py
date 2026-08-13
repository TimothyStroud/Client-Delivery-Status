"""
Combined Aetna Subro tick — mirrors aetnarx_tick.py / aetnahrp_tick.py.

The Windows task "AetnaSubro Tick" runs this every 2 hours, 4:37am-10:37pm — ten
minutes after the AetnaRx Tick (:27), per user 2026-08-13. Each tick:
  - ALWAYS runs the fail-only monitor (ramp_aetnasubro_slack_monitor.py). Any
    'Aetna 0110 Subro Load' FAILED alert is re-emitted as 'POST_SUPPORT|<text>'
    -> post to #data-operations-aetna-updates, then --commit.
  - At a DIGEST SLOT also runs the status digest (ramp_aetnasubro_status_digest.py).
    Its single line is re-emitted as 'POST_DIGEST|<text>' -> the SAME channel. The
    digest's own cross-run + content dedupe apply; it needs no commit.

Subro is a MONTHLY load, so on the ~29 quiet days a month the digest's content
dedupe means these ticks post nothing at all.

--commit  -> commits the monitor state. The poster runs this ONLY after a
             POST_SUPPORT line posts OK (two-phase, so a failed post retries).
"""
import sys, subprocess
from datetime import datetime

BASE = r'C:\Users\tls2\.claude\projects\H--'
PY = sys.executable
MONITOR = BASE + r'\ramp_aetnasubro_slack_monitor.py'
DIGEST = BASE + r'\ramp_aetnasubro_status_digest.py'

DIGEST_HOURS = {4, 6, 8, 10, 12, 14, 16, 18, 20, 22}  # every 2h 4am-10pm, matching
                                                      # the other Aetna digests
DIGEST_DOW = {0, 1, 2, 3, 4, 5, 6}  # every day incl. weekends
EVENING_FROM = 17                 # 5pm+: evening extension. Past the last normal
                                  # slot, keep running the digest --evening ANY day
                                  # so a load finishing late still gets its
                                  # Successful post. The digest self-gates on a real
                                  # load being active/done today, so no-load
                                  # evenings stay silent.


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

    # 2) Status digest — at a normal slot, or as an evening extension (>= 5pm) so a
    #    load finishing after the last daytime slot still posts its Successful line.
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
                print('POST_DIGEST|' + line[len('SLACK|'):])
                posted_any = True

    if not posted_any:
        print('NO_EVENTS')


if __name__ == '__main__':
    main()
