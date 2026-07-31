"""
AetnaRx Eligibility file report -> #data-operations-aetna-updates (Fridays 10am).

Reads [AetnaRx].[etl].[tape] on TRGETL2 for TableID 200 (Eligibility) and, for
each of the 5 file families, reports:
  - the most recent file DATE (parsed out of the FileName, NOT the load date)
  - missing file dates per month, Jan 2026 forward

Missing dates are only reported for dates that fall BEFORE that family's most
recent loaded file (per user: "only note missing files that should have been
loaded before the most recent loaded file"), so a feed that has simply stopped
delivering does not spray a month of false misses -- its stale last-file date is
the signal instead.

Cadence is derived empirically per family from the trailing 120 days of its own
history (which weekdays it actually delivers on), so a weekly feed like
RXBOR-CVSCMK.ELIGCMP (Tuesdays) is not judged against a 7-day expectation.

Output: a single `SLACK|<text>` line with literal '\n' escapes, consumed by
aetnarx_elig_webhook_post.py. The aetna-updates Workflow Builder webhook renders
:emoji: shortcodes ONLY -- no bold/italic/code -- so the body is plain text.

  --force   ignore the once-per-day dedupe (ad-hoc post)
  --dry     print the human-readable message instead of the SLACK| line
"""
import sys, os, json, collections, subprocess
import datetime as dt

BASE = r'C:\Users\tls2\.claude\projects\H--'
STATE_FILE = os.path.join(BASE, 'aetnarx_elig_report_state.json')
SERVER = 'TRGETL2'   # live loader DB. TRGINTP3 (the server in the user's original
                     # query) is a downstream copy that lags ~1 day -- verified
                     # 2026-07-31: INTP3 had 16688 Eligibility rows / latest EIE
                     # 7/3, ETL2 had 16697 / latest 7/4, matching the user's sample.
START = dt.date(2026, 1, 1)          # report window start (per user: Jan 2026 forward)
CADENCE_WINDOW = 120                 # days of own history used to infer cadence
MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July',
          'August', 'September', 'October', 'November', 'December']

# Display order + label, per the user's sample layout.
FAMILIES = [
    ('APCCF',                'Aetna.APCCF.Daily'),
    ('EIE-CVSCMK.MBRELIG',   'EIE-CVSCMK.MBRELIG'),
    ('EIE-CVSCMK.MBRMDCR',   'EIE-CVSCMK.MBRMDCR'),
    ('RXBOR-CVSCMK.MBRELIG', 'RXBOR-CVSCMK.MBRELIG'),
    ('RXBOR-CVSCMK.ELIGCMP', 'RXBOR-CVSCMK.ELIGCMP'),
]

QUERY = r"""SET NOCOUNT ON;
WITH b AS (
    SELECT REVERSE(SUBSTRING(REVERSE(t.FileName), 1,
               CHARINDEX('\', REVERSE(t.FileName)) - 1)) AS fn
    FROM [AetnaRx].[etl].[tape] T (nolock)
    JOIN [AetnaRx].[config].[Table] F (nolock) ON t.TableID = f.TableID
    WHERE t.TableID IN (200)
),
p AS (
    SELECT CASE WHEN fn LIKE '%APCCF%' THEN 'APCCF'
                ELSE LEFT(fn, CHARINDEX('.', fn + '.', CHARINDEX('.', fn + '.') + 1) - 1)
           END AS fam,
           CASE WHEN fn LIKE '%APCCF%'
                THEN SUBSTRING(fn, CHARINDEX('ELIGIBILITY.', fn) + 12, 8)
                ELSE REPLACE(SUBSTRING(fn, CHARINDEX('.', fn, CHARINDEX('.', fn) + 1) + 1, 10), '-', '')
           END AS d
    FROM b
)
SELECT fam, d, COUNT(*) AS n
FROM p
WHERE d LIKE '20[12][0-9][01][0-9][0-3][0-9]'
GROUP BY fam, d
ORDER BY fam, d"""


def fetch():
    """{family: {date: file_count}} keyed off the date inside the FileName."""
    out = subprocess.run(['sqlcmd', '-S', SERVER, '-E', '-W', '-s', '~', '-Q', QUERY],
                         capture_output=True, text=True, timeout=300)
    if out.returncode != 0:
        raise RuntimeError(f"sqlcmd failed: {(out.stderr or out.stdout).strip()[:300]}")
    fams = collections.defaultdict(dict)
    for line in out.stdout.splitlines():
        parts = line.split('~')
        if len(parts) != 3 or parts[0].strip() in ('fam', '---'):
            continue
        fam, d, n = (x.strip() for x in parts)
        try:
            day = dt.date(int(d[:4]), int(d[4:6]), int(d[6:8]))
        except ValueError:
            continue
        fams[fam][day] = int(n)
    return fams


def cadence(days):
    """Infer which weekdays a family delivers on, from its recent own history.

    Returns (set_of_weekdays, human_label). Uses the trailing CADENCE_WINDOW days
    ending at the family's last file so a cadence change (or a weekly feed) is
    respected rather than assuming daily.
    """
    if not days:
        return set(range(7)), 'Daily'
    last = max(days)
    recent = [d for d in days if d > last - dt.timedelta(days=CADENCE_WINDOW)]
    dows = {d.weekday() for d in recent}
    if len(dows) >= 6:
        return set(range(7)), 'Daily'
    if dows == {0, 1, 2, 3, 4}:
        return dows, 'Weekdays'
    names = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    if len(dows) == 1:
        return dows, 'Weekly on ' + names[next(iter(dows))]
    return dows, 'Weekly on ' + ', '.join(names[d] for d in sorted(dows))


def missing_by_month(days, dows):
    """{(year, month): [missing dates]} for expected days in [START, last_file)."""
    if not days:
        return {}
    last, first = max(days), min(days)
    lo = max(START, first)
    out = collections.defaultdict(list)
    d = lo
    while d < last:
        if d.weekday() in dows and d not in days:
            out[(d.year, d.month)].append(d)
        d += dt.timedelta(days=1)
    return out


def md(d):
    return f"{d.month}/{d.day}"


def mdy(d):
    return f"{d.month}/{d.day}/{d:%y}"


def build(fams, today):
    lines = [f"AetnaRx Eligibility File Update - {today.strftime('%a')} {mdy(today)}", ""]
    total_missing = 0
    for key, label in FAMILIES:
        days = fams.get(key, {})
        if not days:
            lines.append(f":question: {label} - no files found")
            lines.append("")
            continue
        dows, cad_label = cadence(days)
        last = max(days)
        head = f"{label} - {mdy(last)}"
        if cad_label != 'Daily':
            head += f" - {cad_label}"
        lines.append(head)

        miss = missing_by_month(days, dows)
        if miss:
            for (y, m) in sorted(miss, reverse=True):
                ds = miss[(y, m)]
                total_missing += len(ds)
                lines.append(f"  :round_pushpin: Missing {MONTHS[m - 1]} Files: "
                             + ', '.join(md(d) for d in ds))
        else:
            lines.append("  :white_small_square: No missing files")
        lines.append("")
    lines.append(f"Source: {SERVER} AetnaRx.etl.tape TableID 200 | dates from FileName "
                 f"| missing window {mdy(START)} forward, up to each feed's latest file")
    return '\n'.join(lines).rstrip(), total_missing


def load_state():
    try:
        return json.load(open(STATE_FILE, encoding='utf-8'))
    except Exception:
        return {}


def save_state(state):
    tmp = STATE_FILE + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(state, f, indent=2)
    os.replace(tmp, STATE_FILE)


def main():
    force = '--force' in sys.argv
    dry = '--dry' in sys.argv
    today = dt.date.today()

    # Two-phase commit: the poster re-invokes with --commit only after the Slack
    # post succeeds, so a failed post is retried rather than silently swallowed.
    if '--commit' in sys.argv:
        save_state({'last_date': today.isoformat()})
        print('committed ' + today.isoformat())
        return 0

    fams = fetch()
    msg, _ = build(fams, today)

    if dry:
        print(msg)
        return 0

    if not force and load_state().get('last_date') == today.isoformat():
        print('SKIP: already posted today')
        return 0

    print('SLACK|' + msg.replace('\n', '\\n'))
    return 0


if __name__ == '__main__':
    sys.exit(main())
