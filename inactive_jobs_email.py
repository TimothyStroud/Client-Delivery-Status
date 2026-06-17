"""
Email report of all inactive RAMP jobs, grouped by Last Run Date.
Columns: Job Name, Last Updated On, Last Updated By, Last Run, Toggle Notes
"""
import subprocess, json, sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta

sys.path.insert(0, r'C:\Users\tls2\.claude\projects\H--')
import send_via_outlook

TO      = 'RDPOperations@Machinify.com'
FROM    = 'DataOperations@machinify.com'
SUBJECT = 'RAMP Inactive Jobs Report'
ADO_URL = 'https://devops.ado.rawlingslou.prod/TFS2012/AppDev/_workitems/edit/{}'

GROUPS = [
    ('Past Week',         0,   7),
    ('Past 1–4 Weeks',    7,   28),
    ('Past 1–3 Months',   28,  90),
    ('Older than 3 Months', 90, None),
]


def curl_json(url):
    result = subprocess.run(
        ['curl', '-s', '--ntlm', '-u', ':', url],
        capture_output=True, text=True
    )
    return json.loads(result.stdout)


def fetch_latest_note(job_id):
    try:
        data = curl_json(f'http://ramp/api/Ramp/Notes/Job/{job_id}')
        notes = data.get('Data', [[]])[0]
        if notes:
            notes_sorted = sorted(notes, key=lambda n: n.get('NoteDate', ''), reverse=True)
            n = notes_sorted[0]
            return job_id, n.get('Message', '').strip(), n.get('NoteDate', ''), n.get('TfsId')
        return job_id, '', '', None
    except Exception:
        return job_id, '', '', None


def fmt_date(dt_str):
    if not dt_str:
        return ''
    try:
        return datetime.fromisoformat(dt_str).strftime('%Y-%m-%d')
    except Exception:
        return dt_str


def strip_domain(user):
    if user and '\\' in user:
        return user.split('\\', 1)[1]
    return user or ''


def age_days(dt_str, now):
    if not dt_str:
        return None
    try:
        return (now - datetime.fromisoformat(dt_str)).days
    except Exception:
        return None


def build_group_table(rows):
    th_style = 'background:#2c5f8a;color:#fff;padding:8px 12px;text-align:left;white-space:nowrap;'
    td_style = 'padding:7px 12px;border-bottom:1px solid #e0e0e0;vertical-align:top;'
    td_alt   = 'padding:7px 12px;border-bottom:1px solid #e0e0e0;vertical-align:top;background:#f5f8fc;'
    headers  = ['Job Name', 'Last Updated On', 'Last Updated By', 'Last Run', 'Toggle Notes']
    thead    = ''.join(f'<th style="{th_style}">{h}</th>' for h in headers)
    tbody_rows = []
    for i, r in enumerate(rows):
        style = td_alt if i % 2 else td_style
        cells = []
        for col_idx, v in enumerate(r):
            if col_idx == 0 and 'snap' in v.lower():
                cells.append(f'<td style="{style}"><span style="background:#ffe0b2;padding:1px 4px;border-radius:3px;font-weight:600;">{v}</span></td>')
            else:
                cells.append(f'<td style="{style}">{v}</td>')
        tbody_rows.append(f'<tr>{"".join(cells)}</tr>')
    return (
        f'<table style="border-collapse:collapse;width:100%;min-width:700px;margin-bottom:28px;">'
        f'<thead><tr>{thead}</tr></thead>'
        f'<tbody>{"".join(tbody_rows)}</tbody>'
        f'</table>'
    )


def build_html(groups_data, total):
    group_html = []
    gh_style        = ('font-family:Segoe UI,Arial,sans-serif;font-size:14px;font-weight:700;'
                       'color:#2c5f8a;margin:20px 0 6px 0;padding-bottom:4px;'
                       'border-bottom:2px solid #2c5f8a;')
    gh_style_recent = ('font-family:Segoe UI,Arial,sans-serif;font-size:14px;font-weight:700;'
                       'color:#c0392b;margin:20px 0 6px 0;padding-bottom:4px;'
                       'border-bottom:2px solid #c0392b;')
    for i, (label, rows) in enumerate(groups_data):
        if not rows:
            continue
        style = gh_style_recent if i == 0 else gh_style
        group_html.append(
            f'<p style="{style}">{label} <span style="font-weight:400;font-size:12px;color:#555;">({len(rows)} job(s))</span></p>'
            + build_group_table(rows)
        )

    return f"""
<html><body style="font-family:Segoe UI,Arial,sans-serif;font-size:13px;color:#222;">
<p style="margin-bottom:8px;">
  <a href="http://ramp/Ramp/Job" style="font-size:13px;">View Inactive Jobs in RAMP</a>
</p>
<p style="margin-bottom:16px;">
  <strong>{total} inactive job(s)</strong> as of {datetime.now().strftime('%Y-%m-%d %H:%M')},
  grouped by Last Run Date. Jobs whose Last Updated date is before 2026 are excluded from this report.
</p>
<div style="margin:0 0 20px 0;padding:12px 16px;background:#fff8e1;border-left:5px solid #f39c12;border-radius:4px;font-size:14px;color:#222;">
  <span style="font-weight:700;color:#b9770e;">Action Requested:</span>
  Please take a few minutes to review the jobs for the clients you support and re-activate any that should be running again. If a job should remain inactive, no action is needed.
</div>
{''.join(group_html)}
</body></html>
"""


def main(to_override=None, from_override=None):
    print('Fetching job list...')
    data = curl_json('http://ramp/api/Ramp/Job/List')
    jobs = data['Data'][0]

    inactive = [j for j in jobs if j.get('Enabled') == 0]

    def updated_in_2026_or_later(j):
        last = (j.get('PreviousJobConfig') or {}).get('TrackLastUpdated') or ''
        if not last:
            return False
        try:
            return datetime.fromisoformat(last).year >= 2026
        except Exception:
            return False

    inactive = [j for j in inactive if updated_in_2026_or_later(j)]
    print(f'Found {len(inactive)} inactive jobs updated in 2026+. Fetching notes...')

    notes_map = {}
    with ThreadPoolExecutor(max_workers=10) as pool:
        futures = {pool.submit(fetch_latest_note, j['JobId']): j['JobId'] for j in inactive}
        for f in as_completed(futures):
            jid, note, note_date, tfs_id = f.result()
            notes_map[jid] = (note, note_date, tfs_id)

    def sort_key(j):
        run = j.get('LatestJobRun') or {}
        return run.get('StartDate') or ''

    inactive.sort(key=sort_key, reverse=True)

    now = datetime.now()

    # Bucket jobs into groups
    buckets = {label: [] for label, _, _ in GROUPS}
    for j in inactive:
        pc  = j.get('PreviousJobConfig') or {}
        run = j.get('LatestJobRun') or {}
        last_updated_str = pc.get('TrackLastUpdated') or ''
        last_run_str = run.get('StartDate') or ''
        days = age_days(last_run_str, now)
        note_text, note_date, tfs_id = notes_map.get(j['JobId'], ('', '', None))
        try:
            note_within_month = note_date and (now - datetime.fromisoformat(note_date)).days <= 30
        except Exception:
            note_within_month = False
        if note_within_month and note_text:
            cell = note_text
            if tfs_id:
                link = f'<a href="{ADO_URL.format(tfs_id)}">ADO {tfs_id}</a>'
                cell = f'{note_text} [{link}]'
        else:
            cell = ''
        row = [
            j.get('JobName', ''),
            fmt_date(last_updated_str),
            strip_domain(pc.get('TrackChangedBy')),
            fmt_date(last_run_str),
            cell,
        ]
        placed = False
        for label, min_days, max_days in GROUPS:
            if days is None:
                continue
            if max_days is None:
                if days >= min_days:
                    buckets[label].append(row)
                    placed = True
                    break
            elif min_days <= days < max_days:
                buckets[label].append(row)
                placed = True
                break
        if not placed:
            buckets[GROUPS[-1][0]].append(row)

    groups_data = [(label, buckets[label]) for label, _, _ in GROUPS]
    html = build_html(groups_data, len(inactive))

    recipient = to_override or TO
    sender    = from_override or FROM
    print(f'Sending email to {recipient} from {sender}...')
    result = send_via_outlook.send(recipient, SUBJECT, html, from_address=sender)
    print(f'Result: {result}')


if __name__ == '__main__':
    import sys as _sys
    override = _sys.argv[1] if len(_sys.argv) > 1 else None
    main(to_override=override, from_override=None)
