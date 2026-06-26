"""
EmblemFacets Response-File SLA monitor.

Monthly process (~15th). Pipeline in RAMP:
    Emblem Facets 0100 Stage -> Emblem Facets 0110 Load -> SFTP Emblem Upload

SLA clock starts when 'Emblem Facets 0110 Load' STARTS. Each of the 5 response
files must be uploaded within 3 days of that start. This script:
  1. Finds the most recent 'Emblem Facets 0110 Load' start (the SLA anchor).
  2. Checks the upload archive for the 5 current-cycle RESP_ files.
  3. ONLY when all 5 are present, emails the final SLA report
     (from DataOperations@machinify.com to RDPOperations@machinify.com),
     coloring in RED any file uploaded after anchor + 3 days.
  4. Records the cycle in a state file so it sends exactly once per cycle.

Run with --status to print current state without sending.
Run with --force to send for the current cycle even if already sent.
"""
import sys, os, json, re, subprocess
from datetime import datetime, timedelta

BASE = r'C:\Users\tls2\.claude\projects\H--'
sys.path.insert(0, BASE)
from send_via_outlook import send

ARCHIVE = r'\\trgllc\Shares\RawlingsOutbound\DIG\EmblemFacets\Archive'
STATE_FILE = os.path.join(BASE, 'emblem_sla_state.json')

FROM_ADDR = 'DataOperations@machinify.com'
TO_ADDR   = 'RDPOperations@machinify.com'

SLA_DAYS = 3
LOAD_JOB_NAME = 'Emblem Facets 0110 Load'

# Only run during the monthly window (per user 2026-06-26): the 15th-20th of the
# month, inclusive. The Scheduled Task still fires daily 9am/2pm, but the script
# no-ops on any other day. --force / --status bypass the window.
WINDOW_START, WINDOW_END = 15, 20

# (display label, filename prefix) for the 5 expected response files
EXPECTED = [
    ("RESP_RWGS_CEE_FF_*",                              "RESP_RWGS_CEE_FF_"),
    ("RESP_EH_RAWLINGS_GPPO_MEMELIG_FA_FM01_PROD_*",    "RESP_EH_RAWLINGS_GPPO_MEMELIG_FA_FM01_PROD_"),
    ("RESP_CCI_RAWLINGS_MULTI_MEMELIG_FA_FM01_PROD_*",  "RESP_CCI_RAWLINGS_MULTI_MEMELIG_FA_FM01_PROD_"),
    ("RESP_CCI_RAWLINGS_CCOMM_EMPGRP_FA_FM01_PROD_*",   "RESP_CCI_RAWLINGS_CCOMM_EMPGRP_FA_FM01_PROD_"),
    ("RESP_EH_RAWLINGS_GPPO_GRP_FA_FM01_PROD_*",        "RESP_EH_RAWLINGS_GPPO_GRP_FA_FM01_PROD_"),
]

# Buffer so clock skew between RAMP's StartDate and the file's mtime can't drop
# a legitimately-current-cycle file.
MATCH_BUFFER = timedelta(hours=6)


def get_load_start():
    """Return datetime of the most recent 'Emblem Facets 0110 Load' start."""
    out = subprocess.run(
        ['curl', '-s', '--ntlm', '-u', ':', 'http://ramp/api/Ramp/Queue/List'],
        capture_output=True, text=True)
    data = json.loads(out.stdout)
    d = data['Data']
    items = d[0] if (isinstance(d, list) and d and isinstance(d[0], list)) else d
    best = None
    for i in items:
        xml = i.get('JobXml', '') or ''
        m = re.search(r'jobname="([^"]+)"', xml)
        if m and m.group(1) == LOAD_JOB_NAME and i.get('StartDate'):
            sd = datetime.fromisoformat(i['StartDate'])
            if best is None or sd > best:
                best = sd
    return best


def find_cycle_files(anchor):
    """For each expected prefix, find the newest archive file in this cycle.
    Returns list of dicts: {label, prefix, found, name, uploaded(datetime)}."""
    cutoff = anchor - MATCH_BUFFER
    results = []
    try:
        entries = os.listdir(ARCHIVE)
    except OSError:
        entries = []
    for label, prefix in EXPECTED:
        best = None
        for fn in entries:
            if fn.startswith(prefix):
                fp = os.path.join(ARCHIVE, fn)
                try:
                    mt = datetime.fromtimestamp(os.path.getmtime(fp))
                except OSError:
                    continue
                if mt >= cutoff and (best is None or mt > best[1]):
                    best = (fn, mt)
        if best:
            results.append({'label': label, 'prefix': prefix, 'found': True,
                            'name': best[0], 'uploaded': best[1]})
        else:
            results.append({'label': label, 'prefix': prefix, 'found': False,
                            'name': None, 'uploaded': None})
    return results


def build_body(anchor, deadline, files):
    any_missed = any(f['found'] and f['uploaded'] > deadline for f in files)
    rows = ""
    for f in files:
        missed = f['found'] and f['uploaded'] > deadline
        color = "#c00000" if missed else "#1a7f37"
        weight = "bold" if missed else "normal"
        up = f['uploaded'].strftime('%-m/%-d/%Y %-I:%M %p') if os.name != 'nt' else f['uploaded'].strftime('%m/%d/%Y %I:%M %p')
        status = ("MISSED SLA" if missed else "Within SLA")
        rows += (
            f'<tr style="color:{color};font-weight:{weight}">'
            f'<td style="padding:6px 10px;border:1px solid #ddd;font-family:Consolas,monospace;font-size:12px">{f["label"]}</td>'
            f'<td style="padding:6px 10px;border:1px solid #ddd">{status}</td>'
            f'<td style="padding:6px 10px;border:1px solid #ddd">{up}</td>'
            f'</tr>'
        )

    anchor_s   = anchor.strftime('%m/%d/%Y %I:%M %p')
    deadline_s = deadline.strftime('%m/%d/%Y %I:%M %p')
    banner = ""
    if any_missed:
        banner = ('<p style="color:#c00000;font-weight:bold;font-size:15px">'
                  '&#9888; One or more response files MISSED the 3-day SLA.</p>')
    headline_color = "#c00000" if any_missed else "#1a7f37"

    body = f"""<html><body style="font-family:Calibri,Arial,sans-serif;font-size:14px;color:#222">
<p><b>SLA Update &ndash; EmblemFacets Response Files</b><br>
Monthly cycle &middot; {SLA_DAYS}-day SLA measured from <i>Emblem Facets 0110 Load</i> start &middot; Source: RAMP Dashboard</p>

{banner}
<p style="color:{headline_color};font-weight:bold">All 5 Emblem Facets response files have been uploaded.</p>

<p><b>Load (SLA clock) started:</b> {anchor_s}<br>
<b>3-day SLA deadline:</b> {deadline_s}</p>

<table style="border-collapse:collapse;font-size:13px">
<tr style="background:#f0f0f0;color:#222;font-weight:bold">
  <th style="padding:6px 10px;border:1px solid #ddd;text-align:left">Response File</th>
  <th style="padding:6px 10px;border:1px solid #ddd;text-align:left">SLA Status</th>
  <th style="padding:6px 10px;border:1px solid #ddd;text-align:left">Uploaded</th>
</tr>
{rows}
</table>

<p style="margin-top:14px"><b>Pipeline:</b> Emblem Facets 0100 Stage &rarr; Emblem Facets 0110 Load &rarr; SFTP Emblem Upload.</p>
<p style="color:#666;font-size:12px">Uploaded response files confirmed in
{ARCHIVE} and delivered to Emblem (/Home/Rawlings/Prod/ToEmblem).</p>
<p style="color:#999;font-size:11px">Automated SLA update generated from RAMP.</p>
</body></html>"""
    return body, any_missed


def build_overdue_body(anchor, deadline, files, now):
    """Overdue-alert email: deadline passed and not all 5 are uploaded."""
    missing_n = sum(1 for f in files if not f['found'])
    rows = ""
    for f in files:
        if f['found']:
            late = f['uploaded'] > deadline
            color = "#c00000" if late else "#1a7f37"
            status = "Uploaded LATE" if late else "Uploaded"
            up = f['uploaded'].strftime('%m/%d/%Y %I:%M %p')
        else:
            color = "#c00000"
            status = "NOT UPLOADED - OVERDUE"
            up = "&mdash;"
        rows += (
            f'<tr style="color:{color};font-weight:bold">'
            f'<td style="padding:6px 10px;border:1px solid #ddd;font-family:Consolas,monospace;font-size:12px">{f["label"]}</td>'
            f'<td style="padding:6px 10px;border:1px solid #ddd">{status}</td>'
            f'<td style="padding:6px 10px;border:1px solid #ddd">{up}</td>'
            f'</tr>'
        )

    anchor_s   = anchor.strftime('%m/%d/%Y %I:%M %p')
    deadline_s = deadline.strftime('%m/%d/%Y %I:%M %p')
    now_s      = now.strftime('%m/%d/%Y %I:%M %p')
    overdue_by = now - deadline
    hrs = int(overdue_by.total_seconds() // 3600)

    body = f"""<html><body style="font-family:Calibri,Arial,sans-serif;font-size:14px;color:#222">
<p><b>SLA Update &ndash; EmblemFacets Response Files</b><br>
Monthly cycle &middot; {SLA_DAYS}-day SLA measured from <i>Emblem Facets 0110 Load</i> start &middot; Source: RAMP Dashboard</p>

<p style="color:#c00000;font-weight:bold;font-size:15px">&#9888; SLA BREACH &ndash; {missing_n} of 5 response files NOT uploaded by the 3-day deadline.</p>

<p><b>Load (SLA clock) started:</b> {anchor_s}<br>
<b>3-day SLA deadline:</b> {deadline_s}<br>
<b>As of:</b> {now_s} (overdue by ~{hrs} hours)</p>

<table style="border-collapse:collapse;font-size:13px">
<tr style="background:#f0f0f0;color:#222;font-weight:bold">
  <th style="padding:6px 10px;border:1px solid #ddd;text-align:left">Response File</th>
  <th style="padding:6px 10px;border:1px solid #ddd;text-align:left">Status</th>
  <th style="padding:6px 10px;border:1px solid #ddd;text-align:left">Uploaded</th>
</tr>
{rows}
</table>

<p style="margin-top:14px"><b>Pipeline:</b> Emblem Facets 0100 Stage &rarr; Emblem Facets 0110 Load &rarr; SFTP Emblem Upload.</p>
<p style="color:#666;font-size:12px">Outstanding files have not reached
{ARCHIVE}. A final confirmation will follow once all 5 upload.</p>
<p style="color:#999;font-size:11px">Automated SLA alert generated from RAMP.</p>
</body></html>"""
    return body


def main():
    status_only = '--status' in sys.argv
    force = '--force' in sys.argv

    day = datetime.now().day
    if not (WINDOW_START <= day <= WINDOW_END) and not force and not status_only:
        print(f"Outside monthly check window ({WINDOW_START}th-{WINDOW_END}th); "
              f"today is the {day}. Skipping.")
        return

    anchor = get_load_start()
    if anchor is None:
        print("No 'Emblem Facets 0110 Load' found in RAMP queue. Exiting.")
        return
    deadline = anchor + timedelta(days=SLA_DAYS)
    files = find_cycle_files(anchor)
    found_n = sum(1 for f in files if f['found'])

    cycle_key = anchor.isoformat()
    state = {}
    if os.path.exists(STATE_FILE):
        try:
            state = json.load(open(STATE_FILE))
        except Exception:
            state = {}

    print(f"Load start (anchor): {anchor}")
    print(f"SLA deadline:        {deadline}")
    print(f"Response files found: {found_n}/5")
    for f in files:
        mark = ("OK " if f['found'] else "-- ")
        when = f['uploaded'].strftime('%Y-%m-%d %H:%M') if f['uploaded'] else ''
        late = " (LATE)" if (f['found'] and f['uploaded'] > deadline) else ""
        print(f"  [{mark}] {f['label']}  {when}{late}")

    now = datetime.now()
    print(f"Now:                 {now}  ({'PAST deadline' if now > deadline else 'before deadline'})")

    if status_only:
        return

    # --- Incomplete cycle ---
    if found_n < 5:
        if now <= deadline:
            print("Not all 5 uploaded yet, still within SLA -- no email (waiting).")
            return
        # Past deadline and incomplete -> overdue alert, at most once per calendar day.
        overdue_key = f"{cycle_key}|{now.strftime('%Y-%m-%d')}"
        if state.get('overdue_last_key') == overdue_key and not force:
            print("Overdue alert already sent today for this cycle. Skipping.")
            return
        body = build_overdue_body(anchor, deadline, files, now)
        month = anchor.strftime('%B %Y')
        subject = f"SLA Update - EmblemFacets Response Files ({month}) - SLA BREACH ({5 - found_n} of 5 outstanding)"
        result = send(to=TO_ADDR, subject=subject, body=body, from_address=FROM_ADDR)
        print(f"Overdue alert send result: {result}")
        if result == 'Sent.':
            state['overdue_last_key'] = overdue_key
            json.dump(state, open(STATE_FILE, 'w'))
            print("Overdue alert recorded.")
        return

    # --- Complete cycle (all 5 uploaded) ---
    if state.get('last_sent_cycle') == cycle_key and not force:
        print("Completion report already sent for this cycle. Use --force to resend.")
        return

    body, any_missed = build_body(anchor, deadline, files)
    subj_flag = " - SLA MISSED" if any_missed else " - Within SLA"
    month = anchor.strftime('%B %Y')
    subject = f"SLA Update - EmblemFacets Response Files ({month}){subj_flag}"
    result = send(to=TO_ADDR, subject=subject, body=body, from_address=FROM_ADDR)
    print(f"Send result: {result}")
    if result == 'Sent.':
        state['last_sent_cycle'] = cycle_key
        json.dump(state, open(STATE_FILE, 'w'))
        print("Cycle marked as sent.")


if __name__ == '__main__':
    main()
