r"""
Monthly Claude-cost email — runs on the 1st, reports the PREVIOUS calendar month.

Drives the OFFICIAL cc-classify.exe (vlognow, v0.12.2) with the local config.toml
initiative mapping, and emails its `report` (TOKEN CAPITALIZATION REPORT) +
`initiatives` + `sessions` output to the user. The .exe numbers are authoritative
and match #proj_july_claude_costs. Sent via the Outlook-COM interactive-task sender.

Also runs `cap-report` (v0.12.1+) to produce the standardised rollup CSV that
finance collects. S3 upload needs an authenticated machinify-dev AWS profile,
which this box does not have, so it falls back to `--local` and the CSV path is
included in the email for manual posting to #proj_july_claude_costs.

Scheduled by Windows Task "CC Monthly Cost Email" (day 1, 8:00 AM, interactive
logon so Outlook COM has a session).

RETENTION: cc-classify reads ~/.claude/projects/*/*.jsonl transcripts, which
Claude Code deletes at startup once they are older than `cleanupPeriodDays`
(default 30). That silently shrinks a month's spend as the month ages -- on
2026-08-10 it had already erased 1-9 July and $86.51. `cleanupPeriodDays` is
pinned to 3650 in %USERPROFILE%\.claude\settings.json; do not lower it.
"""
import os
import sys
import html as _html
import subprocess
import re
from datetime import datetime, timedelta

BASE = r'C:\Users\tls2\.claude\projects\H--'
EXE = os.path.join(BASE, 'cc-classify-bin', 'v0122', 'cc-classify.exe')
CFG = r'C:\Users\tls2\.config\cc-classify\config.toml'
CAPDIR = os.path.join(BASE, 'capreport')
sys.path.insert(0, BASE)
from send_via_outlook import send

TO = 'timothy.stroud@machinify.com'


def prev_month_window(today=None):
    """(since 'YYYY-MM-DD', until 'YYYY-MM-DD' inclusive, 'Month YYYY') for the
    calendar month before `today`."""
    now = today or datetime.now()
    first_this = now.replace(day=1)
    last_prev = first_this - timedelta(days=1)
    since = last_prev.replace(day=1)
    return since.strftime('%Y-%m-%d'), last_prev.strftime('%Y-%m-%d'), since.strftime('%B %Y')


def run_exe(subcommand, since, until):
    """Run `cc-classify.exe --config CFG <subcommand> --since .. --until ..` and
    return stdout (the report text). Returns an error marker on failure."""
    try:
        r = subprocess.run([EXE, '--config', CFG, subcommand,
                            '--since', since, '--until', until],
                           capture_output=True, text=True,
                           encoding='utf-8', errors='replace', timeout=180)
        out = (r.stdout or '').rstrip('\n')
        if not out:
            out = f'({subcommand}: no output)\n' + (r.stderr or '')
        return out
    except Exception as e:
        return f'({subcommand}: failed to run cc-classify.exe: {e})'


def run_cap_report(month):
    """Run `cap-report --month YYYY-MM`, trying the S3 upload first and falling
    back to `--local`. Returns (status_text, csv_path_or_None)."""
    try:
        os.makedirs(CAPDIR, exist_ok=True)
    except Exception as e:
        return f'(cap-report: cannot create {CAPDIR}: {e})', None

    def _run(extra):
        return subprocess.run([EXE, '--config', CFG, 'cap-report',
                               '--month', month] + extra,
                              capture_output=True, text=True, encoding='utf-8',
                              errors='replace', timeout=300, cwd=CAPDIR)
    try:
        r = _run([])
        if r.returncode == 0:
            return 'Uploaded to S3.\n' + (r.stdout or '').rstrip('\n'), None
        s3_err = ((r.stdout or '') + (r.stderr or '')).rstrip('\n')
        r = _run(['--local'])
        csv_path = os.path.join(CAPDIR, f'{month}_{TO}_rollup.csv')
        if r.returncode != 0:
            return ('S3 upload failed AND --local failed.\n' + s3_err + '\n' +
                    ((r.stdout or '') + (r.stderr or '')).rstrip('\n')), None
        return ('S3 upload unavailable (no machinify-dev AWS profile on this box) '
                '- wrote the local rollup CSV instead. Post it to '
                '#proj_july_claude_costs.\n' + (r.stdout or '').rstrip('\n') +
                f'\nCSV: {csv_path}'), (csv_path if os.path.exists(csv_path) else None)
    except Exception as e:
        return f'(cap-report: failed to run: {e})', None


def parse_total(report_text):
    """Pull the total spend string from `report` output for the subject line."""
    m = re.search(r'Total spend:\s*(\$[\d,]+\.\d{2})', report_text)
    if m:
        return m.group(1)
    m = re.search(r'^\s*TOTAL\s+(\$[\s\d,]+\.\d{2})', report_text, re.M)
    return m.group(1).replace(' ', '') if m else '$?'


def asciify(text):
    """Map cc-classify's Unicode box-drawing / marks to ASCII so the report
    survives the Outlook JSON round-trip (PS 5.1 reads the params file as ANSI)."""
    repl = {'─': '-', '═': '=', '│': '|', '║': '|',
            '✓': 'Y', '✗': 'N', '→': '->', '…': '...',
            '‘': "'", '’': "'", '“': '"', '”': '"'}
    for k, v in repl.items():
        text = text.replace(k, v)
    return text.encode('ascii', 'replace').decode('ascii')


def pre(text):
    text = asciify(text)
    return ("<pre style='font-family:Consolas,Courier New,monospace;font-size:12px;"
            "background:#f5f5f5;border:1px solid #ddd;padding:12px;white-space:pre;"
            "overflow-x:auto;'>" + _html.escape(text) + "</pre>")


def build_html(label, report_txt, init_txt, sess_txt, total, cap_txt='', cap_csv=''):
    css = "font-family:Segoe UI,Arial,sans-serif;font-size:13px;"
    p = [f"<div style='{css}'>"]
    p.append(f"<h2 style='color:#2f5496;margin-bottom:2px;'>Claude Cost Report &mdash; {label}</h2>")
    p.append(f"<p style='color:#666;margin-top:0;'>Total spend: <b>{total}</b>. "
             f"Source: official <code>cc-classify.exe</code> v0.12.2 (matches #proj_july_claude_costs).</p>")
    if cap_txt:
        p.append("<h3 style='color:#2f5496;'>cap-report (finance submission)</h3>")
        p.append(pre(cap_txt))
        if cap_csv:
            p.append(f"<p style='color:#a33;'><b>Action required:</b> paste the contents of "
                     f"<code>{_html.escape(cap_csv)}</code> into #proj_july_claude_costs "
                     f"(S3 upload is unavailable on this box).</p>")
    p.append("<h3 style='color:#2f5496;'>Capitalization report</h3>")
    p.append(pre(report_txt))
    p.append("<h3 style='color:#2f5496;'>Initiatives</h3>")
    p.append(pre(init_txt))
    p.append("<h3 style='color:#2f5496;'>Sessions</h3>")
    p.append(pre(sess_txt))
    p.append("<p style='color:#888;font-size:11px;margin-top:14px;'>"
             "Initiative mapping (H:\\ &rarr; \"RDP Data Operations\", capitalizable R&amp;D) is defined in "
             "<code>%USERPROFILE%\\.config\\cc-classify\\config.toml</code>. Buckets and Cap% are computed "
             "by cc-classify itself.</p>")
    p.append("</div>")
    return "".join(p)


def main():
    since, until, label = prev_month_window()
    report_txt = run_exe('report', since, until)
    init_txt = run_exe('initiatives', since, until)
    sess_txt = run_exe('sessions', since, until)
    cap_txt, cap_csv = run_cap_report(since[:7])
    total = parse_total(report_txt)
    html = build_html(label, report_txt, init_txt, sess_txt, total, cap_txt, cap_csv or '')
    subject = f"Claude Cost Report - {label} - {total}"
    result = send(to=TO, subject=subject, body=html)
    print(f"{label}: {total} | cap-csv: {cap_csv or 'n/a'} | email -> {TO} | {result}")


if __name__ == '__main__':
    main()
