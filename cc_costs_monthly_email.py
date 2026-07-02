r"""
Monthly Claude-cost email — runs on the 1st, reports the PREVIOUS calendar month.

Drives the OFFICIAL cc-classify.exe (vlognow, v0.7.0) with the local config.toml
initiative mapping, and emails its `report` (TOKEN CAPITALIZATION REPORT) +
`initiatives` + `sessions` output to the user. The .exe numbers are authoritative
and match #proj_may_claude_costs. Sent via the Outlook-COM interactive-task sender.

Scheduled by Windows Task "CC Monthly Cost Email" (day 1, 8:00 AM, interactive
logon so Outlook COM has a session).
"""
import os
import sys
import html as _html
import subprocess
import re
from datetime import datetime, timedelta

BASE = r'C:\Users\tls2\.claude\projects\H--'
EXE = os.path.join(BASE, 'cc-classify-bin', 'cc-classify.exe')
CFG = r'C:\Users\tls2\.config\cc-classify\config.toml'
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


def build_html(label, report_txt, init_txt, sess_txt, total):
    css = "font-family:Segoe UI,Arial,sans-serif;font-size:13px;"
    p = [f"<div style='{css}'>"]
    p.append(f"<h2 style='color:#2f5496;margin-bottom:2px;'>Claude Cost Report &mdash; {label}</h2>")
    p.append(f"<p style='color:#666;margin-top:0;'>Total spend: <b>{total}</b>. "
             f"Source: official <code>cc-classify.exe</code> v0.7.0 (matches #proj_may_claude_costs).</p>")
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
    total = parse_total(report_txt)
    html = build_html(label, report_txt, init_txt, sess_txt, total)
    subject = f"Claude Cost Report - {label} - {total}"
    result = send(to=TO, subject=subject, body=html)
    print(f"{label}: {total} | email -> {TO} | {result}")


if __name__ == '__main__':
    main()
