r"""
One-off: resend the June 2026 Claude cost email with the full cc-classify.exe
detail PLUS crow-bot's per-model notes for 2026-06-01..2026-06-30
(from #proj_june_claude_costs, thread ts 1782994090.751139).

Reuses the scheduled emailer's helpers so numbers/formatting match the monthly job.
"""
import sys
BASE = r'C:\Users\tls2\.claude\projects\H--'
sys.path.insert(0, BASE)
from cc_costs_monthly_email import run_exe, parse_total, pre, TO
from send_via_outlook import send

SINCE, UNTIL, LABEL = '2026-06-01', '2026-06-30', 'June 2026'

# crow-bot's own (API-side) report for the same window, captured from Slack.
CROWBOT = [
    ('claude-opus-4-8', '$1,249.70'),
    ('claude-opus-4-7', '$278.40'),
    ('claude-haiku-4-5-20251001', '$0.01'),
]
CROWBOT_TOTAL = '$1,528.11'


def crowbot_block():
    rows = "".join(
        f"<tr><td style='padding:3px 14px 3px 0;'>{m}</td>"
        f"<td style='padding:3px 0;text-align:right;font-family:Consolas,monospace;'>{c}</td></tr>"
        for m, c in CROWBOT)
    return (
        "<h3 style='color:#2f5496;'>crow-bot notes &mdash; per-model (2026-06-01 to 2026-06-30)</h3>"
        "<p style='color:#666;margin-top:0;'>crow-bot's own report from "
        "<code>#proj_june_claude_costs</code> (company API-side tool, by model). "
        "This is a separate estimate from cc-classify and totals higher &mdash; crow-bot bills off "
        "logged API usage across all directories, cc-classify prices the local session transcripts "
        "for the <b>RDP Data Operations</b> initiative only.</p>"
        "<table style='border-collapse:collapse;font-size:13px;'>"
        + rows +
        f"<tr style='border-top:1px solid #bbb;font-weight:bold;'>"
        f"<td style='padding:5px 14px 3px 0;'>Total</td>"
        f"<td style='padding:5px 0 3px;text-align:right;font-family:Consolas,monospace;'>{CROWBOT_TOTAL}</td></tr>"
        "</table>")


def main():
    report_txt = run_exe('report', SINCE, UNTIL)
    init_txt = run_exe('initiatives', SINCE, UNTIL)
    sess_txt = run_exe('sessions', SINCE, UNTIL)
    total = parse_total(report_txt)

    css = "font-family:Segoe UI,Arial,sans-serif;font-size:13px;"
    p = [f"<div style='{css}'>"]
    p.append(f"<h2 style='color:#2f5496;margin-bottom:2px;'>Claude Cost Report &mdash; {LABEL}</h2>")
    p.append(f"<p style='color:#666;margin-top:0;'>cc-classify total spend (RDP Data Operations): <b>{total}</b> "
             f"for {SINCE} to {UNTIL}. Source: official <code>cc-classify.exe</code> v0.7.0 "
             f"(matches #proj_june_claude_costs).</p>")
    p.append(crowbot_block())
    p.append("<h3 style='color:#2f5496;'>Capitalization report</h3>")
    p.append(pre(report_txt))
    p.append("<h3 style='color:#2f5496;'>Initiatives</h3>")
    p.append(pre(init_txt))
    p.append("<h3 style='color:#2f5496;'>Sessions</h3>")
    p.append(pre(sess_txt))
    p.append("<p style='color:#888;font-size:11px;margin-top:14px;'>"
             "Initiative mapping (H:\\ &rarr; \"RDP Data Operations\", capitalizable R&amp;D) is defined in "
             "<code>%USERPROFILE%\\.config\\cc-classify\\config.toml</code>. Buckets and Cap% are computed "
             "by cc-classify itself. crow-bot figures are that tool's own per-model estimate.</p>")
    p.append("</div>")
    html = "".join(p)

    subject = f"Claude Cost Report - {LABEL} - {total} (cc-classify) / {CROWBOT_TOTAL} (crow-bot)"
    result = send(to=TO, subject=subject, body=html)
    print(f"{LABEL}: cc-classify {total} | crow-bot {CROWBOT_TOTAL} | email -> {TO} | {result}")


if __name__ == '__main__':
    main()
