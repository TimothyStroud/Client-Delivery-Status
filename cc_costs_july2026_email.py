r"""
One-off Claude cost email for JULY 2026, sent on request (2026-08-10).

Context: David Levinger asked (#proj_july_claude_costs, 2026-08-07 6:42pm) that
everyone move to `cc-classify cap-report` (new in v0.12.1+), which auto-selects
the previous month and uploads a standard-format file to S3.

Two blockers on this box, both reported honestly in the email body:
  1. cap-report needs cc-classify >= 0.12.1. This machine has the v0.7.0 Windows
     build CJ posted to Slack. vlognow/cc-classify is a private repo: the GitHub
     API and the release asset both 404 unauthenticated, there is no `gh` CLI or
     token here, and there is no cargo/crates.io path either. So cap-report
     cannot be run until someone supplies the v0.12.2 Windows zip.
  2. The S3 upload needs the machinify-dev AWS account over SSO. There is no AWS
     CLI and no ~/.aws/config on this box, so even with the new binary this would
     be a `cap-report --local` run (same position Margaret Lane reported).

What this email DOES report: July 2026 from the v0.7.0 binary, with the pricing
table patched in config.toml. That patch matters -- v0.7.0's built-in table only
globs *opus-4*/*sonnet-4*/*haiku-4*, so every claude-opus-5 turn (i.e. nearly all
of this box's usage) was silently priced at $0. Unpatched July read $428.16;
correctly priced it is $461.48.
"""
import os
import sys
import html as _html

BASE = r'C:\Users\tls2\.claude\projects\H--'
sys.path.insert(0, BASE)
from send_via_outlook import send
from cc_costs_monthly_email import run_exe, parse_total, pre

TO = 'timothy.stroud@machinify.com'
SINCE, UNTIL, LABEL = '2026-07-01', '2026-07-31', 'July 2026'


def build_html(report_txt, init_txt, sess_txt, total):
    css = "font-family:Segoe UI,Arial,sans-serif;font-size:13px;"
    warn = ("background:#fff4e5;border-left:4px solid #ff9800;padding:10px 14px;"
            "margin:12px 0;font-size:12.5px;")
    note = ("background:#e8f4fd;border-left:4px solid #2f5496;padding:10px 14px;"
            "margin:12px 0;font-size:12.5px;")
    p = [f"<div style='{css}'>"]
    p.append(f"<h2 style='color:#2f5496;margin-bottom:2px;'>Claude Cost Report &mdash; {LABEL}</h2>")
    p.append(f"<p style='color:#666;margin-top:0;'>Total spend: <b>{total}</b> "
             f"across 107 sessions, ~488.7M tokens. One initiative: "
             f"<b>RDP Data Operations</b> (R&amp;D, capitalizable), ~98% capitalizable.</p>")

    p.append(f"<div style='{warn}'><b>This is not the new <code>cap-report</code> output.</b> "
             "David's 8/7 note asks everyone to upgrade to cc-classify v0.12.1+ and run "
             "<code>cc-classify cap-report</code>, which picks last month automatically and "
             "uploads a standard file to S3. Two things block that on this machine:"
             "<ol style='margin:6px 0 0 18px;padding:0;'>"
             "<li><b>No v0.12.2 binary.</b> <code>vlognow/cc-classify</code> is a private repo &mdash; "
             "the releases API and the <code>...-pc-windows-msvc.zip</code> asset both return 404 "
             "unauthenticated, and there is no <code>gh</code> CLI, GitHub token, or cargo toolchain "
             "on this box. The installed build is the v0.7.0 Windows zip CJ posted to Slack back in June, "
             "which has no <code>cap-report</code> command. "
             "<i>Fix: someone with repo access posts the v0.12.2 Windows zip to Slack, same as CJ did for 0.7.0.</i></li>"
             "<li><b>No AWS access.</b> No AWS CLI and no <code>~/.aws/config</code> here, so the S3 upload "
             "can't happen either &mdash; this would be a <code>cap-report --local</code> run and a CSV posted "
             "to the channel, the same workaround Margaret Lane used.</i></li>"
             "</ol></div>")

    p.append(f"<div style='{note}'><b>Pricing correction &mdash; the number moved.</b> "
             "v0.7.0's built-in pricing table only matches <code>*opus-4*</code>, <code>*sonnet-4*</code> "
             "and <code>*haiku-4*</code>. Every <code>claude-opus-5</code> turn &mdash; effectively all of this "
             "box's usage &mdash; was unpriced and counted as <b>$0</b>. I added the Claude 5 rates "
             "(Opus 5 $5/$25 per MTok, cache-write 5m $6.25, cache-read $0.50) to "
             "<code>config.toml</code>. July reads <b>$461.48</b> corrected, vs <b>$428.16</b> before. "
             "The automated 8/1 email you already received used the uncorrected figure.<br><br>"
             "Note the <code>[[pricing]]</code> block <i>replaces</i> the built-in table rather than merging, "
             "so the stock opus-4/sonnet-4/haiku-4 rows are restated in the config too. All of this becomes "
             "unnecessary once the real v0.12.2 binary is installed &mdash; it prices the 5 family natively.</div>")

    p.append("<h3 style='color:#2f5496;'>Capitalization report</h3>")
    p.append(pre(report_txt))
    p.append("<h3 style='color:#2f5496;'>Initiatives</h3>")
    p.append(pre(init_txt))
    p.append("<h3 style='color:#2f5496;'>Sessions</h3>")
    p.append(pre(sess_txt))
    p.append("<p style='color:#888;font-size:11px;margin-top:14px;'>"
             "Initiative mapping (H:\\ &rarr; \"RDP Data Operations\", capitalizable R&amp;D) and the pricing "
             "patch live in <code>%USERPROFILE%\\.config\\cc-classify\\config.toml</code>. Buckets and Cap% are "
             "computed by cc-classify itself. The monthly Windows task \"CC Monthly Cost Email\" has been "
             "disabled pending the cap-report cutover.</p>")
    p.append("</div>")
    return "".join(p)


def main():
    report_txt = run_exe('report', SINCE, UNTIL)
    init_txt = run_exe('initiatives', SINCE, UNTIL)
    sess_txt = run_exe('sessions', SINCE, UNTIL)
    total = parse_total(report_txt)
    html = build_html(report_txt, init_txt, sess_txt, total)
    subject = f"Claude Cost Report - {LABEL} - {total} (cap-report blocked - see note)"
    result = send(to=TO, subject=subject, body=html)
    print(f"{LABEL}: {total} | email -> {TO} | {result}")


if __name__ == '__main__':
    main()
