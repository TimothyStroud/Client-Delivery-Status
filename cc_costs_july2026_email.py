r"""
One-off Claude cost email for JULY 2026.

First sent 2026-08-10 07:39 reporting $461.48 with two blockers (no v0.12.2
binary, no AWS). RESENT the same day after both were resolved and a third,
worse problem surfaced.

What changed for the resend:
  1. cap-report UNBLOCKED. Derek Leverenz posted the v0.12.2 Windows build to
     #proj_july_claude_costs at 11:07am. Installed to cc-classify-bin\v0122\.
  2. Pricing patch REMOVED. v0.12.2 prices the Claude 5 family natively at
     exactly the rates that were patched into config.toml, and adds gpt-5.6-*
     rows that the patch would have silently dropped (a [[pricing]] block
     replaces the built-in table rather than merging).
  3. S3 still unavailable -- no machinify-dev AWS profile -- so cap-report was
     run with --local and the rollup CSV is embedded below and posted to Slack.
  4. TRANSCRIPT RETENTION LOSS. cc-classify reads ~/.claude/projects/**/*.jsonl,
     and Claude Code deletes transcripts older than cleanupPeriodDays (default
     30) at startup. Between the 7:39am run and 11:20am the same day a /clear
     triggered that cleanup and erased 1-9 July: $461.48/107 sessions became
     $374.97/79, on BOTH binaries. $86.51 (19%) unrecoverable -- transcripts are
     gitignored and were not in the recycle bin, a shadow copy, or OneDrive.
     cleanupPeriodDays is now pinned to 3650 in ~/.claude/settings.json.

So $461.48 is the truer July figure and $374.97 is all the tool can now
reproduce. Both are stated in the email rather than picking one silently.
"""
import os
import sys
import io
import html as _html

BASE = r'C:\Users\tls2\.claude\projects\H--'
sys.path.insert(0, BASE)
from send_via_outlook import send
from cc_costs_monthly_email import run_exe, parse_total, pre, run_cap_report

TO = 'timothy.stroud@machinify.com'
SINCE, UNTIL, LABEL = '2026-07-01', '2026-07-31', 'July 2026'

PRE_PRUNE_TOTAL = '$461.48'
PRE_PRUNE_SESSIONS = 107
PRE_PRUNE_TOKENS = '488,654,092'
SLACK_LINK = 'https://machinify.slack.com/archives/C0B8K5U4U7P/p1786376708834119'


def build_html(report_txt, init_txt, sess_txt, total, cap_txt, cap_csv_text):
    css = "font-family:Segoe UI,Arial,sans-serif;font-size:13px;"
    alert = ("background:#fdecea;border-left:4px solid #d93025;padding:10px 14px;"
             "margin:12px 0;font-size:12.5px;")
    ok = ("background:#e6f4ea;border-left:4px solid #188038;padding:10px 14px;"
          "margin:12px 0;font-size:12.5px;")
    note = ("background:#e8f4fd;border-left:4px solid #2f5496;padding:10px 14px;"
            "margin:12px 0;font-size:12.5px;")
    p = [f"<div style='{css}'>"]
    p.append(f"<h2 style='color:#2f5496;margin-bottom:2px;'>Claude Cost Report &mdash; {LABEL} "
             "<span style='font-size:13px;color:#d93025;'>(corrected resend)</span></h2>")
    p.append(f"<p style='color:#666;margin-top:0;'>Supersedes the 7:39am email. "
             f"Two figures, both real:<br>"
             f"&bull; <b>{PRE_PRUNE_TOTAL}</b> &mdash; {PRE_PRUNE_SESSIONS} sessions, "
             f"~{PRE_PRUNE_TOKENS} tokens. Captured 7:39am, before transcripts were pruned. "
             f"<b>The truer number.</b><br>"
             f"&bull; <b>{total}</b> &mdash; what <code>cap-report</code> can reproduce now, "
             f"and therefore what the standard CSV below says.<br>"
             f"One initiative: <b>RDP Data Operations</b> (R&amp;D, capitalizable), ~98% capitalizable.</p>")

    p.append(f"<div style='{alert}'><b>Why the number dropped $86.51 in four hours.</b> "
             "cc-classify prices the local Claude Code transcripts in "
             "<code>~\\.claude\\projects\\**\\*.jsonl</code>. Claude Code <b>deletes transcripts older "
             "than <code>cleanupPeriodDays</code> (default 30) at startup</b>. A <code>/clear</code> "
             "mid-morning triggered that cleanup and erased <b>1&ndash;9 July</b>. The same July window "
             f"returned {PRE_PRUNE_TOTAL} / {PRE_PRUNE_SESSIONS} sessions at 7:39am and {total} / 79 "
             "sessions at 11:20am &mdash; on <i>both</i> the 0.7.0 and 0.12.2 binaries, so this is data "
             "loss, not a version difference.<br><br>"
             "<b>Unrecoverable:</b> the transcripts are gitignored in the auto-synced repo, and were not "
             "in the Recycle Bin, a shadow copy, or OneDrive.<br><br>"
             "<b>Fixed going forward:</b> <code>\"cleanupPeriodDays\": 3650</code> is now pinned in "
             "<code>%USERPROFILE%\\.claude\\settings.json</code>. Do not lower it &mdash; every month "
             "reported more than ~30 days late is otherwise an undercount, and that likely applies to "
             "other people's channel numbers too. Flagged in Slack: "
             f"<a href='{SLACK_LINK}'>#proj_july_claude_costs</a>.</div>")

    p.append(f"<div style='{ok}'><b>Both previous blockers are cleared.</b>"
             "<ol style='margin:6px 0 0 18px;padding:0;'>"
             "<li><b>v0.12.2 installed.</b> Derek Leverenz posted the Windows build to "
             "#proj_july_claude_costs at 11:07am today; it is now at "
             "<code>cc-classify-bin\\v0122\\cc-classify.exe</code> and <code>cap-report</code> runs.</li>"
             "<li><b>Pricing patch removed.</b> v0.12.2 prices <code>*opus-5*</code> at $5/$25 per MTok "
             "(cache-write $6.25, cache-read $0.50) natively &mdash; identical to the hand-patched rates &mdash; "
             "and adds <code>gpt-5.6-*</code> rows the patch would have silently dropped, since a "
             "<code>[[pricing]]</code> block replaces the built-in table rather than merging. "
             "Old config kept at <code>config.toml.v0.7.0-pricing-patch.bak</code>.</li>"
             "<li><b>S3 still unavailable</b> &mdash; no <code>machinify-dev</code> AWS profile on this box "
             "(<code>cap-report</code> fails with \"A region must be set\"), so this is a "
             "<code>--local</code> run, same as Margaret Lane and Joshua Hart. The CSV is below and has "
             "been posted to the channel.</li>"
             "</ol></div>")

    p.append("<h3 style='color:#2f5496;'>cap-report rollup CSV (the finance artifact)</h3>")
    p.append(pre(cap_csv_text))
    p.append(pre(cap_txt))

    p.append(f"<div style='{note}'><b>Still open:</b> which July figure gets filed. "
             f"{total} is the clean tool-generated artifact; {PRE_PRUNE_TOTAL} is the truer number from "
             "the same tool and config a few hours earlier. Asked David Levinger to choose in the channel. "
             "A corrected CSV was deliberately <i>not</i> hand-written, because the token and engaged-minute "
             "detail for the deleted sessions is gone and inventing it would corrupt a machine-read file.</div>")

    p.append("<h3 style='color:#2f5496;'>Capitalization report</h3>")
    p.append(pre(report_txt))
    p.append("<h3 style='color:#2f5496;'>Initiatives</h3>")
    p.append(pre(init_txt))
    p.append("<h3 style='color:#2f5496;'>Sessions</h3>")
    p.append(pre(sess_txt))
    p.append("<p style='color:#888;font-size:11px;margin-top:14px;'>"
             "Initiative mapping (H:\\ &rarr; \"RDP Data Operations\", capitalizable R&amp;D) lives in "
             "<code>%USERPROFILE%\\.config\\cc-classify\\config.toml</code>. Buckets and Cap% are computed "
             "by cc-classify itself. The monthly Windows task \"CC Monthly Cost Email\" is re-enabled "
             "(day 1, 8:00 AM) and now runs <code>cap-report</code> automatically, falling back to "
             "<code>--local</code> when S3 is unavailable.</p>")
    p.append("</div>")
    return "".join(p)


def main():
    report_txt = run_exe('report', SINCE, UNTIL)
    init_txt = run_exe('initiatives', SINCE, UNTIL)
    sess_txt = run_exe('sessions', SINCE, UNTIL)
    cap_txt, cap_csv = run_cap_report('2026-07')
    cap_csv_text = '(rollup CSV not produced)'
    if cap_csv and os.path.exists(cap_csv):
        cap_csv_text = io.open(cap_csv, encoding='utf-8').read().rstrip('\n')
    total = parse_total(report_txt)
    html = build_html(report_txt, init_txt, sess_txt, total, cap_txt, cap_csv_text)
    subject = (f"Claude Cost Report - {LABEL} - {PRE_PRUNE_TOTAL} pre-prune / {total} reproducible "
               f"(CORRECTED RESEND)")
    result = send(to=TO, subject=subject, body=html)
    print(f"{LABEL}: {total} (pre-prune {PRE_PRUNE_TOTAL}) | email -> {TO} | {result}")


if __name__ == '__main__':
    main()
