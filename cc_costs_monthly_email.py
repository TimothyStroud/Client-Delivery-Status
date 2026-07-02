r"""
Monthly Claude-cost email — runs on the 1st, reports the PREVIOUS calendar month.

Reuses the pricing/scan logic in cc_costs_by_initiative.py, builds an HTML
spend-by-initiative + per-session report, and emails it to the user via the
established Outlook-COM interactive-task sender (send_via_outlook.send).

Scheduled by Windows Task "CC Monthly Cost Email" (day 1 of each month, 8:00 AM,
interactive logon so Outlook COM has a session).
"""
import os
import sys
import glob
import json
import html as _html
import subprocess
from datetime import datetime, timedelta, timezone

BASE = r'C:\Users\tls2\.claude\projects\H--'
sys.path.insert(0, BASE)
import cc_costs_by_initiative as cc          # cost_of, parse_ts, PROJECTS_DIR
from send_via_outlook import send

TO = 'timothy.stroud@machinify.com'
PS1 = os.path.join(BASE, 'cc-classify-powershell', 'cc-classify.ps1')


def run_ps_summary(since, until_exclusive):
    """Run cc-classify.ps1 `summary` for the given window and return its text
    (the CAPITALIZATION REPORT + INITIATIVES sections, verbatim). until is
    exclusive; the .ps1 -Until is inclusive, so pass the last covered day."""
    last_day = (until_exclusive - timedelta(days=1)).strftime('%Y-%m-%d')
    first_day = since.strftime('%Y-%m-%d')
    r = subprocess.run(
        ['powershell', '-NoProfile', '-NonInteractive', '-File', PS1,
         'summary', '-Since', first_day, '-Until', last_day],
        capture_output=True, text=True, timeout=180)
    out = (r.stdout or '').rstrip('\n')
    if not out:
        out = '(cc-classify.ps1 produced no output)\n' + (r.stderr or '')
    return out


def prev_month_window(today=None):
    """Return (since_utc, until_utc_exclusive, label) for the calendar month
    before `today` (defaults to now). until is exclusive (first of this month)."""
    now = today or datetime.now(timezone.utc)
    first_this = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    last_prev = first_this - timedelta(days=1)          # last day of prev month
    since = last_prev.replace(day=1)                    # first day of prev month
    return since, first_this, since.strftime('%B %Y')


def scan(since, until):
    """Aggregate spend by initiative and by session in [since, until)."""
    init_agg = {}    # name -> {cost, tokens, sessions:set}
    sess_agg = {}    # (init, sid) -> {cost, tokens, first, last}
    unpriced = {}
    for jf in glob.glob(os.path.join(cc.PROJECTS_DIR, '*', '*.jsonl')):
        initiative = os.path.basename(os.path.dirname(jf))
        try:
            fh = open(jf, encoding='utf-8')
        except Exception:
            continue
        with fh:
            for line in fh:
                line = line.strip()
                if not line:
                    continue
                try:
                    o = json.loads(line)
                except Exception:
                    continue
                if o.get('type') != 'assistant':
                    continue
                usage = (o.get('message') or {}).get('usage')
                if not usage:
                    continue
                ts = cc.parse_ts(o.get('timestamp'))
                if ts is None or not (since <= ts < until):
                    continue
                model = (o.get('message') or {}).get('model') or ''
                usd, toks = cc.cost_of(usage, model)
                if usd is None:
                    unpriced[model] = unpriced.get(model, 0) + 1
                    continue
                sid = o.get('sessionId') or os.path.splitext(os.path.basename(jf))[0]
                e = init_agg.setdefault(initiative, {'cost': 0.0, 'tokens': 0, 'sessions': set()})
                e['cost'] += usd
                e['tokens'] += toks
                e['sessions'].add(sid)
                s = sess_agg.setdefault((initiative, sid),
                                        {'cost': 0.0, 'tokens': 0, 'first': ts, 'last': ts})
                s['cost'] += usd
                s['tokens'] += toks
                if ts < s['first']:
                    s['first'] = ts
                if ts > s['last']:
                    s['last'] = ts
    return init_agg, sess_agg, unpriced


def build_html(label, summary_text, init_agg, sess_agg, unpriced):
    tot_cost = sum(v['cost'] for v in init_agg.values())
    tot_tok = sum(v['tokens'] for v in init_agg.values())
    tot_sess = sum(len(v['sessions']) for v in init_agg.values())

    css = ("font-family:Segoe UI,Arial,sans-serif;font-size:13px;")
    th = "background:#2f5496;color:#fff;padding:6px 10px;text-align:left;"
    thr = th.replace('left', 'right')
    td = "padding:5px 10px;border-bottom:1px solid #e0e0e0;"
    tdr = td + "text-align:right;"

    parts = [f"<div style='{css}'>"]
    parts.append(f"<h2 style='color:#2f5496;margin-bottom:2px;'>Claude Cost Report &mdash; {label}</h2>")
    parts.append(f"<p style='color:#666;margin-top:0;'>Total: <b>${tot_cost:,.2f}</b> across "
                 f"{tot_tok:,} tokens in {tot_sess} sessions.</p>")

    # cc-classify.ps1 `summary` output verbatim (CAPITALIZATION REPORT + INITIATIVES)
    parts.append("<h3 style='color:#2f5496;'>cc-classify summary</h3>")
    parts.append("<pre style='font-family:Consolas,Courier New,monospace;font-size:12px;"
                 "background:#f5f5f5;border:1px solid #ddd;padding:12px;white-space:pre;"
                 "overflow-x:auto;'>" + _html.escape(summary_text) + "</pre>")

    # By session
    parts.append("<h3 style='color:#2f5496;'>By session</h3>")
    parts.append("<table style='border-collapse:collapse;'>")
    parts.append(f"<tr><th style='{th}'>Session start (local)</th><th style='{th}'>Initiative</th>"
                 f"<th style='{th}'>ID</th><th style='{thr}'>Cost</th><th style='{thr}'>Tokens</th>"
                 f"<th style='{thr}'>$%</th></tr>")
    srows = sorted(sess_agg.items(), key=lambda kv: kv[1]['cost'], reverse=True)
    for (init, sid), v in srows:
        pct = (v['cost'] / tot_cost * 100) if tot_cost else 0
        start_local = v['first'].astimezone()
        parts.append(f"<tr><td style='{td}'>{start_local:%Y-%m-%d %H:%M}</td><td style='{td}'>{init}</td>"
                     f"<td style='{td}'>{sid[:8]}</td><td style='{tdr}'>${v['cost']:,.2f}</td>"
                     f"<td style='{tdr}'>{v['tokens']:,}</td><td style='{tdr}'>{pct:.1f}%</td></tr>")
    parts.append("</table>")

    note = ("Rates: Opus $5/$25, Sonnet $3/$15, Haiku $1/$5 per MTok; cache write 5m 1.25&times;, "
            "1h 2&times;, read 0.1&times; input. Capitalization buckets above default every initiative to "
            "Dev &mdash; real cc-classify assigns Dev/COS/Mixed/Strategy from config.toml rules not "
            "available here.")
    if unpriced:
        note += " Skipped unpriced turns: " + ", ".join(f"{m} ({n})" for m, n in unpriced.items()) + "."
    parts.append(f"<p style='color:#888;font-size:11px;margin-top:14px;'>{note}</p>")
    parts.append("</div>")
    return "".join(parts)


def main():
    since, until, label = prev_month_window()
    init_agg, sess_agg, unpriced = scan(since, until)
    summary_text = run_ps_summary(since, until)
    html = build_html(label, summary_text, init_agg, sess_agg, unpriced)
    tot_cost = sum(v['cost'] for v in init_agg.values())
    subject = f"Claude Cost Report - {label} - ${tot_cost:,.2f}"
    result = send(to=TO, subject=subject, body=html)
    print(f"{label}: ${tot_cost:,.2f} | email -> {TO} | {result}")


if __name__ == '__main__':
    main()
