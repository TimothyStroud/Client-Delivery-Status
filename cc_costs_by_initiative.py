r"""
cc-costs-by-initiative (Windows/Python port of cc-classify's "initiatives" report)

Scans local Claude Code session transcripts (~/.claude/projects/<initiative>/*.jsonl),
prices every assistant turn with authoritative Anthropic rates (incl. prompt-cache
5m/1h write + read tiers), and prints spend grouped by initiative (project dir).

NOTE: This reproduces cc-classify's SPEND-BY-INITIATIVE view only. It does NOT
reproduce the Dev/COS/Mixed/Strategy capitalization buckets — those come from
cc-classify's own config rules, which aren't published here.

Usage:
  python cc_costs_by_initiative.py --since 1month
  python cc_costs_by_initiative.py --since 2026-06-01 --until 2026-06-30
"""
import argparse
import glob
import json
import os
from datetime import datetime, timedelta, timezone

PROJECTS_DIR = os.path.expanduser(r"~\.claude\projects")

# Base per-MILLION-token rates (USD). Cache multipliers applied below.
#   cache write 5m = 1.25x input | write 1h = 2x input | read = 0.1x input
RATES = {  # match by substring of the model id
    "claude-fable-5":   (10.0, 50.0),
    "claude-opus-4-8":  (5.0, 25.0),
    "claude-opus-4-7":  (5.0, 25.0),
    "claude-opus-4-6":  (5.0, 25.0),
    "claude-opus-4-5":  (5.0, 25.0),
    "claude-opus-4":    (5.0, 25.0),   # 4.0/4.1 fallback
    "claude-sonnet-4-6": (3.0, 15.0),
    "claude-sonnet-4-5": (3.0, 15.0),
    "claude-sonnet-4":  (3.0, 15.0),
    "claude-haiku-4-5": (1.0, 5.0),
    "claude-haiku":     (1.0, 5.0),
}


def rate_for(model):
    if not model:
        return None
    for key, r in RATES.items():
        if model.startswith(key):
            return r
    return None


def parse_since(s):
    s = (s or "").strip().lower()
    now = datetime.now(timezone.utc)
    if s in ("1month", "1m", "month"):
        return now - timedelta(days=30)
    if s.endswith("month"):
        return now - timedelta(days=30 * int(s[:-5]))
    if s.endswith("d") or s.endswith("day"):
        n = int(s.rstrip("day").rstrip("d"))
        return now - timedelta(days=n)
    # explicit date
    return datetime.strptime(s, "%Y-%m-%d").replace(tzinfo=timezone.utc)


def parse_ts(ts):
    if not ts:
        return None
    try:
        return datetime.fromisoformat(ts.replace("Z", "+00:00"))
    except Exception:
        return None


def cost_of(usage, model):
    """Return (usd, total_tokens) for one assistant turn's usage block."""
    r = rate_for(model)
    if r is None:
        return None, 0
    inp_rate, out_rate = r
    write_5m_rate = inp_rate * 1.25
    write_1h_rate = inp_rate * 2.0
    read_rate = inp_rate * 0.10

    inp = usage.get("input_tokens", 0) or 0
    out = usage.get("output_tokens", 0) or 0
    read = usage.get("cache_read_input_tokens", 0) or 0
    # Prefer the exact 5m/1h split when present; else treat all creation as 5m.
    cc = usage.get("cache_creation") or {}
    w5 = cc.get("ephemeral_5m_input_tokens")
    w1 = cc.get("ephemeral_1h_input_tokens")
    if w5 is None and w1 is None:
        w5 = usage.get("cache_creation_input_tokens", 0) or 0
        w1 = 0
    w5 = w5 or 0
    w1 = w1 or 0

    usd = (inp * inp_rate + out * out_rate + w5 * write_5m_rate
           + w1 * write_1h_rate + read * read_rate) / 1_000_000
    return usd, inp + out + w5 + w1 + read


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--since", default="1month")
    ap.add_argument("--until", default=None)
    ap.add_argument("--sessions", metavar="INITIATIVE", default=None,
                    help="break out per-session detail for one initiative")
    args = ap.parse_args()

    since = parse_since(args.since)
    until = (datetime.strptime(args.until, "%Y-%m-%d").replace(tzinfo=timezone.utc)
             + timedelta(days=1)) if args.until else datetime.now(timezone.utc)

    # initiative -> {cost, tokens, sessions:set}
    agg = {}
    unpriced = {}
    # per-session detail (only populated when --sessions is used):
    #   (initiative, sessionId) -> {cost, tokens, first_ts, last_ts, model}
    sess = {}
    for jf in glob.glob(os.path.join(PROJECTS_DIR, "*", "*.jsonl")):
        initiative = os.path.basename(os.path.dirname(jf))
        try:
            fh = open(jf, encoding="utf-8")
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
                if o.get("type") != "assistant":
                    continue
                msg = o.get("message") or {}
                usage = msg.get("usage")
                if not usage:
                    continue
                ts = parse_ts(o.get("timestamp"))
                if ts is None or not (since <= ts < until):
                    continue
                model = msg.get("model") or ""
                usd, toks = cost_of(usage, model)
                if usd is None:
                    unpriced[model] = unpriced.get(model, 0) + 1
                    continue
                e = agg.setdefault(initiative, {"cost": 0.0, "tokens": 0, "sessions": set()})
                e["cost"] += usd
                e["tokens"] += toks
                sid = o.get("sessionId") or os.path.splitext(os.path.basename(jf))[0]
                e["sessions"].add(sid)
                if args.sessions and initiative == args.sessions:
                    s = sess.setdefault((initiative, sid),
                                        {"cost": 0.0, "tokens": 0, "first": ts, "last": ts, "model": model})
                    s["cost"] += usd
                    s["tokens"] += toks
                    if ts < s["first"]:
                        s["first"] = ts
                    if ts > s["last"]:
                        s["last"] = ts

    rows = sorted(agg.items(), key=lambda kv: kv[1]["cost"], reverse=True)
    tot_cost = sum(v["cost"] for _, v in rows)
    tot_tok = sum(v["tokens"] for _, v in rows)
    tot_sess = sum(len(v["sessions"]) for _, v in rows)

    print()
    print("  cc-costs-by-initiative (PowerShell-free Python port)")
    print(f"  Window: {since:%Y-%m-%d} to {until - timedelta(days=1):%Y-%m-%d}")
    print(f"  Projects dir: {PROJECTS_DIR}")
    print()
    print("  SPEND BY INITIATIVE  (sorted by cost)")
    print("  " + "-" * 78)
    print(f"  {'INITIATIVE':<28}{'COST':>13}{'TOKENS':>18}{'SESS':>7}{'$%':>8}")
    print("  " + "-" * 78)
    for name, v in rows:
        pct = (v["cost"] / tot_cost * 100) if tot_cost else 0
        print(f"  {name:<28}{'$'+format(v['cost'], ',.2f'):>13}"
              f"{format(v['tokens'], ','):>18}{len(v['sessions']):>7}{pct:>7.1f}%")
    print("  " + "-" * 78)
    print(f"  {'TOTAL':<28}{'$'+format(tot_cost, ',.2f'):>13}"
          f"{format(tot_tok, ','):>18}{tot_sess:>7}{100.0:>7.1f}%")
    if args.sessions:
        srows = sorted(
            [(sid, v) for (init, sid), v in sess.items() if init == args.sessions],
            key=lambda kv: kv[1]["cost"], reverse=True)
        s_cost = sum(v["cost"] for _, v in srows)
        s_tok = sum(v["tokens"] for _, v in srows)
        print(f"  SESSIONS in '{args.sessions}'  (sorted by cost)")
        print("  " + "-" * 84)
        print(f"  {'SESSION START (local)':<22}{'SESSION ID':<16}{'COST':>12}{'TOKENS':>18}{'$%':>8}")
        print("  " + "-" * 84)
        for sid, v in srows:
            pct = (v["cost"] / s_cost * 100) if s_cost else 0
            start_local = v["first"].astimezone()
            print(f"  {start_local:%Y-%m-%d %H:%M}     {sid[:8]:<16}"
                  f"{'$'+format(v['cost'], ',.2f'):>12}{format(v['tokens'], ','):>18}{pct:>7.1f}%")
        print("  " + "-" * 84)
        print(f"  {'TOTAL ('+str(len(srows))+' sessions)':<38}"
              f"{'$'+format(s_cost, ',.2f'):>12}{format(s_tok, ','):>18}{100.0:>7.1f}%")
        print()

    if unpriced:
        print("  [warning] unpriced models skipped (add to RATES):")
        for m, n in unpriced.items():
            print(f"    {m!r}: {n} turns")
        print()


if __name__ == "__main__":
    main()
