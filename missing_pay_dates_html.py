"""Generate a static HTML version of the MissingPayDates report.

Self-contained single .html file: embedded JSON data + vanilla JS filtering.
No external dependencies, no server, no auth. Drop it on a shared drive and
users open it in any modern browser.

Output paths (overwritten each run):
- \\\\trgfile1\\Shared\\DIG\\Data Business Delivery Team\\Delivery Schedule\\Daily Status Reports\\MissingPayDates.html
- C:\\Users\\tls2\\.claude\\projects\\H--\\MissingPayDates.html
- C:\\Users\\tls2\\OneDrive - Machinify\\Documents\\Reports\\MissingPayDates.html

Window: 2025-10-01 forward (configurable via WINDOW_START).

Run:
    python C:\\Users\\tls2\\.claude\\projects\\H--\\missing_pay_dates_html.py
"""

import json
import os
import sys
import shutil
from datetime import date, datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from missing_pay_dates_sync import build_rows, WINDOW_START

OUTPUT_PATHS = [
    r"\\trgfile1\Shared\DIG\Data Business Delivery Team\Delivery Schedule\Daily Status Reports\MissingPayDates.html",
    r"C:\Users\tls2\.claude\projects\H--\MissingPayDates.html",
    r"C:\Users\tls2\OneDrive - Machinify\Documents\Reports\MissingPayDates.html",
]


HTML_TEMPLATE = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Missing Pay Dates</title>
<style>
  :root {
    --bg: #f4f6f9;
    --card: #ffffff;
    --border: #d8dee6;
    --text: #1f2a37;
    --muted: #5b6776;
    --accent: #2c5f8a;
    --accent-dark: #1f3d5c;
    --pink: #ffc7ce;
    --pink-dark: #9c0006;
    --gray: #e6e6e6;
    --gray-dark: #555;
    --green: #c6efce;
    --green-dark: #006100;
  }
  * { box-sizing: border-box; }
  body {
    margin: 0;
    font-family: "Segoe UI", -apple-system, BlinkMacSystemFont, sans-serif;
    background: var(--bg);
    color: var(--text);
    font-size: 14px;
  }
  header {
    background: var(--accent-dark);
    color: #39ff14;
    padding: 16px 24px;
  }
  header h1 { margin: 0; font-size: 18px; font-weight: 600; }
  header .meta { font-size: 12px; opacity: 0.85; margin-top: 4px; }
  main { padding: 16px 24px 32px; }
  .kpis {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
    gap: 12px;
    margin-bottom: 16px;
  }
  .kpi {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 6px;
    padding: 12px 16px;
  }
  .kpi .label { font-size: 11px; color: var(--muted); text-transform: uppercase; letter-spacing: 0.5px; }
  .kpi .value { font-size: 22px; font-weight: 600; margin-top: 2px; color: var(--accent-dark); }
  .filters {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 6px;
    padding: 12px 16px;
    margin-bottom: 16px;
    display: flex;
    flex-wrap: wrap;
    gap: 12px;
    align-items: end;
  }
  .field { display: flex; flex-direction: column; gap: 4px; }
  .field label { font-size: 11px; color: var(--muted); text-transform: uppercase; letter-spacing: 0.5px; }
  .field select, .field input {
    border: 1px solid var(--border);
    border-radius: 4px;
    padding: 6px 8px;
    font-size: 13px;
    background: #fff;
    min-width: 120px;
  }
  .field input[type=date] { min-width: 140px; }
  button {
    background: var(--accent);
    color: #fff;
    border: 0;
    border-radius: 4px;
    padding: 6px 14px;
    cursor: pointer;
    font-size: 13px;
  }
  button.secondary { background: #fff; color: var(--accent); border: 1px solid var(--accent); }
  table {
    width: 100%;
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 6px;
    border-collapse: separate;
    border-spacing: 0;
    overflow: hidden;
  }
  thead { background: var(--accent); color: #39ff14; position: sticky; top: 0; z-index: 1; }
  th {
    text-align: left;
    padding: 6px 8px;
    font-weight: 600;
    font-size: 12px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    cursor: pointer;
    user-select: none;
    white-space: nowrap;
  }
  th .arrow { opacity: 0.5; font-size: 10px; margin-left: 3px; }
  th.sort-asc .arrow, th.sort-desc .arrow { opacity: 1; }
  td {
    padding: 4px 8px;
    border-top: 1px solid var(--border);
    font-size: 13px;
    white-space: nowrap;
  }
  table { width: auto; min-width: 0; }
  td.amount { text-align: right; font-variant-numeric: tabular-nums; }
  td.amount.negative { color: #c00; font-weight: 600; }
  tr:hover td { background: #f8fafc; }
  .badge {
    display: inline-block;
    padding: 2px 8px;
    border-radius: 10px;
    font-size: 11px;
    font-weight: 600;
  }
  .badge.notrec { background: var(--pink); color: var(--pink-dark); }
  .badge.weekend { background: var(--gray); color: var(--gray-dark); }
  .badge.holiday { background: #fff2cc; color: #806000; }
  .badge.paid { background: var(--green); color: var(--green-dark); }
  .badge.negpaid { background: #f8d7da; color: #c00; }
  .badge + .badge { margin-left: 4px; }
  .pager { display: flex; align-items: center; gap: 12px; margin-top: 12px; justify-content: flex-end; font-size: 13px; color: var(--muted); }
  .empty { padding: 24px; text-align: center; color: var(--muted); }
</style>
</head>
<body>
<header>
  <h1>Missing Pay Dates</h1>
  <div class="meta">Generated __GENERATED__ &middot; window __WINDOW_START__ &rarr; __WINDOW_END__ &middot; __ROW_COUNT__ rows total</div>
</header>
<main>
  <section class="kpis">
    <div class="kpi"><div class="label">Total Rows</div><div class="value" id="kpi-rows">&mdash;</div></div>
    <div class="kpi"><div class="label">Not Received</div><div class="value" id="kpi-notrec">&mdash;</div></div>
    <div class="kpi"><div class="label">Weekend</div><div class="value" id="kpi-weekend">&mdash;</div></div>
    <div class="kpi"><div class="label">Holiday</div><div class="value" id="kpi-holiday">&mdash;</div></div>
    <div class="kpi"><div class="label">Negative Paid</div><div class="value" id="kpi-negpaid">&mdash;</div></div>
    <div class="kpi"><div class="label">Paid</div><div class="value" id="kpi-paid">&mdash;</div></div>
    <div class="kpi"><div class="label">Sum Paid</div><div class="value" id="kpi-sum">&mdash;</div></div>
  </section>

  <section class="filters">
    <div class="field">
      <label for="f-client">Database Name</label>
      <select id="f-client"><option value="">All clients</option></select>
    </div>
    <div class="field">
      <label for="f-year">Year</label>
      <select id="f-year"><option value="">All</option></select>
    </div>
    <div class="field">
      <label for="f-month">Month</label>
      <select id="f-month">
        <option value="">All</option>
        <option value="1">January</option><option value="2">February</option>
        <option value="3">March</option><option value="4">April</option>
        <option value="5">May</option><option value="6">June</option>
        <option value="7">July</option><option value="8">August</option>
        <option value="9">September</option><option value="10">October</option>
        <option value="11">November</option><option value="12">December</option>
      </select>
    </div>
    <div class="field">
      <label for="f-comment">Status</label>
      <select id="f-comment">
        <option value="">All</option>
        <option value="Not Received">Not Received only</option>
        <option value="Weekend">Weekend only</option>
        <option value="Holiday">Holiday only</option>
        <option value="Negative Paid">Negative Paid only</option>
        <option value="Paid">Paid only</option>
      </select>
    </div>
    <div class="field">
      <label for="f-from">From</label>
      <input type="date" id="f-from">
    </div>
    <div class="field">
      <label for="f-to">To</label>
      <input type="date" id="f-to">
    </div>
    <div class="field">
      <label for="f-low">Low &lt; $</label>
      <input type="number" id="f-low" placeholder="e.g. 1000" step="100" min="0">
    </div>
    <button class="secondary" id="btn-reset">Reset</button>
  </section>

  <table id="grid">
    <thead>
      <tr>
        <th data-key="date">Date<span class="arrow">&#8597;</span></th>
        <th data-key="weekday">Day of Week<span class="arrow">&#8597;</span></th>
        <th data-key="client">Database Name<span class="arrow">&#8597;</span></th>
        <th data-key="amount" class="amount">Sum Paid<span class="arrow">&#8597;</span></th>
        <th data-key="comment">Comment<span class="arrow">&#8597;</span></th>
      </tr>
    </thead>
    <tbody id="grid-body"></tbody>
  </table>
  <div class="pager">
    <span id="pager-info"></span>
    <button class="secondary" id="pg-prev">&laquo; Prev</button>
    <button class="secondary" id="pg-next">Next &raquo;</button>
  </div>
</main>

<script type="application/json" id="data">__DATA_JSON__</script>
<script>
(function() {
  const ROWS = JSON.parse(document.getElementById('data').textContent);
  const PAGE_SIZE = 100;
  let state = {
    client: '', year: '', month: '', comment: '', from: '', to: '', low: '',
    sortKey: 'date', sortDir: 'desc', page: 0,
  };

  const fmtMoney = (v) => v == null ? '' : v.toLocaleString('en-US', { style: 'currency', currency: 'USD' });
  const fmtMoneyWhole = (v) => v == null ? '' : Math.round(v).toLocaleString('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 0 });
  const $ = (id) => document.getElementById(id);

  // Populate client + year dropdowns from data
  const clients = [...new Set(ROWS.map(r => r.client))].sort();
  const years = [...new Set(ROWS.map(r => r.year))].sort();
  for (const c of clients) {
    const opt = document.createElement('option');
    opt.value = c; opt.textContent = c;
    $('f-client').appendChild(opt);
  }
  for (const y of years) {
    const opt = document.createElement('option');
    opt.value = y; opt.textContent = y;
    $('f-year').appendChild(opt);
  }

  // A row can carry more than one status (e.g. "Not Received; Holiday: MLK Day").
  // statusesOf returns the set of badges to render; primaryStatus returns the
  // dominant one for filter matching.
  function statusesOf(row) {
    const c = row.comment || '';
    const out = [];
    if (c.indexOf('Not Received') !== -1) out.push('Not Received');
    if (c.indexOf('Weekend') !== -1) out.push('Weekend');
    if (c.indexOf('Holiday:') !== -1) out.push('Holiday');
    if (c.indexOf('Negative Paid') !== -1) out.push('Negative Paid');
    // Paid = received any amount and not flagged Not Received/Weekend.
    // Negative Paid implies paid (just a refund/adjustment) so we still
    // add Paid alongside it.
    if (row.amount != null && out.indexOf('Not Received') === -1 && out.indexOf('Weekend') === -1) {
      out.push('Paid');
    }
    if (out.length === 0) out.push('Paid');
    return out;
  }
  function holidayNameOf(row) {
    const c = row.comment || '';
    const i = c.indexOf('Holiday:');
    if (i === -1) return '';
    return c.slice(i + 'Holiday:'.length).split(';')[0].trim();
  }

  function applyFilters() {
    return ROWS.filter(r => {
      if (state.client && r.client !== state.client) return false;
      if (state.year && r.year != state.year) return false;
      if (state.month && r.month != state.month) return false;
      if (state.from && r.date < state.from) return false;
      if (state.to && r.date > state.to) return false;
      if (state.comment) {
        if (statusesOf(r).indexOf(state.comment) === -1) return false;
      }
      if (state.low !== '' && state.low !== null) {
        const lo = Number(state.low);
        if (r.amount == null || r.amount >= lo) return false;
      }
      return true;
    });
  }

  function sortRows(rows) {
    const k = state.sortKey;
    const dir = state.sortDir === 'asc' ? 1 : -1;
    return rows.slice().sort((a, b) => {
      let av = a[k], bv = b[k];
      if (k === 'amount') {
        av = av == null ? -Infinity : av;
        bv = bv == null ? -Infinity : bv;
      }
      if (av == null) av = '';
      if (bv == null) bv = '';
      if (av < bv) return -1 * dir;
      if (av > bv) return  1 * dir;
      return 0;
    });
  }

  function render() {
    const filtered = applyFilters();
    const sorted = sortRows(filtered);

    // KPIs (a row can count toward multiple status buckets)
    let notrec = 0, weekend = 0, holiday = 0, negpaid = 0, paid = 0, sum = 0;
    for (const r of filtered) {
      const ss = statusesOf(r);
      if (ss.indexOf('Not Received') !== -1) notrec++;
      if (ss.indexOf('Weekend') !== -1) weekend++;
      if (ss.indexOf('Holiday') !== -1) holiday++;
      if (ss.indexOf('Negative Paid') !== -1) negpaid++;
      if (ss.indexOf('Paid') !== -1) { paid++; sum += r.amount || 0; }
    }
    $('kpi-rows').textContent = filtered.length.toLocaleString();
    $('kpi-notrec').textContent = notrec.toLocaleString();
    $('kpi-weekend').textContent = weekend.toLocaleString();
    $('kpi-holiday').textContent = holiday.toLocaleString();
    $('kpi-negpaid').textContent = negpaid.toLocaleString();
    $('kpi-paid').textContent = paid.toLocaleString();
    $('kpi-sum').textContent = fmtMoneyWhole(sum);

    // Pagination
    const pageCount = Math.max(1, Math.ceil(sorted.length / PAGE_SIZE));
    if (state.page >= pageCount) state.page = pageCount - 1;
    if (state.page < 0) state.page = 0;
    const start = state.page * PAGE_SIZE;
    const slice = sorted.slice(start, start + PAGE_SIZE);

    // Body
    const body = $('grid-body');
    body.innerHTML = '';
    if (slice.length === 0) {
      const tr = document.createElement('tr');
      tr.innerHTML = '<td colspan="5" class="empty">No rows match the current filters.</td>';
      body.appendChild(tr);
    } else {
      for (const r of slice) {
        const tr = document.createElement('tr');
        const ss = statusesOf(r);
        const hn = holidayNameOf(r);
        const cls = {'Not Received': 'notrec', 'Weekend': 'weekend', 'Holiday': 'holiday', 'Negative Paid': 'negpaid', 'Paid': 'paid'};
        const badges = ss.map(s => {
          const label = (s === 'Holiday' && hn) ? ('Holiday: ' + hn) : s;
          return '<span class="badge ' + cls[s] + '">' + label + '</span>';
        }).join('');
        const amtClass = 'amount' + (r.amount != null && r.amount < 0 ? ' negative' : '');
        tr.innerHTML =
          '<td>' + r.date + '</td>' +
          '<td>' + r.weekday + '</td>' +
          '<td>' + r.client + '</td>' +
          '<td class="' + amtClass + '">' + (r.amount == null ? '' : fmtMoney(r.amount)) + '</td>' +
          '<td>' + badges + '</td>';
        body.appendChild(tr);
      }
    }

    // Sort indicators
    for (const th of document.querySelectorAll('th[data-key]')) {
      th.classList.remove('sort-asc', 'sort-desc');
      const arrow = th.querySelector('.arrow');
      if (th.dataset.key === state.sortKey) {
        th.classList.add('sort-' + state.sortDir);
        arrow.textContent = state.sortDir === 'asc' ? '▲' : '▼';
      } else {
        arrow.textContent = '↕';
      }
    }

    // Pager
    $('pager-info').textContent = sorted.length === 0
      ? '0 rows'
      : (start + 1).toLocaleString() + '–' + (start + slice.length).toLocaleString() + ' of ' + sorted.length.toLocaleString();
  }

  function bindFilters() {
    const ids = ['f-client', 'f-year', 'f-month', 'f-comment', 'f-from', 'f-to', 'f-low'];
    const keys = ['client', 'year', 'month', 'comment', 'from', 'to', 'low'];
    ids.forEach((id, i) => {
      $(id).addEventListener('input', () => {
        state[keys[i]] = $(id).value;
        state.page = 0;
        render();
      });
    });
    $('btn-reset').addEventListener('click', () => {
      ids.forEach(id => $(id).value = '');
      Object.assign(state, { client: '', year: '', month: '', comment: '', from: '', to: '', low: '', page: 0 });
      render();
    });
    document.querySelectorAll('th[data-key]').forEach(th => {
      th.addEventListener('click', () => {
        const k = th.dataset.key;
        if (state.sortKey === k) {
          state.sortDir = state.sortDir === 'asc' ? 'desc' : 'asc';
        } else {
          state.sortKey = k;
          state.sortDir = (k === 'amount' || k === 'date') ? 'desc' : 'asc';
        }
        render();
      });
    });
    $('pg-prev').addEventListener('click', () => { state.page--; render(); });
    $('pg-next').addEventListener('click', () => { state.page++; render(); });
  }

  bindFilters();
  render();
})();
</script>
</body>
</html>
"""


def generate_html(window_start: date, window_end: date) -> str:
    rows = build_rows(window_start, window_end)
    payload = [
        {
            "client": r["client"],
            "date": r["date"].isoformat(),
            "weekday": r["weekday"],
            "amount": (round(r["amount"], 2) if r["amount"] is not None else None),
            "comment": r["comment"] or "",
            "year": r["year"],
            "month": r["month"],
        }
        for r in rows
    ]
    html = (HTML_TEMPLATE
            .replace("__GENERATED__", datetime.now().strftime("%Y-%m-%d %H:%M"))
            .replace("__WINDOW_START__", window_start.isoformat())
            .replace("__WINDOW_END__", window_end.isoformat())
            .replace("__ROW_COUNT__", f"{len(payload):,}")
            .replace("__DATA_JSON__", json.dumps(payload, separators=(",", ":"))))
    return html


def main():
    today = date.today()
    print(f"[info] Building HTML for {WINDOW_START} -> {today}")
    html = generate_html(WINDOW_START, today)
    primary = OUTPUT_PATHS[0]
    try:
        os.makedirs(os.path.dirname(primary), exist_ok=True)
        with open(primary, "w", encoding="utf-8") as f:
            f.write(html)
        print(f"[done] Wrote {primary}")
    except (PermissionError, OSError) as e:
        print(f"[warn] Couldn't write primary path: {e}")

    for path in OUTPUT_PATHS[1:]:
        try:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            shutil.copyfile(primary, path) if os.path.exists(primary) else open(path, "w", encoding="utf-8").write(html)
            print(f"[done] Copy: {path}")
        except (PermissionError, OSError) as e:
            print(f"[warn] Couldn't write {path}: {e}")


if __name__ == "__main__":
    main()
