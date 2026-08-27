# Sourcing E % / MC % / # of Invs for the MMSEA_Report tab

**Status:** dashboard currently *copies* these three columns from Adam West's newest
emailed `MMSEA_Report_<YYYYMMDD>.xlsx`. That covers **112 of 372 loads (30%)** and is
stale by construction. This note is the ask needed to make it live and complete.

---

## What we established (so the ask doesn't get re-litigated)

These three columns are **not computable** from `TRGRepSQL3 / cmse_new`. Proof:

Comparing the **2025-06-05** and **2025-08-04** exports over the **782 loads present in
both with byte-identical count columns** (Import Count, Success, Failed, Age, Dis, ESRD,
U M Count):

| column | loads whose value changed |
|---|---|
| E % | 138 |
| MC % | 448 |
| # of Invs | 212 |

Every candidate input frozen, outputs moved. So they are **downstream MSP
investigation metrics that keep accruing after the load finishes** — not load-time
arithmetic.

Also ruled out, exhaustively:
- 2.1M-ratio search over 1,445 derived values (every subset sum, pairwise difference and
  base-minus-subset of the eight count columns plus the per-StagingStatus counts), under
  both round-half-up and truncation → **zero** formula pairs where E % and MC % share a
  denominator.
- Null test: invented targets (73.41, 55.28, 12.35) returned *more* matches than the real
  values, confirming every individual hit is coincidence.
- Best single-formula fit across all 782 loads: 12.5%, and that one is an artifact
  (`Success/ImportCount` = 100% whenever Failed = 0, overlapping the 98 loads where
  E % = 100).

Two loads that kill any record-level predicate on their own: **HNE MSP 32225** and
**GEHA MSP 31824** — near-identical ImportStaging profiles, yet 0%/0% vs 100%/100%.

---

## The ask — for Adam West

1. **What produces `E %`, `MC %` and `# of Invs` on the MMSEA_Report?** Specifically:
   - Which database and table/view do they come from?
   - Is there a stored proc or SSRS/Excel query behind the export we could be pointed at?
2. **What do they mean?** Our reading is that `MC %` is a Medicare-coverage match rate and
   `# of Invs` is investigations opened, both accruing per load over time — please
   confirm, because the tab needs a correct tooltip more than it needs a number.
3. **What is the denominator?** On the loads where they're non-zero it looks like the
   **ImportStaging row count**, not `SourceLog.RecordCount` (load 26733 is 162 vs 161).
4. **Can the export be scheduled, or better, can we read the source directly?** A monthly
   emailed xlsx means the dashboard is always up to a month behind and can never cover
   recent loads.

## The ask — for whoever owns DB permissions

Read access for `tls2` to the **`MMSEA` database on TRGRepSQL3**. It exists and is
visible, but the account cannot log into it. No proc in any readable database references
`ImportStaging`, so that database is the most likely home of these three metrics.

If access is granted, the fix is small: swap `read_mmsea_snapshot()` in
`cmse_report.py` for a `sql()` call and the tab goes to 100% coverage and live values.

---

## Meanwhile

- Snapshots are auto-ingested from `~\Documents`, `~\Downloads` and the Outlook attachment
  cache into `mmsea_snapshots/`. **Just leave Adam's email in Outlook** — the next
  dashboard run picks the attachment up on its own and coverage improves automatically.
- The Outlook attachment cache is volatile (hashed folders, purged without warning), which
  is why local copies are kept.
