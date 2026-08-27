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

## Servers already checked — don't re-hunt these

### SqlUtilAudit (= TRGUSER5A), checked 2026-08-27 — **does not have it**

Readable databases: `ACT`, `cmse_new`, `MIR`, `Scratch`, `SmartII`, `master`, `msdb`,
`tempdb`. Everything else on the instance returns `HAS_DBACCESS = 0` (including
`MemberResponse`, `PersonMatch`, `EnterpriseMasterData`).

- **`cmse_new` here is the same database as on TRGRepSQL3** — identical table list, row
  counts identical bar one row of replication lag on `ImportStaging`. Nothing extra.
- **No MMSEA objects anywhere.** No table/view/proc named `%MMSEA%` or `%Entitlement%`, and
  no module text referencing `MMSEA`, `EntitlementAgeCount`, `U M Count` or `# of Invs`.
- **The `MMSEA` database is not on this instance** — it is still only on TRGRepSQL3, still
  not accessible to `tls2`.
- **`MIR` is the closest system but cannot produce the numbers.** See below.

### MIR database (on SqlUtilAudit) — genuinely related, still not the source

`MIR` is the MMSEA response-file processing system: `MIRFile` (148 rows), `MIRRecord`
(68.6M), `MIREligibilityInfo` (1.0M), `PersonInvestigation` (299,652).

**Useful discovery: `MIR.dbo.MIRFile.CMSESourceLogId` joins straight to
`cmse_new..SourceLog.SourceLogId`.** But:

- It is **Aetna-only** — `FeedClientCode` is just `2-AHP` (TRADITIONAL, 127 files) and
  `2-AEHMO` (HMO, 21 files). The report covers 21 clients.
- Only **46** of its files overlap the 2025-08-04 export, all "Aetna Traditional AIS NonMSP".
- **`# of Invs` equals none of MIR's counts** — 0 of 46 against `RecordsReadAndRecorded`,
  `NetNewPersons`, `TotalProcessedPersons`, `TotalDistributablePersons`,
  `TotalDistributableLetters`.
- A ratio search allowing a *different* formula per load still matched only **6 of 46**
  (E %) and **5 of 46** (MC %). A real formula would match 46 of 46.
- Counting investigations through the record chain
  (`MIRFile → MIRRecord → PersonInvestigation`) gives ~65,000 per file where the report
  says ~1,600 — off by 40x (it counts every investigation those persons ever had).
- `MIRRecord.EntitlementFlag` (values A / G / B / blank) and
  `MedAEffectiveDate` / `MedBEffectiveDate` are the best semantic fits, but the rates are
  wrong: for CMSESourceLogId 34592 (report says E % 98.66, MC % 91.44) the flag is
  populated on 99.98% of records, Med A or B on 99.93%, and `(A+B)/total` = 91.28% — close
  to MC % but not equal, and it does not hold across loads.

### What this did settle — the column meanings

- **MC = Medicare Coverage.** The `MC_` column family in `SmartII.dbo.tblCOB2` and
  `Scratch..InvestigationRecap` makes it unambiguous: `MC_2728OnFile`,
  `MC_CoordinationPeriodApplies`, `MC_DiagnosisCode`, `MC_DialysisFirstDate`,
  `MC_ApprovedFacility`, `MC_DisabilityDate`.
- **E = Entitlement** (not "Eligibility"), matching `MIRRecord.EntitlementFlag` and
  cmse_new's `Entitlement{Age,Disability,Esrd}Count`.

Worth confirming both with Adam West rather than assuming.

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
