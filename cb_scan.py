"""Scan the VENDOR.CB-CLAIMS-EXTRACT CSV (from stdin) for 11/12/2025 pay dates.

Reads CSV from stdin (piped via `cat` in Git Bash, since the Windows python
child can't resolve the UNC mount directly). Reports counts + SUM(PAID_AMOUNT)
for target date 2025-11-12 across the claims-file date fields, plus a November
neighbor distribution to mirror yesterday's AetnaHRP review.
"""
import sys, csv

csv.field_size_limit(10_000_000)
TARGET = "20251112"

r = csv.reader(sys.stdin)
header = next(r)
idx = {name.strip().strip('"'): i for i, name in enumerate(header)}

# Date fields that exist in a claims extract (no PAYMENT_DATE here).
DATE_FIELDS = ["PAYABLE_DATE", "MOST_RECENT_PROCESS_DATE", "CLEAN_CLAIM_DATE"]
i_paid = idx.get("PAID_AMOUNT")
i_payable = idx.get("PAYABLE_DATE")

has_payment_date = "PAYMENT_DATE" in idx

cnt = {f: 0 for f in DATE_FIELDS}
amt = {f: 0.0 for f in DATE_FIELDS}
nov_cnt = {}   # PAYABLE_DATE 202511xx -> count
nov_amt = {}   # PAYABLE_DATE 202511xx -> sum PAID_AMOUNT
total = 0
bad_paid = 0


def to_amt(s):
    global bad_paid
    try:
        return float(s)
    except (ValueError, TypeError):
        bad_paid += 1
        return 0.0


for row in r:
    if len(row) <= i_paid:
        continue
    total += 1
    paid = to_amt(row[i_paid]) if i_paid is not None else 0.0
    for f in DATE_FIELDS:
        j = idx.get(f)
        if j is not None and j < len(row) and row[j] == TARGET:
            cnt[f] += 1
            amt[f] += paid
    # November 2025 distribution by PAYABLE_DATE
    if i_payable is not None and i_payable < len(row):
        pd = row[i_payable]
        if len(pd) == 8 and pd.startswith("202511"):
            nov_cnt[pd] = nov_cnt.get(pd, 0) + 1
            nov_amt[pd] = nov_amt.get(pd, 0.0) + paid

print(f"TOTAL_DATA_ROWS|{total}")
print(f"HAS_PAYMENT_DATE_COLUMN|{has_payment_date}")
print(f"BAD_PAID_VALUES|{bad_paid}")
for f in DATE_FIELDS:
    print(f"TARGET_{f}|{cnt[f]}|{amt[f]:.2f}")
print("--- November 2025 PAYABLE_DATE distribution ---")
for d in sorted(nov_cnt):
    print(f"NOV|{d}|{nov_cnt[d]}|{nov_amt[d]:.2f}")
