"""
RAMP Client Delivery Status Report — calendar view.

Data sources (all combined per scheduled cell):
  - SQL [DHTStats].[DHT].[TableList] on TRGUTIL10  — canonical certification date
  - RAMP /api/Ramp/Snap/SnapQueueStatus            — snap completion (for snap-only clients)
  - RAMP /api/Ramp/Queue/List                      — load-job completion
  - RAMP /api/Ramp/Job/List                        — to detect Inactive
  - ADO WIQL                                       — tickets tagged 'Delivery Ticket'

Each (client, scheduled-day) cell in the calendar resolves to:
  - Date  (MM/DD)  if the client certified that day in DHT
  - "Snap"          if the client snapped that day but does not certify (snap-only)
  - "L"             if a load/snap is currently in progress for that client today
  - blank           otherwise

Client name suffix conventions (matching the All Clients tab key):
  - (s)  SLA Client
  - (p)  Rx Client Post Snap
  - (n)  Not Delivered (special)
  - M -  Monthly client prefix (placed dynamically on day ticket fired / snap completed)
"""
import calendar
import json
import os
import re
import subprocess
import tempfile
from collections import defaultdict
from datetime import date, datetime, timedelta

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ============================================================
#                    static configuration
# ============================================================
ADO_BASE   = "https://devops.ado.rawlingslou.prod/TFS2012/Rawlings"
ADO_LINK   = "https://devops.ado.rawlingslou.prod/TFS2012/AppDev/_workitems/edit/{}"
RAMP_BASE  = "http://ramp"
SQL_SERVER = "TRGUTIL10"
SQL_DB     = "DHTStats"
OUTPUT_DIR = r"\\trgfile1\Shared\DIG\Data Business Delivery Team\Delivery Schedule\Daily Status Reports"
# 2026 monthly schedule files maintained by the team. Their 'All Clients' tab
# is the source of truth for past months (Jan–Apr 2026); the live DHT/RAMP
# data we fetch has only 3 months of cert history and rolling queue/snap, so
# we overlay those files for closed months instead of re-deriving cells.
EXPECTED_DATES_DIR = r"\\trgfile1\Shared\DIG\Data Business Delivery Team\Delivery Schedule\2026"
EXPECTED_DATES_FILES = {
    1: "202601_ExpectedClientDates_JAN.xlsx",
    2: "202602_ExpectedClientDates_FEB.xlsx",
    3: "202603_ExpectedClientDates_MAR.xlsx",
    4: "202604_ExpectedClientDates_APR.xlsx",
}
# Marker strings the team uses in those files that should pink-shade the cell.
ALL_CLIENTS_ALERT_MARKERS = {"no data", "load failure", "missing files",
                             "deployment", "snap hold", "field changes",
                             "outreach", "empty", "disabled", "inactive",
                             "cleanup"}
# Clients whose empty/True date cells in the snapshot files should render as
# ✓ (per user 2026-06-03 — Feb–Apr files use boolean True as a verify-load
# placeholder for non-cert days, but the team reads them as "loaded/snapped").
ALL_CLIENTS_FILL_CHECKMARK = {
    "AetnaHRP", "AetnaRCE", "AetnaRx", "NCStateAetna",
    "CVSPBMRx", "ESIPBMRx", "PrimePBMRx", "MedImpactPBMRx",
}
# Project-folder copy (canonical filename — overwritten each run).
LOCAL_COPY_DIR = r"C:\Users\tls2\.claude\projects\H--"
# OneDrive copy with fixed filename so a single Notion link stays valid run-to-run.
ONEDRIVE_COPY_PATH = r"C:\Users\tls2\OneDrive - Machinify\Documents\Reports\ClientDeliveryStatus.xlsx"

# Recognise three title formats: Snap and Mine | Load and Snap | Kaiser - SNAP/MINE
TITLE_RE = re.compile(
    r"^\s*(Snap and Mine|Load and Snap|Kaiser\s*-\s*SNAP/MINE)\s*-\s*([^-\s][^-]*?)\s*(?:-|$)",
    re.I,
)

# ADO ticket client identifier → additional normalized strings to match against
# RAMP JobName (substring) / FeedName / ClientName (equality).
CLIENT_ALIASES = {
    "SamaritanHealth":      ["samaritan"],
    "HealthNetCA":          ["healthnet"],
    # JohnsHopkins RAMP feed is named "JHHC Medical" — without these aliases,
    # 'JHHC Medical 0110 Load' wouldn't match find_matching_jobs / snap_idx.
    # Added 2026-05-19 per user: "JohnsHopkins 'JHHC Medical 0110 Load' should be an 'L'".
    "JohnsHopkins":         ["johnshopkins", "jhhc", "jhhcmedical"],
    # JHHCPassfile (monthly): the 'JHHC Passfile Email' job → snap_idx key
    # "jhhcpassfileemail" (no digit code in the name). Distinct from
    # JohnsHopkins ('JHHC Medical 0110 Load'). Per user 2026-06-08.
    "JHHCPassfile":         ["jhhcpassfileemail"],
    # ElevanceMMMRx (daily): RAMP jobs are 'ElevanceMMMRx Masterload 0110 Load' /
    # '0120 Snap'. build_snap_index strips the trailing "load" off "Masterload",
    # so the snap-index key is "elevancemmmrxmaster" — alias it so snap_on_day's
    # strict-equality match fires for ✓. Per user 2026-06-16: daily ✓ on load+snap.
    "ElevanceMMMRx":        ["elevancemmmrxmaster"],
    "BCBSNorthCarolinaFEP": ["bcbsncfep"],
    # NCStateAetna's daily load runs through 'Aetna RCE 310 ETL Load' (Feed
    # "RCE Medical" → key "aetnarce" after stripping). Include aetnarce alias
    # so NCStateAetna gets ✓ from those same daily completions.
    "NCStateAetna":         ["ncstateaetna", "aetnarce"],
    "WPRxDMGCOBMining":     ["wellpointedwardrxdmgcobmining", "wpwedmgcobmining"],
    "HumanaRx":             ["humanarx"],
    # OptumPBMRx: RAMP load is "Optum 0110 PBM Load" → digit-collapsed
    # "optumpbmload"; alias "optumpbm" gives the substring match for both
    # is_loading_today and snap_idx lookups.
    "OptumPBMRx":           ["optumpbm"],
    # GEHA: the actual RAMP load is "GEHA UMR 0110 Load" → prefix "gehaumr".
    # Plain "geha" alone wouldn't strict-match "gehaumr" in snap_idx.
    "GEHA":                 ["geha", "gehaumr"],
    "EmblemFacets":         ["emblemfacets"],
    "AetnaQNXT":            ["aetnaqnxt"],
    "AetnaQNXTRx":          ["aetnaqnxtrx"],
    "ElixirRx":             ["elixirrx"],
    "Tufts_PublicPlan":     ["tuftspublicplan"],
    "MedicalMutualMHS":     ["medicalmutualmhs"],
    # WellpointEdwardRx — RAMP jobs are named "Wellpoint RX Claims …" /
    # "Wellpoint RX Claims HealthSun …" so we need the prefix-derived
    # snap_idx keys to be reachable via strict-equality aliases. Per user
    # 2026-05-20: "'Wellpoint RX Claims HealthSun 0130 Load' is running,
    # so there should be an 'L' on WellpointEdwardRx."
    "WellpointEdwardRx":    ["wellpointedwardrx", "wellpointrxclaims",
                             "wellpointrxclaimshealthsun"],
    # WellpointRxElig (monthly 'M - WellpointRx Elig'): the Wellpoint RX Elig
    # pipeline (0100 Stage -> 0110 Load -> 0120 Snap -> 0130 Mine). Both the
    # Load and Snap snap_idx entries derive the prefix key "wellpointrxelig"
    # (distinguished by kind). Per user 2026-06-12: L off '0110 Load', ✓ off
    # '0120 Snap'. Distinct from WellpointEdwardRx (Claims) jobs.
    "WellpointRxElig":      ["wellpointrxelig"],
    # HarvardPilgrim — RAMP job is "HarvardPilgrim Claims 0110 Load" so
    # the snap_idx prefix key is "harvardpilgrimclaims" (the trailing-
    # word strip doesn't catch "Claims"). Per user 2026-05-20:
    # "'HarvardPilgrim Claims 0110 Load' ran last night, so there should
    # be an 'L' until certification."
    "HarvardPilgrim":       ["harvardpilgrim", "harvardpilgrimclaims"],
    # UPMC — RAMP job "UPMC Masterload 0110 Load". The trailing-keyword
    # strip incorrectly takes "load" off "Masterload" leaving snap_idx
    # key "upmcmaster". Per user 2026-05-20: "'UPMC Masterload 0110 Load'
    # finished, so it should show an 'L' until certification."
    "UPMC":                 ["upmc", "upmcmaster", "upmcmasterload"],
    # BCBSAR is a MEDICAL client — its only jobs are 'BCBSAR Medical 0100/0110'.
    # The bare "bcbsar" alias was dropped 2026-06-12 because it substring-matched
    # the sibling BCBSARRx jobs ('BCBSARRx COBC/MasterLoad …'), wrongly pulling
    # the BCBSARRx COBC load failure onto BCBSAR. Matched via "bcbsarmedical"
    # only; see CLIENT_PRIMARY_KEY_OVERRIDE. (DHT cert key "bcbsar" is still
    # yielded by _keys_for_client's base normalize, so cert lookup is unaffected.)
    "BCBSAR":               ["bcbsarmedical"],
    "BCBSARRx":             ["bcbsarrx"],
    "MedStar":              ["medstar"],
    # HealthNewEngland shows up as "HNE Medical" on the RAMP Dashboard.
    "HealthNewEngland":     ["healthnewengland", "hnemedical", "hne"],
    # AetnaRx queue variants: each distinct JobName prefix becomes its own
    # snap_idx key under JobName-only indexing.
    # "aetnarxclaims" covers the snap step "AetnaRx Claims 0130 Start Snap"
    # (plural Claims) — distinct from the load step "AetnaRX Claim 0120 Load".
    "AetnaRx":              ["aetnarx", "aetnarxo20", "aetnarxclaim", "aetnarxclaims",
                             "aetnarxnewelig", "aetnarxmining",
                             "aetnarxcobcmasterload", "aetnarxcobc",
                             "aetnarxcaqh", "aetnarxtrr", "aetnarxihp"],
    # CenteneFidelis / CenteneFidelisRx job prefixes
    "CenteneFidelis":       ["centenefidelis", "centenefidelismedical"],
    "CenteneFidelisRx":     ["centenefidelisrx", "centenefidelisrxmasterload"],
    # WellCare / WellCareRx job prefixes
    # 'wellcare' as a substring also lives inside 'wellcarerx', so we use
    # CLIENT_PRIMARY_KEY_OVERRIDE to swap the auto-derived primary key for
    # 'wellcaremedical'. The substring match then no longer crosses into
    # WellCareRx territory.
    "WellCare":             ["wellcaremedical"],
    # WellCareRx: per user 2026-05-18, ancillary jobs (COBC, ABII) are NOT
    # load indicators — exclude their snap_idx keys from matching so that a
    # successful COBC load this week does not trip the "loaded this week" L.
    "WellCareRx":           ["wellcarerx", "wellcarerxmasterload"],
    # Oscar / OscarRx job prefixes
    "Oscar":                ["oscar", "oscarmedical"],
    "OscarRx":              ["oscarrx", "oscarrxmasterload", "oscarrxabii"],
    # Centene / CenteneRx job prefixes
    "Centene":              ["centene", "centenemedical"],
    "CenteneRx":            ["centenerx", "centenerxhnt", "centenerxhntelig"],
    # AetnaHRP / AetnaRCE job prefixes
    "AetnaHRP":             ["aetnahrp"],
    "AetnaRCE":             ["aetnarce"],
    "AetnaQNXTRx":          ["aetnaqnxtrx", "aetnaqnxtrxmasterload"],
    "AetnaQNXT":            ["aetnaqnxt", "aetnaqnxtmasterload", "aetnaqnxtmspi", "aetnaqnxtcaqh"],
    # NCStateAetna alias to aetnarce (uses same ETL load)
    "NCStateAetna":         ["ncstateaetna", "aetnarce", "ncstateaetnamasterload"],
    # ESIPBMRx 0120 Start Snap → "esipbmrx"
    "ESIPBMRx":             ["esipbmrx"],
    # CareSource variants
    "CareSource":           ["caresource"],
    "CareSourceRx":         ["caresourcerx"],
    # Kaiser WA: ONLY exact "kaiserwa" — keep separate from KaiserWARx etc.
    "Kaiser_WA":            ["kaiserwa"],
    "Kaiser_WARx":          ["kaiserwarx"],
    # KaiserPrePayCOB
    "KaiserPrePayCOB":      ["kaiserprepaycob", "kaiserpareoprepay"],
    # Kaiser SC/NC Pareo — RAMP feeds are "Kaiser Pareo SC" and
    # "Kaiser Pareo NC&TPMG" so the normalized form has "pareo" before
    # the state code. Without these aliases find_matching_jobs misses them.
    # Per user 2026-05-19: "KaiserSCPareo & KaiserNCPareo should have an 'L'".
    "KaiserSCPareo":        ["kaiserscpareo", "kaiserpareosc"],
    "KaiserNCPareo":        ["kaiserncpareo", "kaiserpareonc"],
    # Kaiser_MASTapestry — RAMP renamed feed to "Kaiser Pareo MAS" per user
    # 2026-06-03. Keeps default `kaisermastapestry` match for historical
    # 'Kaiser MAS Tapestry' JobNames; adds `kaiserpareomas` for the new form.
    "Kaiser_MASTapestry":   ["kaisermastapestry", "kaiserpareomas"],
    # MMOH (WC) monthly: the report's "MMOH (WC)" row tracks the Workers' Comp
    # load 'MMO 0110 WC Load' (Stage->Load, no snap/cert). build_snap_index
    # emits the feed+sub-feed key "mmowc" for it (the bare-prefix key "mmo" is
    # too short to index). Per user 2026-06-08 (correction): this row is the WC
    # load, not the MMOHRx monthly claim. Distinct from MedicalMutualOH
    # ('MMOH Claims 0110 Load') and the weekly MMOHRx ('MMOHRx Weekly Claim').
    "MMOH":                 ["mmowc"],
    # MedicalMutualOH (monthly, cert-only): its load is 'MMOH Claims 0110 Load'
    # → snap_idx key "mmohclaims" (per user 2026-06-08). The auto-derived
    # primary key "medicalmutualoh" never matches the RAMP "MMOH ..." JobNames.
    "MedicalMutualOH":      ["mmohclaims"],
    # MMOHRx weekly Tue: only 'MMOHRx Weekly Claim 0110 Load' counts. COBC
    # alias dropped 2026-05-18 so MMOHRx COBC Successful loads don't trip L.
    "MMOHRx":               ["mmohrx", "mmohrxweeklyclaim"],
    # MMOHRxMonthly (monthly 'M - MMOHRx'): the MMOHRx Monthly Claim pipeline
    # (Stage->Load 0110->Snap 0120). Both the Load and Snap snap_idx entries
    # derive key "mmohrxmonthlyclaim" (distinguished by kind). Per user
    # 2026-06-09: L off '0110 Load', ✓ off '0120 Snap'. Distinct from the
    # weekly MMOHRx (Weekly Claim) and from MMOH (WC).
    "MMOHRxMonthly":        ["mmohrxmonthlyclaim"],
    # Cigna variants
    "CignaFacets":          ["cignafacets"],
    # CignaRx: only 'Cigna RX 0110 Load' counts (key 'cignarx'). COBC and
    # Daily PassFile have their own snap_idx keys and shouldn't trigger L
    # for CignaRx — per user 2026-05-18.
    "CignaRx":              ["cignarx"],
    "CignaPower":           ["cignapower"],
    "CignaProClaims":       ["cignaproclaims"],
    # Premera / PremeraMedAdv*
    "Premera":              ["premera"],
    "PremeraMedAdvVIS":     ["premeramedadvvis"],
    "PremeraMedAdvRx":      ["premeramedadvrx"],
    # Tufts variants
    "TuftsMedPref":         ["tuftsmedpref"],
    "Tufts_Audit_CIT":      ["tuftsauditcit"],
    "WebTPA":               ["webtpa"],
    "BCBSPuertoRico":       ["bcbspuertorico"],
    "NCState":              ["ncstate"],
    "MedImpactPBMRx":       ["medimpactpbmrx"],
    "CenteneRx":            ["centenerx"],
    # BCBSFLEligibilityLoad: the RAMP job is "BCBSFL Eligibility ..." → key
    # "bcbsfleligibility" after stripping load/stage/digits.
    "BCBSFLEligibilityLoad": ["bcbsfleligibility"],
    # Kaiser monthly aliases
    "Kaiser_GE":            ["kaiserge"],
    "Kaiser_AmbCO":         ["kaiserambulanceco", "kaiserambco"],
    "Kaiser_AmbGA":         ["kaiserambulancega", "kaiserambga"],
    "Kaiser_AmbHI":         ["kaiserambulancehi", "kaiserambhi"],
    "Kaiser_AmbNW":         ["kaiserambulancenw", "kaiserambnw"],
    "Kaiser_AmbN":          ["kaiserambulancenc", "kaiserambn"],
    "Kaiser_AmbS":          ["kaiserambulancesc", "kaiserambs"],
    "Kaiser_AmbM":          ["kaiserambulancemas", "kaiserambulancema", "kaiserambm"],
}

# --------- canonical client lists (from the manual ExpectedClientDates sheet) ---------
# Daily clients (load + snap every weekday; cert cadence varies).
DAILY_CLIENTS = ["AetnaHRP", "AetnaRCE", "AetnaRx", "ElevanceMMMRx", "NCStateAetna"]
KAISER_PREPAY_CLIENT = "KaiserPrePayCOB"

# Fixed display order for the Daily section (per user 2026-06-16) — NOT
# alphabetical. Every daily client shows on every weekday (blank or "-" when it
# didn't load). Matched by label prefix so client suffixes like "(s)"/"(p)"
# still resolve. "Kaiser Submission" sits between KaiserPrePayCOB and
# NCStateAetna. (Aetna MSPI moved to Monthly Ad Hoc 2026-06-25.)
DAILY_ORDER = [
    "AetnaHRP",
    "AetnaRCE",
    "AetnaRx",
    "ElevanceMMMRx",
    "KaiserPrePayCOB",
    "Kaiser Submission",
    "NCStateAetna",
]


def _daily_order_key(label):
    """Sort key mapping a daily row label to its fixed-order index."""
    for idx, name in enumerate(DAILY_ORDER):
        if label.startswith(name):
            return idx
    return len(DAILY_ORDER)

# Weekly clients keyed by canonical name -> list of weekday names they deliver on.
WEEKLY_CLIENTS = {
    # === MONDAY ===
    "BCBSKSMedAdv":          ["Monday"],
    "Cambia":                ["Monday"],
    "CignaPower":            ["Monday"],
    "CignaProClaims":        ["Monday"],
    "CVSPBMRx":              ["Monday"],
    "EverNorthRx":           ["Monday"],
    "GEHA":                  ["Monday"],
    "HealthNetCA":           ["Monday"],
    "MedicaDean":            ["Monday"],
    "Tufts_Audit_CIT":       ["Monday"],
    "TuftsMedPref":          ["Monday"],
    "TuftsRx":               ["Monday"],
    # === TUESDAY ===
    "BCBSAR":                ["Tuesday"],
    "BCBSARRx":              ["Tuesday"],
    "BCBSFL":                ["Tuesday"],
    "Centene":               ["Tuesday"],
    "CenteneQualChoice":     ["Tuesday"],
    "CignaFacets":           ["Tuesday"],
    "CignaRx":               ["Tuesday"],
    "HMSA":                  ["Tuesday"],
    "HMSA_Rx":               ["Tuesday"],
    "JohnsHopkins":          ["Tuesday"],
    "MedStar":               ["Tuesday"],
    "MMOHRx":                ["Tuesday"],
    "UPMC":                  ["Tuesday"],
    "Wellmark":              ["Tuesday"],
    # === WEDNESDAY ===
    "CareSource":            ["Wednesday"],
    "CenteneFidelis":        ["Wednesday"],
    "CenteneFidelisRx":      ["Wednesday"],
    "EmblemRx":              ["Wednesday"],
    "ExcellusRx":            ["Wednesday"],
    "HarvardPilgrim":        ["Wednesday"],
    "Medica":                ["Wednesday"],
    "Oscar":                 ["Wednesday"],
    "WellpointEdwardRx":     ["Wednesday"],
    # === THURSDAY ===
    "HealthNewEngland":      ["Thursday"],
    "Kaiser_CO":             ["Thursday"],
    "Kaiser_GA":             ["Thursday"],
    "Kaiser_HI":             ["Thursday"],
    "Kaiser_MASTapestry":    ["Thursday"],
    "CareSourceRx":          ["Thursday"],
    "Kaiser_NW":             ["Thursday"],
    "KaiserNCPareo":         ["Thursday"],
    "KaiserSCPareo":         ["Thursday"],
    "Premera":               ["Thursday"],
    "PrimePBMRx":            ["Thursday"],
    # === FRIDAY ===
    "CenteneRx":             ["Friday"],
    "OscarRx":               ["Friday"],
    "WebTPA":                ["Friday"],
    "WellCare":              ["Friday"],
    "WellCareRx":            ["Friday"],
    # Snap-only Monday slots kept from earlier user instructions:
    # ESIPBMRx is MONTHLY (handled by MONTHLY_CLIENTS) — not weekly.
    # OptumPBMRx is monthly (loaded once per month via TRGETL3 tape RAW1/2/3) —
    # placed dynamically by determine_monthly() on the actual tape-load date.
}

# Clients that should always show as Inactive (pink shade) regardless of
# RAMP/DHT detection. User-confirmed list.
# HealthNetCA added 2026-05-18: 'HealthNet 0100 Claims Stage' disabled in RAMP.
# Kaiser_AmbM removed 2026-05-19 (no longer inactive). Snap re-enabled
# 2026-06-08 — now treated as a normal monthly Kaiser_Amb cert feed (see
# MONTHLY_CERT_ONLY_CLIENTS / SNAP_KIND_ONLY_CLIENTS).
FORCED_INACTIVE = {"Tufts_PublicPlan", "TuftsRx", "HealthNetCA"}

# Clients whose load is running but snap step is disabled in RAMP — show
# marker "Snap" with pink shading on the expected delivery day. Mechanism kept
# wired for future use; Kaiser_AmbM removed 2026-06-08 (snap re-enabled).
SNAP_DISABLED_CLIENTS = set()

# For certain clients, the "L" (currently loading) indicator should only fire
# when a Ready/Running job's JobName contains one of the listed substrings.
# Ancillary jobs (COBC, IHP, ABII, etc.) running do NOT mean the client is
# loading. Per user 2026-05-18: "A WellCareRx COBC job ran, but that is not
# the indicator for Loading. We only want to look for MasterLoad or Claims Load."
LOAD_NAME_REQUIRED = {
    # AetnaRx: only the main "Claim 0110 Split Load" / "Claim 0120 Load" /
    # "Claim 0130 Start Snap" / "MasterLoad" steps count. Ancillary
    # "Claim 0150 RTA Load", "0132 ETL4 O20 Load", "IHP", "COBC" etc.
    # running Ready does NOT count as the client currently loading.
    # "claim 0130" added 2026-05-20: when the 0130 Start Snap completes
    # in early morning (attribution date = today, after the 0120 Load
    # finishes overnight on yesterday's date), today's cell gets ✓ from
    # the snap step. Per user: "AetnaRx should be a checkmark for today."
    "AetnaRx":           ("claim 0110", "claim 0120", "claim 0130", "masterload"),
    # BCBSFL (weekly Tue): only 'BCBSFL 0110 Claims Load' counts — not CMS
    # Referral Load, Claims Stage, or Claims Start Snap.
    "BCBSFL":            ("claims load",),
    # MMOH (WC) monthly: only 'MMO 0110 WC Load' is the load indicator (per
    # user 2026-06-08 correction — this row is the Workers' Comp load). The
    # "wc load" keyword matches it but not the WC Stage step or any other
    # MMOH* job (MMOH Claims / GEN / ICD10 / MMOHRx...).
    "MMOH":              ("wc load",),
    # MedicalMutualOH (monthly, cert-only): only 'MMOH Claims 0110 Load' counts
    # (plural "Claims" — distinct from the MMOHRx monthly/weekly singular).
    "MedicalMutualOH":   ("claims 0110 load",),
    # MMOHRx weekly Tue: only 'MMOHRx Weekly Claim 0110 Load' counts —
    # filter excludes the Monthly Claim Stage and Weekly Claim Stage.
    "MMOHRx":            ("weekly claim 0110 load",),
    # MMOHRxMonthly: the 'MMOHRx Monthly Claim' Load (0110) and Snap (0120)
    # steps. "monthly claim" matches both (excludes Weekly Claim / COBC); the
    # Stage step is dropped by is_loading_today's stage filter, and ✓ comes
    # from the Snap step via SNAP_KIND_ONLY_CLIENTS.
    "MMOHRxMonthly":     ("monthly claim",),
    # CignaRx (weekly Tue): only 'Cigna RX 0110 Load' counts. COBC Load,
    # Daily PassFile, and other ancillary jobs share the "cignarx" matching
    # prefix and would otherwise trip the L indicator. Per user 2026-05-19:
    # "'Cigna RX 0110 Load' is not running, so CignaRx should not have an 'L'."
    # Pattern "rx 0110 load" matches "Cigna RX 0110 Load" but not
    # "Cigna RX COBC 0110 Load" (the "COBC" between "RX" and "0110" breaks
    # the substring).
    "CignaRx":           ("rx 0110 load",),
    # ElevanceMMMRx (daily): only the Masterload load + snap steps signal
    # delivery. Excludes the ancillary 'ElevanceMMMRx COBC 0110 Load' (and the
    # 0100 Stage / 0130 Post Snap) from both is_loading_today and the ✓ lookup.
    # Per user 2026-06-16.
    "ElevanceMMMRx":     ("masterload 0110 load", "masterload 0120 snap"),
    # WellCareRx: narrowed to the Masterload Load step 2026-06-08 (was
    # "masterload","claim") so 'WellCareRx Masterload 0100 Stage' no longer
    # trips L via the snap-index activity path. Delivery = Masterload 0110 Load.
    "WellCareRx":        ("masterload 0110 load",),
    # BCBSARRx: delivery = 'BCBSARRx MasterLoad 0110 Load'. Narrowed 2026-06-12
    # so the ancillary 'BCBSARRx COBC 0110 Load' failure no longer trips a
    # "Load Failure" (nor L) for the claims cycle. Per user: the failure is COBC,
    # not claims.
    "BCBSARRx":          ("masterload 0110 load",),
    # OscarRx: main load is 'Oscar RX 0110 Load' (no "claim"/"masterload"
    # in the name). Added "rx 0110 load" 2026-05-20 per user:
    # "Oscar Rx 0110 Load is running and was not picked up."
    "OscarRx":           ("masterload", "claim", "rx 0110 load"),
    # CenteneRx: narrowed to the Claims Load step 2026-06-08 — bare "claim"
    # matched 'Centene RX 0130 Claims Stage'. Delivery = '0140 Claims Load'.
    "CenteneRx":         ("claims load",),
    # CenteneFidelisRx: only 'Centene Fidelis Rx 0130 Claims Load' (and the
    # MasterLoad variant) should drive L. Narrowed from "claim" to "claims
    # load" 2026-06-08 — the bare "claim" matched 'Centene Fidelis Rx 0120
    # Claims Stage', which (as a completed entry in the snap index) tripped the
    # cert-client "activity -> L" path. The is_loading_today stage filter
    # already excluded a Ready/Running stage, but the snap-index activity path
    # had no stage guard, so the narrower keyword is the fix.
    "CenteneFidelisRx":  ("masterload", "claims load"),
    # AetnaQNXT / AetnaQNXTRx: narrowed to the Masterload Load step 2026-06-08
    # so 'Masterload 0100 Stage' no longer trips L via the activity path.
    "AetnaQNXTRx":       ("masterload 0110 load",),
    "AetnaQNXT":         ("masterload 0110 load",),
    # WellpointEdwardRx: narrowed to the two Claims Load steps 2026-06-08 so
    # 'Claims 0100 Stage' / 'Claims HealthSun 0120 Stage' no longer trip L.
    "WellpointEdwardRx": ("claims 0110 load", "claims healthsun 0130 load"),
    # WellpointRxElig: only the Elig pipeline counts. "rx elig" matches BOTH
    # 'Wellpoint RX Elig 0110 Load' (→ L) and 'Wellpoint RX Elig 0120 Snap'
    # (→ ✓ via the snap-index path), but excludes the Claims/COBC jobs.
    "WellpointRxElig":   ("rx elig",),
    # HAPRx (per user 2026-06-04): only the main Claims load counts —
    # `HAPRx 0110 Load` and any future `HAPRx Masterload 0110 ...`. The
    # substring `haprx 0110 load` matches the main load but NOT the COBC /
    # TPLCov variants (`HAPRx COBC 0110 Load`, `HAPRx TPLCov 0110 Load`)
    # because the modifier breaks the substring — same trick as CignaRx.
    "HAPRx":             ("haprx 0110 load", "haprx masterload 0110"),
    # ElixirRx and PremeraMedAdvRx (per user 2026-06-04): only Claim Load
    # jobs trigger L; snap / mine / SFTP jobs must not. Patterns match
    # `Premera MedAdv Rx Claims 0110 Load` and any future Elixir Claims
    # 0110 Load variant.
    "ElixirRx":          ("claim 0110 load", "claims 0110 load", "masterload"),
    "PremeraMedAdvRx":   ("claim 0110 load", "claims 0110 load", "masterload"),
    # BCBSNC (per user 2026-06-05): only the main `BCBSNC Claims 0110 Load`
    # counts. Excludes CAQH, COBC, Claris Health, MSPI, Rx, Daily Passfile —
    # all of which have their own 0110 Load lines.
    "BCBSNC":            ("bcbsnc claims 0110 load",),
}

# Manual cell overrides — (client, scheduled_date) → marker. Marker can be:
#   - a date object (rendered as MM/DD/YY)
#   - one of the marker strings: "✓", "L", "No Data", "Load Failure",
#     "Inactive", "Deployment", "" (blank)
# Use sparingly — only for one-off corrections that the data sources can't
# express on their own (e.g. retroactively assigning a cert date to a Friday
# cell, or marking a known deployment-blocked Wednesday).
MANUAL_OVERRIDES = {
    ("AetnaHRP",  date(2026, 5, 1)): "✓",
    ("WebTPA",    date(2026, 5, 1)): "No Data",
    ("CenteneRx", date(2026, 5, 1)): date(2026, 5, 5),
    ("CenteneRx", date(2026, 5, 8)): date(2026, 5, 5),
    ("Medica",    date(2026, 5, 6)): "Deployment",
    # 2026-05-20: CignaFacets 5/12 Tue cycle certified 5/19 (Mon, outside the
    # default Mon-Fri 5/11-5/15 backward window). Per user: "missing past
    # dates … CignaFacets on 5/12/26." Pin the late cert explicitly.
    ("CignaFacets", date(2026, 5, 12)): date(2026, 5, 19),
    # 2026-05-21: Premera certified 5/20 14:32 (Wed, one day before its Thu
    # scheduled cycle). Forward cert window 5/21-5/27 misses it; pin the
    # cert date explicitly. Per user: "Premera Load was certified on 5/20."
    ("Premera",      date(2026, 5, 21)): date(2026, 5, 20),
    # 2026-05-28: ExcellusRx 5/20 finished loading and certified today;
    # pin 5/28 as the cert date for the 5/20 cycle cell.
    ("ExcellusRx",   date(2026, 5, 20)): date(2026, 5, 28),
    # 2026-05-29: ExcellusRx 5/27 cycle loaded and certified today.
    ("ExcellusRx",   date(2026, 5, 27)): date(2026, 5, 29),
    # 2026-05-28: CenteneRx 5/22 cycle certified today alongside the 5/29
    # load job. Pin 5/28 cert date for the 5/22 cell. (5/29 cell will pick
    # up the cert naturally via DHT detection.)
    ("CenteneRx",    date(2026, 5, 22)): date(2026, 5, 28),
    # 2026-05-28: KaiserSCPareo certified today (Thu). Pin the cert date
    # explicitly so the cell shows the cert regardless of when DHT indexes.
    ("KaiserSCPareo", date(2026, 5, 28)): date(2026, 5, 28),
    # 2026-05-28: Premera Commercial 0110 Load finished + certified today.
    ("Premera",       date(2026, 5, 28)): date(2026, 5, 28),
    # 2026-06-01: AetnaRx 6/1 was a dupe of 5/31 files; flipped to ✓ on
    # 2026-06-05 once the associated data landed.
    ("AetnaRx",       date(2026, 6, 1)): "✓",
    # 2026-06-03: AetnaRx Snap failed — staff will manually correct. Force ✓
    # for today; remove this override once the next day's run picks up the
    # corrected snap completion naturally.
    ("AetnaRx",       date(2026, 6, 3)): "✓",
    # Kaiser_MASTapestry / KaiserSCPareo blank-until-cert overrides cleared
    # 2026-06-04 — both certified, auto-detection now surfaces the date.
    # 2026-06-04: TuftsRx Mon 6/1 — force back to Inactive. The past-day ✓
    # came from a stale snap entry that beat the FORCED_INACTIVE fallback.
    ("TuftsRx",       date(2026, 6, 1)): "Inactive",
    # 2026-06-08: BCBSFL weekly delivery skipped this week (one-off). Show
    # "Skip" on the 6/9/26 Tuesday cell only. Remove after this week.
    ("BCBSFL",        date(2026, 6, 9)): "Skip",
    # 2026-06-09: CignaFacets certified last week per user (TapeID 3593,
    # [dbo].[vwMiningCache_Full], CertTimestamp 6/2). DHT.TableList row is still
    # at "Email sent, Ready for Certification review" (not "Certified"), so the
    # auto cert lookup misses it and the 6/2 cell flagged with "!". Pin the
    # 6/2/26 cert date. Remove once DHT flips the 3593 row to Certified (then
    # cert_in_week surfaces it automatically).
    ("CignaFacets",   date(2026, 6, 2)): date(2026, 6, 2),
    # 2026-06-16: CignaFacets 6/9 cycle (TapeID 3598) was Certified and shown as
    # 6/9/26, but DHT reverted it to "Email sent, Ready for Certification review"
    # (a known client issue), so the live cert lookup misses it and the cell went
    # to a pink "!". Pin the cert date back. Also seeds the sticky-cert cache so
    # the regression can't recur. Per user: "Once a client gets certified, do not
    # change the cell to an '!'." Remove once DHT flips 3598 back to Certified.
    ("CignaFacets",   date(2026, 6, 9)): date(2026, 6, 9),
    # WellCare / WellCareRx delivery tracking (updated 2026-06-24):
    #  - WellCare Medical: 6/12 cell certified 6/17; 6/19 cell certified 6/18.
    #  - WellCareRx loaded 6/12+6/19 together and CERTIFIED 6/24 (per user; DHT
    #    6/24 cert covers StatTimestamp weeks 6/19 + 6/24). Per user, pin 6/24 on
    #    the 6/12, 6/19 AND 6/26 cells — the combined load's 6/12 week has no own
    #    StatTimestamp so 6/12 needs the explicit pin; 6/19 & 6/26 pinned too.
    ("WellCare",      date(2026, 6, 12)): date(2026, 6, 17),
    ("WellCare",      date(2026, 6, 19)): date(2026, 6, 18),
    ("WellCareRx",    date(2026, 6, 12)): date(2026, 6, 24),
    ("WellCareRx",    date(2026, 6, 19)): date(2026, 6, 24),
    ("WellCareRx",    date(2026, 6, 26)): date(2026, 6, 24),
    # 2026-06-24: CenteneRx/OscarRx 6/22 certs land on the 6/19 cell
    # automatically via the StatTimestamp system. The 6/26 blank overrides were
    # REMOVED 2026-06-24 — OscarRx & CenteneRx are now LOADING their 6/26
    # delivery and must show "L" (the blanks were hiding it).
}

# --- Sticky certifications --------------------------------------------------
# Once a (client, scheduled-day) cell has rendered a real cert date, remember it
# so a later DHT status reversion (a row flipping from "Certified" back to e.g.
# "Email sent, Ready for Certification review") can't regress the cell to a
# blank / "No Data" pink "!". Per user 2026-06-16: "Once a client gets certified,
# do not change the cell to an '!'." Live certs always win (a fresh cert date
# overwrites the remembered one); the cache only fills in when the live lookup
# would otherwise leave the cell empty. Persisted next to the script as JSON
# {"client|YYYY-MM-DD": "YYYY-MM-DD"}. Keyed by exact cell day, so it's reliable
# for fixed-weekday weekly/daily clients (monthly cells, which move to the cert
# date, are best-effort).
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CERT_STICKY_PATH = os.path.join(_SCRIPT_DIR, "cert_sticky_cache.json")
STICKY_CERTS = {}


def load_sticky_certs():
    global STICKY_CERTS
    try:
        with open(CERT_STICKY_PATH, "r", encoding="utf-8") as f:
            STICKY_CERTS = json.load(f)
    except (OSError, ValueError):
        STICKY_CERTS = {}
    return STICKY_CERTS


def save_sticky_certs():
    try:
        with open(CERT_STICKY_PATH, "w", encoding="utf-8") as f:
            json.dump(STICKY_CERTS, f, indent=0, sort_keys=True)
    except OSError as e:
        print(f"[warn] couldn't write sticky cert cache: {e}")


def apply_sticky_cert(client, day, marker, alert, from_manual):
    """Record live cert dates and restore a remembered cert if the freshly
    computed cell would regress to a pink "!".

    - `marker` is a date (a real cert) → remember it; live cert always wins.
    - Otherwise, when the marker is NOT a manual-override value and is one of the
      states that renders "!" (blank or "No Data"), restore a remembered cert
      date for this (client, day) and clear the alert. Manual overrides (e.g. an
      intentional blank) are left untouched.
    """
    key = f"{client}|{day.isoformat()}"
    if isinstance(marker, date):
        STICKY_CERTS[key] = marker.isoformat()
        return marker, alert
    if from_manual:
        return marker, alert
    if marker in ("", "No Data") and key in STICKY_CERTS:
        try:
            return date.fromisoformat(STICKY_CERTS[key]), False
        except ValueError:
            pass
    return marker, alert


# ADO ticket IDs to hyperlink onto specific Load-Failure cells. Keyed by
# (client, day) — same convention as MANUAL_OVERRIDES. When a cell renders
# "Load Failure" AND has an entry here, the marker text becomes a clickable
# link to the TFS work item. Per user 2026-05-20: "For Load failures that
# have an ADO, like 954657 for Centene Medical, added as a link to the
# 'Load Failure' comment."
LOAD_FAILURE_ADO_LINKS = {
    ("ExcellusRx", date(2026, 5, 20)): 955578,  # 'Excellus - Rx - ExcellusRx 0110 Load'
    ("AetnaHRP",   date(2026, 5, 22)): 956353,  # 'Aetna 0110 HRP Load' failure
    # 2026-05-26: current-week Load-Failure links per user.
    ("CignaFacets", date(2026, 5, 26)): 956575,  # 'ProdSupp - Cigna - Audit/Subro - Cigna Facets 0110 Load Failure'
    ("ExcellusRx",  date(2026, 5, 27)): 955578,  # same active ExcellusRx 0110 Load story
    # Centene removed 2026-05-21 — 0110 Claims Load restarted (Ready 14:46),
    # is_loading_today now auto-returns "L".
    # MMOHRx removed 2026-05-21 — Weekly Claim 0110 Load finished.
}

# Monthly clients whose placement (day AND/OR marker) should be forced,
# overriding determine_monthly's auto-detection. Keyed by client. Used for
# one-off corrections: e.g. EDW feeds certified late but anchored to their
# expected day, or AetnaQNXT mid-cycle visible on a specific day.
# Value: (placement_date, marker). marker may be a date object (cert-style)
# or a string ("L", "Load Failure", "No Data", etc.).
MONTHLY_PLACEMENT_OVERRIDES = {
    # 2026-05-21: EDW feeds certified this morning at 07:54-07:56 but user
    # wants them anchored to 5/20 (their expected delivery day). Marker
    # shows the actual 5/21 cert date.
    "EDW_ASE":    (date(2026, 5, 20), date(2026, 5, 21)),
    "EDW_C_FAC":  (date(2026, 5, 20), date(2026, 5, 21)),
    "EDW_C_NAS":  (date(2026, 5, 20), date(2026, 5, 21)),
    "EDW_Empire": (date(2026, 5, 20), date(2026, 5, 21)),
    # EDW_WGS certified 2026-05-22 per user. Anchored to 5/20 (expected
    # delivery day); marker shows the actual 5/22 cert date.
    "EDW_WGS":    (date(2026, 5, 20), date(2026, 5, 22)),
    # 2026-05-21: AetnaQNXT — Masterload started 5/18; anchored to 5/19.
    # "AUTO" marker: cell shows L while pending, then auto-updates to the
    # actual cert date once it lands — but placement stays on 5/19 even
    # though the cert may arrive on a different day (per user 2026-05-22:
    # "At the next update, AetnaQNXT will have been certified. Please
    # leave it on the 5/19 date.").
    "AetnaQNXT":  (date(2026, 5, 19), "AUTO"),
    # Kaiser_AmbM override removed 2026-06-08 — snap re-enabled; it now follows
    # the standard Kaiser_Amb cert-only placement (Thursday anchor).
}

# Extra rows injected into the calendar after standard placement runs. Use for
# one-off catch-up entries that don't fit the regular weekly/monthly cadence.
# Tuple: (section, day, label, marker, alert, highlight)
#   section ∈ {"daily", "weekly", "monthly", "kaiser"}
ADDITIONAL_ENTRIES = [
    # Medica catch-up for 5/1/26 claims — certified 2026-05-18 09:13:42
    # (DHT). Display cert date in the Mon cell.
    ("weekly", date(2026, 5, 18), "Medica (5/1/26)", date(2026, 5, 18), False, None),
    # EverNorthRx backsweep files 21, 22, 23 certified 2026-05-22 per user.
    ("weekly", date(2026, 5, 22), "EverNorthRx (21,22,23 BS)", date(2026, 5, 22), False, None),
]

# CignaRx EOM/SOM cycle — at the start of each month a second CignaRx cycle
# closes out the prior month's tail; user 2026-06-03: "typically marked as an
# exception after certification." A `CignaRx (EOM/SOM)` row is injected on
# the first Tuesday of every month (matching regular CignaRx Tuesday). Marker
# is auto-detected from cert/load activity in a window straddling the month
# boundary; override here when the cert is recorded as an exception.
# Key: (year, month) of the SOM side. Value: a date (cert date) OR a marker
# string ("✓", "L", "No Data", "Load Failure", "").
CIGNARX_EOM_SOM_OVERRIDES = {
    # ("2026,6": date(2026, 6, ?)),  # populate when the exception cert lands
}

# Per-client cert window direction (default = backward / same Mon-Fri week).
# "forward" = look forward 7 days from scheduled day (used when a cert that
# lands after the scheduled day belongs to that scheduled day's cycle, like
# Premera where a Mon 5/11 cert completes the previous Thursday 5/7 cycle).
CERT_DIRECTION = {
    "Premera": "forward",
}

# Monthly clients that must remain anchored to their expected day even after
# a snap/load completes — only a DHT cert moves them. Per-user spec:
# "BCBSKS & BCBSKSMedAdv Monthly clients should always be on the 15th".
# 2026-05-15: Kaiser_Amb* feeds (CO/GA/HI/N/NW/S) added — user wants them
# anchored to the Thursday cert day even while loading. Kaiser_AmbM added
# 2026-06-08 (snap re-enabled) — now certifies with the rest of the feeds.
MONTHLY_CERT_ONLY_CLIENTS = {
    "BCBSKS", "BCBSKSMedAdv", "BCBSSCRx", "CareFirstRx",
    "Kaiser_AmbCO", "Kaiser_AmbGA", "Kaiser_AmbHI",
    "Kaiser_AmbM",
    "Kaiser_AmbN", "Kaiser_AmbNW", "Kaiser_AmbS",
    # Kaiser_WA: per user 2026-05-18, load completion alone is not delivery —
    # the cell should stay L on the expected day until the cycle truly
    # completes. Previously in LOAD_AS_DELIVERY, which auto-✓'d on a Successful
    # load even when the actual data was empty/incomplete.
    "Kaiser_WA",
    # Christus: per user 2026-06-01, stays L until DHT cert lands — snap
    # completion alone should not flip the cell to ✓.
    "Christus",
    # BCBSNorthCarolinaFEP: per user 2026-06-02, no ✓ from snap activity —
    # stays L on the expected Friday until DHT cert lands.
    "BCBSNorthCarolinaFEP",
    # BSCA_Facets: per user 2026-06-03, snap activity alone must not surface
    # ✓ — stays L (or No Data) until cert lands.
    "BSCA_Facets",
    # HealthSpring_FWA: per user 2026-06-03, same rule — loaded today but
    # cert will arrive separately; no snap-only ✓.
    "HealthSpring_FWA",
    # ElixirRx and PremeraMedAdvRx per user 2026-06-04: only Claim Load
    # job and DHT certification should move them off No Data / L. A snap
    # alone must not surface a ✓ for either client.
    "ElixirRx", "PremeraMedAdvRx",
    # BCBSVT and BSCA_Medicare per user 2026-06-08: leave as "L" until
    # certification — snap activity alone must not flip the cell to ✓.
    "BCBSVT", "BSCA_Medicare",
    # HAP_Medical and HAPRx per user 2026-06-08: same — stay "L" until cert.
    "HAP_Medical", "HAPRx",
    # Per user 2026-06-08: the following monthly clients should never show a
    # checkmark — stay "L" until the cert date lands (snap activity alone must
    # not flip to ✓). New 'MasterLoad 0110 Load' implementations also default
    # to this cert-style behavior (see the auto-discovery block in main()).
    "AetnaQNXT", "AetnaQNXTRx",
    "BCBSNC", "BCBSNC_Rx", "BCBSPuertoRico", "BCBSSC",
    "CareFirstDC", "CareFirstFacets", "CareFirstNasco",
    "Chickering",
    "EDW_ASE", "EDW_C_FAC", "EDW_C_NAS", "EDW_Empire", "EDW_WGS",
    "EmblemFacets",
    "Kaiser_WARx",
    "MedicalMutualMHS", "MedicalMutualOH",
    "NCStateRx",
    "PremeraMedAdvVIS",
    "SamaritanHealth",
}

# Monthly clients that should show an empty Date cell (rather than "No Data")
# until the cert lands. Per-user: "M - BCBSKSMedAdv had data, so have it blank
# each month until the ticket, Snap, Certification process finish on/near 15th".
MONTHLY_BLANK_UNTIL_CERT = {"BCBSKSMedAdv"}

# Monthly clients whose "No Data" should always be shaded regardless of the
# 7-day grace window (e.g. ElixirRx hasn't received data in a long time).
FORCE_SHADE_NO_DATA = {"ElixirRx", "MedicalMutualMHS"}

# Clients that should be rendered with bold label (no fill).
BOLD_LABEL = {"Aetna NMSP - MMSEA", "AetnaMMSEA"}
# Clients with yellow label fill (kept empty per user — Aetna NMSP changed to bold only).
YELLOW_HIGHLIGHT = set()

MONTHLY_CLIENTS = {
    "AetnaQNXT", "AetnaQNXTRx", "AetnaRx_LegacyDMG", "AetnaSubro",
    "BCBSFLEligibilityLoad", "BCBSKS", "BCBSKSMedAdv",
    "BCBSNC", "BCBSNC_Rx", "BCBSNorthCarolinaFEP",
    "BCBSPuertoRico", "BCBSSC", "BCBSSCRx", "BCBSVT",
    "BSCA_Facets", "BSCA_Medicare",
    "CareFirstDC", "CareFirstFacets", "CareFirstNasco", "CareFirstRx",
    "Chickering", "Christus",
    "EDW_ASE", "EDW_C_FAC", "EDW_C_NAS", "EDW_Empire", "EDW_WGS",
    "ElixirRx",
    "EmblemFacets",
    "HAP_Medical", "HAPRx", "HealthSpring_FWA", "HumanaRx",
    "JHHCPassfile",
    "Kaiser_AmbCO", "Kaiser_AmbGA", "Kaiser_AmbHI", "Kaiser_AmbM", "Kaiser_AmbN",
    "Kaiser_AmbNW", "Kaiser_AmbS",
    "Kaiser_GE",
    "Kaiser_WA", "Kaiser_WARx",
    "MedicalMutualMHS", "MedicalMutualOH", "MedImpactPBMRx",
    "MMOH", "MMOHRxMonthly", "NCState", "NCStateRx",
    "WellpointRxElig",                  # monthly load->snap (L on load, ✓ on snap)
    "ESIPBMRx",                         # monthly snap-only (RAMP snap-driven)
    "OptumPBMRx",                       # monthly, tape-driven
    "PremeraMedAdvRx", "PremeraMedAdvVIS",
    "SamaritanHealth", "Tufts_PublicPlan", "TuftsRx",
}

# Ad-hoc MONTHLY cert clients (per user 2026-06-25): appear ONCE per month on
# their DHT certification date, ONLY after they certify — no fixed expected day,
# never flagged "No Data"/missing before cert. Deliberately NOT in
# MONTHLY_CLIENTS (that would flag them pre-cert). Value = DHT DatabaseName key.
ADHOC_MONTHLY_CERT_CLIENTS = {
    "UnitedRx": "UnitedRx",   # ticket certifying ~2026-06-25; surfaces on cert
}

# Override display name for a client (the label only; client_key stays the same).
CLIENT_DISPLAY_NAME = {
    "BCBSFLEligibilityLoad": "BCBSFL Elig",
    "MMOH":                  "MMOH (WC)",
    "MMOHRxMonthly":         "MMOHRx",
    "WellpointRxElig":       "WellpointRx Elig",
    "JHHCPassfile":          "JHHC Passfile",
    "Kaiser_AmbCO":          "KaiserAmbCO",
    "Kaiser_AmbGA":          "KaiserAmbGA",
    "Kaiser_AmbHI":          "KaiserAmbHI",
    "Kaiser_AmbM":           "Kaiser_AmbM",
    "Kaiser_AmbN":           "KaiserAmbN",
    "Kaiser_AmbNW":          "KaiserAmbNW",
    "Kaiser_AmbS":           "Kaiser_AmbS",
    "Kaiser_GE":             "KaiserGE",
}

# Expected delivery window per monthly client, as (start_day, end_day) of month.
# Used for both placement (end of range) and "late" detection (today > end + 7).
MONTHLY_EXPECTED_DAY_RANGE = {
    "Chickering":             (1, 1),
    "Christus":               (1, 1),
    "MedicalMutualMHS":       (1, 1),
    "NCStateRx":              (1, 1),
    "MedicalMutualOH":        (3, 8),
    "MedImpactPBMRx":         (5, 10),
    "AetnaQNXTRx":            (5, 10),
    "BCBSVT":                 (5, 10),
    "BSCA_Facets":            (5, 10),
    "BSCA_Medicare":          (5, 10),
    "HAP_Medical":            (5, 10),
    "HAPRx":                  (5, 10),
    "HealthSpring_FWA":       (5, 10),
    "MMOH":                   (5, 10),
    "NCState":                (5, 10),
    # WellpointRxElig loads ~11th, snaps ~12th each month (per RAMP history).
    "WellpointRxElig":        (10, 12),
    "PremeraMedAdvVIS":       (5, 10),
    "PremeraMedAdvRx":        (5, 10),
    "TuftsRx":                (10, 10),
    "AetnaQNXT":              (10, 15),
    "AetnaSubro":             (11, 11),
    "BCBSNC":                 (10, 15),
    "BCBSNorthCarolinaFEP":   (10, 15),
    "BCBSPuertoRico":         (10, 15),
    "ElixirRx":               (10, 15),
    "Kaiser_WA":              (10, 15),
    "Kaiser_WARx":            (10, 15),
    "SamaritanHealth":        (10, 15),
    "Tufts_PublicPlan":       (10, 15),
    "BCBSKS":                 (15, 15),
    "BCBSKSMedAdv":           (15, 15),
    "BCBSNC_Rx":              (15, 15),
    "AetnaMMSEA":             (15, 15),
    "AetnaRx_LegacyDMG":      (16, 16),
    "BCBSSC":                 (15, 20),
    # CareFirst clients moved to 19 to spread out from the 5/20 cluster
    "CareFirstDC":            (15, 19),
    "CareFirstFacets":        (15, 19),
    "CareFirstNasco":         (15, 19),
    "CareFirstRx":            (15, 19),
    "EmblemFacets":           (15, 20),
    "EDW_ASE":                (20, 20),
    "EDW_C_FAC":              (20, 20),
    "EDW_C_NAS":              (20, 20),
    "EDW_Empire":             (20, 20),
    "EDW_WGS":                (20, 20),
    "BCBSFLEligibilityLoad":  (25, 25),
    "AetnaRx_LegacyDMG":      (16, 16),
    # BCBSSCRx delays one week — per user, this month's load belongs next week.
    "BCBSSCRx":               (18, 19),
    # Kaiser monthly clients
    "Kaiser_GE":              (15, 20),
    "Kaiser_AmbCO":           (21, 21),
    "Kaiser_AmbGA":           (21, 21),
    "Kaiser_AmbHI":           (21, 21),
    "Kaiser_AmbM":            (21, 21),
    "Kaiser_AmbN":            (21, 21),
    "Kaiser_AmbNW":           (21, 21),
    "Kaiser_AmbS":            (21, 21),
    # HumanaRx: no fixed expected day — placed on snap date or today when loading.
}

# Spread monthly clients across the Mon-Fri week of their anchor day.
# Weekday is 0=Mon, 1=Tue, 2=Wed, 3=Thu, 4=Fri. The override snaps
# expected_date to the same Mon-Fri work-week as the original anchor day.
# Clients NOT in this dict stay on their MONTHLY_EXPECTED_DAY_RANGE end day.
# Per user 2026-05-26: avoid stacking everyone on the 10th and 15th. Pinned
# clients (BCBSKS, BCBSKSMedAdv, Kaiser_WA, BCBSSCRx, CareFirst*) are left
# off this list — they keep their explicit anchor.
MONTHLY_PLACEMENT_WEEKDAY = {
    # Day-10 cluster — spread Mon-Fri of the work-week of the 10th.
    "AetnaQNXTRx":            0,  # Mon
    "BCBSVT":                 1,  # Tue
    "BSCA_Facets":            2,  # Wed
    "BSCA_Medicare":          3,  # Thu
    "HAP_Medical":            4,  # Fri
    "HAPRx":                  0,  # Mon
    "HealthSpring_FWA":       1,  # Tue
    "MMOH":                   2,  # Wed
    "MedImpactPBMRx":         3,  # Thu
    "NCState":                4,  # Fri
    "PremeraMedAdvRx":        0,  # Mon
    "PremeraMedAdvVIS":       1,  # Tue
    "TuftsRx":                2,  # Wed
    # Day-15 cluster — spread Mon-Fri of the work-week of the 15th.
    # (AetnaMMSEA omitted — handled separately by nmsp_mmsea_date.)
    "AetnaQNXT":              1,  # Tue
    "BCBSNC":                 2,  # Wed
    "BCBSNC_Rx":              3,  # Thu
    "BCBSNorthCarolinaFEP":   4,  # Fri
    "BCBSPuertoRico":         0,  # Mon
    "ElixirRx":               1,  # Tue
    "Kaiser_WARx":            2,  # Wed
    "SamaritanHealth":        3,  # Thu
    "Tufts_PublicPlan":       4,  # Fri
}

# Monthly clients whose anchor day should snap to the CLOSEST weekday
# (Sat → previous Fri, Sun → next Mon) instead of the default
# next_monday_if_weekend (Sat → next Mon). Per user 2026-05-26: BCBSFL Elig
# anchor 25th should land on the closest weekday in months where 25 is Sat/Sun.
CLOSEST_WEEKDAY_CLIENTS = {"BCBSFLEligibilityLoad"}

# Clients in implementation phase — render "Implementation" in the date cell
# on every scheduled day from start_date until the first DHT cert lands. Cells
# before start_date are suppressed entirely.
# Per user 2026-05-26: BCBSAR was a new Tuesday weekly implementation from June.
# BCBSAR removed 2026-06-08 — Implementation marker dropped; certifying this
# week, so it now behaves as a normal weekly Tuesday cert-style client.
# Mechanism kept wired for the next implementation client.
IMPLEMENTATION_CLIENTS = {}

# Weekly clients whose cells BEFORE a given date render empty (no row, no pink
# "!") — for clients with no real history before they went live. Keyed by
# client → cutoff date; any scheduled day strictly before the cutoff is skipped.
# Per user 2026-06-08: BCBSAR — blank all dates before this week (Mon 6/8/26).
BLANK_BEFORE = {
    "BCBSAR": date(2026, 6, 8),
    # ElevanceMMMRx — only show June 2026 forward (per user 2026-06-16); earlier
    # cells were implementation-phase noise.
    "ElevanceMMMRx": date(2026, 6, 1),
}

# Clients whose "is delivered" signal is exclusively from TRGETL3 tape loads.
# Lookups for these clients ignore RAMP snap entries entirely.
TAPE_ONLY_CLIENTS = {"OptumPBMRx", "ESIPBMRx", "MedImpactPBMRx"}

# Jobs that signal a delivery but whose name lacks the usual load/stage/snap/
# mine keyword, so build_snap_index would otherwise skip them. Matched on the
# lowercased, stripped JobName. Per user 2026-06-08: 'JHHC Passfile Email'
# completion is the JHHCPassfile monthly ✓.
EXTRA_INDEXED_JOBS = {"jhhc passfile email"}

# Snap destination filter — when a client uses a specific snap destination,
# only count snap entries matching that destination string.
# (MMOH's Pharmacy filter removed 2026-06-08 — MMOH is now the WC load, which
# delivers via 'MMO 0110 WC Load' completion, not a Pharmacy snap.)
SNAP_DESTINATION_FILTER = {
}

# Clients that show ✓ when snapped (not just blank when no cert).
# Combines:  daily clients (✓ on snap-only days), PBMRx clients, and the
# small set of "select" snap-deliverable clients the user named.
SNAP_ONLY_CLIENTS = {
    "OptumPBMRx",
    "ESIPBMRx", "CVSPBMRx",
    "MedImpactPBMRx", "PrimePBMRx",
    "AetnaSubro", "HumanaRx",
    "WPRxDMGCOBMining",
    "BCBSKSMedAdv", "TuftsRx",   # snap weekly, cert monthly
    "NCState",                   # blocked from DHT cert by Chimera; track via snap
}

# These clients ONLY get ✓ when an actual SNAP step completes — a load-step
# completion alone doesn't trigger ✓.
SNAP_KIND_ONLY_CLIENTS = {
    # PBMRx snap clients
    "ESIPBMRx", "MedImpactPBMRx", "PrimePBMRx", "CVSPBMRx",
    # Other snap-required monthly/weekly clients
    # (MMOH removed 2026-06-08 — now the WC load, ✓ on load completion not snap)
    "AetnaSubro", "TuftsRx", "NCState", "WPRxDMGCOBMining",
    # MMOHRxMonthly: ✓ off the 'MMOHRx Monthly Claim 0120 Snap' step (per user
    # 2026-06-09) — L during the 0110 Load, ✓ when the Snap completes.
    "MMOHRxMonthly",
    # WellpointRxElig: ✓ off the 'Wellpoint RX Elig 0120 Snap' step (per user
    # 2026-06-12) — L during the 0110 Load, ✓ at the Snap step.
    "WellpointRxElig",
    # Kaiser_GE needs snap-step completion (0120 Snap).
    "Kaiser_GE",
    # Kaiser ambulance feeds: per user 2026-05-15, must wait for an actual
    # snap step (Kaiser Ambulance NC/CO/GA/HI/MAS/NW/S 0120 Snap) — a load-step
    # completion alone leaves the cell in "L" (load done, snap pending).
    # Kaiser_AmbM added 2026-06-08 (snap re-enabled).
    "Kaiser_AmbCO", "Kaiser_AmbGA", "Kaiser_AmbHI", "Kaiser_AmbM",
    "Kaiser_AmbN", "Kaiser_AmbNW", "Kaiser_AmbS",
    # AetnaHRP added 2026-05-19 per user — snap step must complete; the load
    # alone is not delivery (cell stays "L" between load done and snap done).
    "AetnaHRP",
    # NOTE: Daily Aetna clients AetnaRCE, AetnaRx, NCStateAetna are still
    # NOT in this set — their ✓ fires on the respective Load job completion
    # (Aetna RCE 310 ETL Load / AetnaRX Claim 0120 Load), not the snap step.
}

# These clients get ✓ on LOAD completion (load = delivery for them).
# BCBSKSMedAdv weekly: ✓ after 'BCBSKS Med Adv 0110 Load' finishes.
# Daily Aetna clients (AetnaRCE, AetnaRx, NCStateAetna): ✓ on the LOAD step
# — the subsequent Start Snap step should NOT keep the cell in "L".
# AetnaHRP REMOVED 2026-05-19 per user: "AetnaHRP did not Snap yet from the
# 5/18/26 load. The 5/18/26 HRP should still be an 'L'." AetnaHRP now requires
# the snap step to complete — see SNAP_KIND_ONLY_CLIENTS.
LOAD_AS_DELIVERY_CLIENTS = {
    "OptumPBMRx", "HumanaRx", "BCBSKSMedAdv",
    "AetnaRCE", "AetnaRx", "NCStateAetna",
}

# Clients whose "L" is driven ONLY by the LOAD step — a running snap/mine step
# does NOT keep them at L. Unlike LOAD_AS_DELIVERY_CLIENTS, this set does NOT
# affect ✓ resolution (so these clients can still be SNAP_KIND_ONLY and only
# get ✓ from an actual snap completion). Used solely by is_loading_today.
# Per user 2026-06-12 for WellpointRxElig: L during '0110 Load', then ✓ at the
# '0120 Snap' — the post-snap '0130 Mine' running must NOT revert ✓ back to L.
L_ON_LOAD_ONLY_CLIENTS = {
    "WellpointRxElig",
}

# Clients we are NOT actively working (no certification expected), but whose
# load pipeline is running for implementation/testing. Cell behavior:
#   - currently loading → "L"
#   - load+snap finished → blank (NOT a ✓, NOT a cert date)
# New auto-discovered MasterLoad implementations that are NOT PBMRx default
# into this set; promote out of it once the client is being actively delivered.
# (ElevanceMMMRx promoted out 2026-06-16 — now a DAILY_CLIENTS client with ✓ on
# load+snap days; see DAILY_CLIENTS / CLIENT_ALIASES / LOAD_NAME_REQUIRED.)
IMPLEMENTATION_LOAD_ONLY_CLIENTS = set()

# Override the auto-derived primary key for clients whose name is a substring
# of another client's name (causing spurious substring matches in
# find_matching_jobs). Per user 2026-05-18: WellCare jobs were detected from
# WellCareRx Ready entries because 'wellcare' ⊂ 'wellcarerx'.
CLIENT_PRIMARY_KEY_OVERRIDE = {
    "WellCare": "wellcaremedical",
    # AetnaQNXT's normalized key "aetnaqnxt" is a prefix-substring of
    # "aetnaqnxtrx", so find_matching_jobs would wrongly match AetnaQNXTRx's
    # jobs. Override to "aetnaqnxtmasterload" 2026-06-08 — matches AetnaQNXT's
    # Masterload jobs (its delivery; CAQH/MSPI are ancillary) but NOT any
    # AetnaQNXTRx job ("aetnaqnxt" + "rx" + "masterload" is not contiguous).
    # _keys_for_client still yields base "aetnaqnxt" for the DHT cert lookup.
    "AetnaQNXT": "aetnaqnxtmasterload",
    # BCBSAR (Medical) ⊂ BCBSARRx — without this, "bcbsar" substring-matched the
    # BCBSARRx jobs and pulled the BCBSARRx COBC load failure onto BCBSAR. Match
    # only the Medical jobs. Per user 2026-06-12. _keys_for_client still yields
    # base "bcbsar" for the DHT cert lookup.
    "BCBSAR": "bcbsarmedical",
    # "medica" is a substring of "medical", so find_matching_jobs("Medica")
    # falsely matched every "*Medical*" load (Centene Medical, WellCare Medical,
    # BCBSAR Medical, …) — a failed 'WellCare Medical 0110 Load' was showing up
    # as a "Load Failure" on Medica's cell (per user 2026-06-16: "Medica 0110
    # Load ran w/o error"). Override to "medica0", which matches Medica's own
    # step jobs ('Medica 0110 Load' → "medica0110load") but NOT "…medical0…".
    # _keys_for_client still yields base "medica" for DHT cert / snap-index
    # lookups (those use strict equality against the "medica" prefix key).
    "Medica": "medica0",
}

# NYShip_Rx fires four times per month — on the 1st, 8th, 16th, 24th
# (or the next Monday if that date is a weekend).
NYSHIP_DAYS = [1, 8, 16, 24]
NYSHIP_LABEL = {1: "1st", 8: "8th", 16: "16th", 24: "24th"}

# One-off per-cell overrides for the NYShip_Rx rotation, keyed by
# (year, month, daynum). Value is a marker string ("L") or a date(...) cert
# date (renders MM/DD/YY). Bypasses resolve_marker for that specific cell.
# Per user 2026-06-08: the 1st & 8th loaded together this cycle — both marked
# "L", and both get the SAME certification date once it lands. One-off; clear
# these entries after this June 2026 cycle certifies.
NYSHIP_OVERRIDES = {
    # 1st & 8th loaded together this cycle and both certified 6/9/26 (DHT
    # NYSHIP_RX, status Certified). Both cells show the shared cert date. The
    # 1st cell needs the override because a 6/9 cert is outside its 6/1-6/5
    # auto-detection window. One-off; clear after the June 2026 cycle.
    (2026, 6, 1): date(2026, 6, 9),
    (2026, 6, 8): date(2026, 6, 9),
    # 2026-06-23: the 16th load ran 6/22 (for the 6/16 data) but has NOT
    # certified yet (DHT NYSHIP_RX still only shows the 6/9 cert), so show "L".
    # The next load is for the 24th. Swap to the cert date once it certifies.
    # Per user.
    (2026, 6, 16): "L",
}

# Suffix conventions per the All Clients tab key:
#   (s) SLA Client | (p) Rx Post Snap | (n) Not Delivered
CLIENT_SUFFIXES = {
    "AetnaHRP":              "(s)",
    "AetnaRCE":              "(s)",
    "AetnaRx":               "(p)(s)",
    "CareSource":            "(s)",
    "CareSourceRx":          "(n)",
    "CenteneFidelisRx":      "(p)",
    "CenteneRx":             "(p)",
    "CignaRx":               "(p)",
    "EmblemRx":              "(p)",
    "EverNorthRx":           "(p)",
    "ExcellusRx":            "(p)",
    "BCBSARRx":              "(p)",
    "HMSA_Rx":               "(n)",
    "WellCareRx":            "(p)",
    "MMOHRx":                "(n)(p)",
    "KaiserPrePayCOB":       "(s)",
    "OscarRx":               "(n)(p)",
    "PremeraMedAdvRx":       "(n)(p)",
    "NCStateRx":             "(n)(p)",
    "Wellmark":              "(s)",
    "EmblemFacets":          "(s)",
    "HealthSpring_FWA":      "(s)",
    "Cambia":                "(n)",
    "HAPRx":                 "(p)",
    "P32-TuftsRx":           "(p)",
    # Monthly-specific suffixes per user spec 2026-05-13
    "TuftsRx":               "(p)",
    "AetnaSubro":            "(n)",
    "BCBSNC_Rx":             "(p)",
    "CareFirstRx":           "(n)(p)",
    "BCBSSCRx":              "(n)(p)",
    "EDW_ASE":               "(n)",
    "EDW_C_FAC":             "(n)",
    "EDW_C_NAS":             "(n)",
    "EDW_Empire":            "(n)",
    "EDW_WGS":               "(n)",
    "ElixirRx":              "(p)",
    "Kaiser_WARx":           "(n)",
}

WEEKDAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]

# ---------- Client Owner tab ----------
# Per user 2026-05-20. Each owner has an "upper" list (active/current
# clients) and a "lower" list (separated by a blank row in the source).
# "*" suffix preserved verbatim from the user-provided list.
# Priority: 1 = highest, 4 = lowest.
CLIENT_OWNERS = {
    "Dave": {
        "upper": [
            ("Centene*", 1),
            ("Anthem/Wellpoint*", 2),
            ("BCBSNC*", 2),
            ("Humana*", 2),
            ("Arkansas Blue*", 3),
            ("Christus", 3),
            ("CVSPBM Rx", 3),
            ("Elixir/MCS", 3),
            ("FrontRunner", 3),
            ("HMSA*", 3),
            ("Medstar", 3),
            ("Molina", 3),
            ("NYSHIP", 3),
            ("Waystar", 3),
        ],
        "lower": [
            ("BCBSMN*", 4),
            ("BCBSLA*", 4),
            ("MVP*", 4),
            ("HealthPartners*", 4),
        ],
    },
    "Emmanuel": {
        "upper": [
            ("Aetna", 1),
            ("Evernorth Rx", 1),
            ("Excellus", 2),
            ("Medical Mutual OH*", 2),
            ("Oscar*", 2),
            ("Johns Hopkins*", 3),
            ("CareSource", 3),
            ("ESI PBM Rx", 3),
            ("HAP", 3),
            ("Ingenio", 3),
            ("Maxor", 3),
            ("Medica", 3),
            ("Work Comp", 3),
            ("United", 4),
        ],
        "lower": [
            ("BCBS_Assoc*", 4),
            ("IndepenenceHealth*", 4),
            ("BCBSND*", 4),
            ("Highmark*", 4),
        ],
    },
    "Holly": {
        "upper": [
            ("Kaiser*", 1),
            ("BSCA*", 2),
            ("Emblem*", 2),
            ("Point 32 (Tufts/Harvard Pilgrim)*", 2),
            ("Premera*", 2),
            ("Wellmark*", 2),
            ("BCBSKS", 3),
            ("BCBSSC*", 3),
            ("HealthNewEngland", 3),
            ("Medispan", 3),
            ("UPMC", 3),
            ("WebTPA", 3),
            ("NPI", 4),
        ],
        "lower": [
            ("CapitalBlueCross*", 4),
            ("BlueCrossIdaho*", 4),
            ("BCBSRI*", 4),
            ("KPS*", 4),
        ],
    },
    "Adam": {
        "upper": [
            ("Cigna*", 1),
            ("BCBSFL*", 2),
            ("CareFirst*", 2),
            ("GEHA", 2),
            ("BCBS Puerto Rico", 3),
            ("BCBSVT*", 3),
            ("HealthNow*", 3),
            ("Medimpact PBM Rx", 3),
            ("Optum", 3),
            ("Prime PBM Rx", 3),
            ("Samaritan Health", 3),
            ("Cambia", 4),
            ("Provider Solutions", 4),
        ],
        "lower": [
            ("HCSC (Cigna)*", 4),
            ("BCBSTN*", 3),
            ("BCBSMA*", 4),
        ],
    },
}


# ============================================================
#                          helpers
# ============================================================
def curl_json(url):
    r = subprocess.run(
        ["curl", "-s", "--ntlm", "-u", ":", url],
        capture_output=True, text=True, check=False,
    )
    return json.loads(r.stdout)


def curl_post_json(url, body):
    with tempfile.NamedTemporaryFile("w", suffix=".json", delete=False, encoding="utf-8") as f:
        json.dump(body, f)
        path = f.name
    try:
        r = subprocess.run(
            ["curl", "-s", "--ntlm", "-u", ":",
             "-H", "Content-Type: application/json",
             "--data-binary", f"@{path}", url],
            capture_output=True, text=True, check=False,
        )
        return json.loads(r.stdout)
    finally:
        os.unlink(path)


def normalize(s):
    return re.sub(r"[\s_\-]+", "", s or "").lower()


def parse_dt(s):
    if not s:
        return None
    s = s.replace("Z", "").rstrip()
    try:
        return datetime.fromisoformat(s.split(".")[0])
    except Exception:
        return None


# ============================================================
#                       SQL cert fetch
# ============================================================
def fetch_dht_certs(since):
    """Query [DHTStats].[DHT].[TableList] for certifications since `since`.
    Returns a list of dicts: {DatabaseName, Name, CertTimestamp, CurrentStatus, PCN}.
    """
    q = (
        "SET NOCOUNT ON; "
        "SELECT DatabaseName, [Name], PCN, CertTimestamp, CurrentStatus, StatTimestamp "
        f"FROM [DHTStats].[DHT].[TableList] "
        f"WHERE CertTimestamp >= '{since.isoformat()}' "
        "ORDER BY CertTimestamp"
    )
    r = subprocess.run(
        ["sqlcmd", "-S", SQL_SERVER, "-d", SQL_DB, "-E", "-Q", q,
         "-W", "-s", "\t", "-h", "-1"],
        capture_output=True, text=True, check=False,
    )
    rows = []
    for line in r.stdout.splitlines():
        if not line or line.startswith("---") or "rows affected" in line:
            continue
        parts = line.split("\t")
        if len(parts) < 5:
            continue
        try:
            ts = parse_dt(parts[3])
        except Exception:
            continue
        if not ts:
            continue
        rows.append({
            "DatabaseName":   parts[0].strip(),
            "Name":           parts[1].strip(),
            "PCN":            parts[2].strip(),
            "CertTimestamp":  ts,
            "CurrentStatus":  parts[4].strip(),
            "StatTimestamp":  parse_dt(parts[5]) if len(parts) > 5 else None,
        })
    return rows


def build_cert_index(certs):
    """Return: normalized_db -> list of (datetime, status) sorted asc."""
    idx = defaultdict(list)
    for c in certs:
        key = normalize(c["DatabaseName"])
        if key:
            idx[key].append((c["CertTimestamp"], c["CurrentStatus"]))
    for k in idx:
        idx[k].sort()
    return idx


def stat_week_monday(d):
    """Monday of the delivery week a StatTimestamp (data date) belongs to.

    Per user 2026-06-24, validated against DHT: the delivery week runs
    Mon-Sat, and SUNDAY rolls forward into the next week. So a Sat data date
    stays with the week that just ended; a Sun data date belongs to the
    upcoming week.
      - Thu 6/18 -> Mon 6/15  (that week)
      - Sat 6/20 -> Mon 6/15  (week just ended; CenteneRx -> 6/19 cell)
      - Sun 6/21 -> Mon 6/22  (next week; Centene -> 6/23 cell)
    """
    if isinstance(d, datetime):
        d = d.date()
    wd = d.weekday()           # Mon=0 .. Sat=5, Sun=6
    if wd == 6:                # Sunday -> next week's Monday
        return d + timedelta(days=1)
    return d - timedelta(days=wd)   # Mon-Sat -> this week's Monday


# normalized_db -> {week_monday(date): latest Certified CertTimestamp}. Built in
# main() from DHT StatTimestamp so a weekly client's cert lands on the cell for
# the week its DATA covers (not the week it happened to certify). Used by
# cert_in_week for non-forward weekly clients.
CERT_WEEK_IDX = {}


def build_cert_week_index(certs):
    """normalized_db -> {week_monday: latest CertTimestamp} over Certified rows,
    grouping by the delivery week of each row's StatTimestamp (falls back to the
    CertTimestamp's own week when StatTimestamp is missing)."""
    idx = defaultdict(dict)
    for c in certs:
        if c.get("CurrentStatus") != "Certified":
            continue
        key = normalize(c["DatabaseName"])
        if not key:
            continue
        certdt = c["CertTimestamp"]
        wk = stat_week_monday(c.get("StatTimestamp") or certdt)
        cur = idx[key].get(wk)
        if cur is None or certdt > cur:
            idx[key][wk] = certdt
    return idx


def cert_on_day(client, day, cert_idx):
    """Return the latest CertTimestamp datetime for `client` on calendar `day`, else None.
    Friday cells additionally pick up Sat/Sun certs so weekend deliveries
    surface on the prior Friday's cell (e.g. AetnaRCE 5/9 Sat → Fri 5/8)."""
    days_to_check = {day}
    if day.weekday() == 4:  # Friday: also include Sat/Sun
        days_to_check.add(day + timedelta(days=1))
        days_to_check.add(day + timedelta(days=2))
    best = None
    for key in _keys_for_client(client):
        for dt, status in cert_idx.get(key, ()):
            if status != "Certified":
                continue
            if dt.date() in days_to_check:
                if best is None or dt > best:
                    best = dt
    return best


def cert_in_week(client, scheduled_day, cert_idx):
    """Latest cert attributed to this client's delivery week (the week of
    `scheduled_day`).

    Default: attribute each cert to the week its DATA covers via the
    StatTimestamp of the certified tapes (CERT_WEEK_IDX, built by
    build_cert_week_index). Per user 2026-06-24 a single cert run can certify
    two data weeks at once — e.g. Centene's 6/22 cert had StatTimestamp 6/18
    (-> the 6/16 cell) AND 6/21 (Sun -> the 6/23 cell) — and StatTimestamp
    picks the right cell for each. Falls back to the CertTimestamp week for rows
    with no StatTimestamp (handled in build_cert_week_index).

    `CERT_DIRECTION[client] = "forward"` (Premera): keeps the explicit 7-day
    forward window on CertTimestamp.
    """
    if CERT_DIRECTION.get(client) == "forward":
        cycle_start = scheduled_day
        cycle_end   = scheduled_day + timedelta(days=6)
        best = None
        for key in _keys_for_client(client):
            for dt, status in cert_idx.get(key, ()):
                if status != "Certified":
                    continue
                if cycle_start <= dt.date() <= cycle_end:
                    if best is None or dt > best:
                        best = dt
        return best

    cell_monday = scheduled_day - timedelta(days=scheduled_day.weekday())
    best = None
    for key in _keys_for_client(client):
        dt = CERT_WEEK_IDX.get(key, {}).get(cell_monday)
        if dt and (best is None or dt > best):
            best = dt
    return best


def latest_cert(client, cert_idx, on_or_before=None):
    best = None
    for key in _keys_for_client(client):
        for dt, status in cert_idx.get(key, ()):
            if status != "Certified":
                continue
            if on_or_before and dt.date() > on_or_before:
                continue
            if best is None or dt > best:
                best = dt
    return best


def _keys_for_client(client):
    # Always yield the base normalize(client) so cert/snap lookups still
    # find the natural DHT/RAMP key even when CLIENT_PRIMARY_KEY_OVERRIDE
    # has remapped the substring-target for find_matching_jobs. Per user
    # 2026-05-20: WellCare's DHT cert key is "wellcare", but the override
    # changed _keys_for_client to "wellcaremedical" only — past cert
    # dates were silently missed.
    base = normalize(client)
    seen = set()
    if base:
        seen.add(base)
        yield base
    primary = CLIENT_PRIMARY_KEY_OVERRIDE.get(client)
    if primary and primary not in seen:
        seen.add(primary)
        yield primary
    for alias in CLIENT_ALIASES.get(client, []):
        k = normalize(alias)
        if k and k not in seen:
            seen.add(k)
            yield k


# ============================================================
#                          ADO fetch
# ============================================================
def fetch_ado_tickets(min_changed_date):
    """Fetch ADO user stories tagged 'Delivery Ticket' changed since min_changed_date."""
    wiql = (
        "SELECT [System.Id] FROM WorkItems WHERE "
        "[System.TeamProject] = 'Rawlings' "
        "AND [System.Tags] CONTAINS 'Delivery Ticket' "
        f"AND [System.ChangedDate] >= '{min_changed_date.isoformat()}' "
        "AND [System.WorkItemType] = 'User Story'"
    )
    res = curl_post_json(f"{ADO_BASE}/_apis/wit/wiql?api-version=5.0", {"query": wiql})
    ids = [w["id"] for w in res.get("workItems", [])]
    if not ids:
        return []
    out = []
    fields = ",".join([
        "System.Id", "System.Title", "System.State", "System.AreaPath",
        "System.IterationPath", "System.ChangedDate", "System.CreatedDate",
        "System.AssignedTo", "System.Tags",
    ])
    for i in range(0, len(ids), 200):
        batch = ids[i:i+200]
        url = (f"{ADO_BASE}/_apis/wit/workitems?ids={','.join(map(str, batch))}"
               f"&fields={fields}&api-version=5.0")
        for w in curl_json(url).get("value", []):
            f = w["fields"]
            title = f.get("System.Title", "")
            m = TITLE_RE.match(title)
            client = m.group(2).strip() if m else ""
            kind = m.group(1) if m else ""
            assigned = f.get("System.AssignedTo", "")
            if isinstance(assigned, dict):
                assigned = assigned.get("displayName") or assigned.get("uniqueName", "")
            elif isinstance(assigned, str) and "<" in assigned:
                assigned = assigned.split("<")[0].strip()
            out.append({
                "id":       f["System.Id"],
                "title":    title,
                "state":    f.get("System.State", ""),
                "kind":     kind,
                "client":   client,
                "iter":     f.get("System.IterationPath", ""),
                "changed":  f.get("System.ChangedDate", ""),
                "created":  f.get("System.CreatedDate", ""),
                "assigned": assigned,
                "tags":     f.get("System.Tags", ""),
            })
    return out


# ============================================================
#                          RAMP fetch
# ============================================================
def fetch_ramp_jobs():
    return curl_json(f"{RAMP_BASE}/api/Ramp/Job/List").get("Data", [[]])[0]


def fetch_ramp_queue():
    """Pull RAMP queue from SQL [TRGUTIL10].RAMP.ramp.Queue directly.
    The REST endpoint /api/Ramp/Queue/List caps at 1000 items and SFTP/LogFile
    churn rotates real load entries out within hours. SQL gives full history.
    Returns dicts shaped like the REST response so downstream code is unchanged.
    """
    q = (
        "SET NOCOUNT ON; "
        "SELECT q.QueueId, q.JobId, q.Status, "
        "       CONVERT(varchar(23), q.StartDate, 121) AS StartDate, "
        "       CONVERT(varchar(23), q.EndDate, 121)   AS EndDate, "
        "       CONVERT(varchar(23), q.CreateDate, 121) AS CreateDate, "
        "       CAST(q.JobXml AS varchar(MAX)) AS JobXml "
        "FROM [RAMP].[ramp].[Queue] q "
        "WHERE q.CreateDate >= DATEADD(day, -45, GETDATE()) "
        "ORDER BY q.QueueId DESC"
    )
    SEP = "\x1f"   # ASCII unit separator — unlikely to appear in any XML/data
    r = subprocess.run(
        ["sqlcmd", "-S", SQL_SERVER, "-d", "RAMP", "-E", "-Q", q,
         "-W", "-s", SEP, "-h", "-1"],
        capture_output=True, text=True, check=False,
    )
    rows = []
    for line in r.stdout.splitlines():
        if not line or line.startswith("---") or "rows affected" in line:
            continue
        parts = line.split(SEP, 6)
        if len(parts) < 7:
            continue
        try:
            qid    = int(parts[0])
            job_id = int(parts[1])
        except ValueError:
            continue
        rows.append({
            "QueueId":    qid,
            "JobId":      job_id,
            "Status":     parts[2].strip(),
            "StartDate":  parts[3].strip() if parts[3].strip() != "NULL" else "",
            "EndDate":    parts[4].strip() if parts[4].strip() != "NULL" else "",
            "CreateDate": parts[5].strip(),
            "JobXml":     parts[6],
        })
    return rows


def fetch_ramp_snaps():
    data = curl_json(f"{RAMP_BASE}/api/Ramp/Snap/SnapQueueStatus").get("Data", [[]])
    rows = data[0] if data and isinstance(data[0], list) else data
    cutoff = datetime.now() - timedelta(days=60)
    out = []
    for s in rows:
        end = parse_dt(s.get("End"))
        if end is None or end >= cutoff:
            out.append(s)
    return out


def fetch_aetna_nmsp_loads(since):
    r"""Query SQLUtilAudit.cmse_new.SourceLog for Aetna NonMSP file loads
    started since `since`. Drives "M - Aetna NMSP - MMSEA":
      - "L"  once a file at \\trgdatacap2\MMSEA\Aetna\<year>\NonMSP has an
             ImportStartDate but no ImportCompleteDate yet (loading in CMSE),
      - "✓"  once ImportCompleteDate lands (placed on the completion date).
    Returns list of dicts {"start": datetime, "done": datetime|None}.
    """
    q = (
        "SET NOCOUNT ON; "
        "SELECT CONVERT(varchar(23), ImportStartDate, 121) AS Started, "
        "ISNULL(CONVERT(varchar(23), ImportCompleteDate, 121), '') AS Done "
        "FROM [cmse_new].[dbo].[SourceLog] WITH (NOLOCK) "
        "WHERE EntryName LIKE '%MMSEA\\Aetna\\2026\\NonMSP%' "
        f"AND ImportStartDate >= '{since.isoformat()}' "
        "ORDER BY ImportStartDate"
    )
    r = subprocess.run(
        ["sqlcmd", "-S", "SQLUtilAudit", "-d", "cmse_new", "-E", "-Q", q,
         "-W", "-s", "\t", "-h", "-1"],
        capture_output=True, text=True, check=False,
    )
    out = []
    for line in r.stdout.splitlines():
        line = line.rstrip()
        if not line or line.startswith("---") or "rows affected" in line:
            continue
        parts = line.split("\t")
        start = parse_dt(parts[0].strip()) if parts else None
        done = parse_dt(parts[1].strip()) if len(parts) > 1 and parts[1].strip() else None
        if start:
            out.append({"start": start, "done": done})
    return out


def fetch_tape_loads(db, since, server="TRGETL3", name_like=None):
    """Query <server>.<db>.etl.Tape for recent successful loads (ProcessStatus=50).
    Returns list of dicts: {FileName, FileLoadDate (datetime)}.

    `server` defaults to TRGETL3 (the PBMRx tape server); JHHC Passfile lives on
    TRGINTP3.JohnsHopkins. `name_like` optionally restricts to filenames matching
    a LIKE pattern (e.g. 'PassFile') so a shared client DB only yields the rows
    for the intended feed.
    """
    where = f"ProcessStatus = 50 AND FileLoadDate >= '{since.isoformat()}'"
    if name_like:
        where += f" AND FileName LIKE '%{name_like}%'"
    q = (
        "SET NOCOUNT ON; "
        "SELECT FileName, FileLoadDate FROM [etl].[Tape] "
        f"WHERE {where} "
        "ORDER BY FileLoadDate"
    )
    r = subprocess.run(
        ["sqlcmd", "-S", server, "-d", db, "-E", "-Q", q,
         "-W", "-s", "\t", "-h", "-1"],
        capture_output=True, text=True, check=False,
    )
    rows = []
    for line in r.stdout.splitlines():
        if not line or line.startswith("---") or "rows affected" in line:
            continue
        parts = line.split("\t")
        if len(parts) < 2:
            continue
        dt = parse_dt(parts[1])
        if dt:
            rows.append({"FileName": parts[0].strip(), "FileLoadDate": dt})
    return rows


# Map of client → (database name, snap-index source key).
# Server defaults to TRGETL3 (see TAPE_LOAD_SERVER for overrides).
TAPE_LOAD_SOURCES = {
    "OptumPBMRx":      ("OptumPBMRx",      "optumpbmrx"),
    "ESIPBMRx":        ("ESIPBMRx",        "esipbmrx"),
    "MedImpactPBMRx":  ("MedImpactPBMRx",  "medimpactpbmrx"),
    # JHHC Passfile loads to TRGINTP3.JohnsHopkins.etl.Tape (NOT TRGETL3, and the
    # JohnsHopkins DB also holds the main JHHC Medical feed — so filter to
    # PassFile filenames). Added 2026-06-16 per user: the 6/12 passfile reload
    # wasn't captured by the fragile 'JHHC Passfile Email' RAMP job; the tape is
    # authoritative. Placement lands ✓ on the latest load date (single cell).
    "JHHCPassfile":    ("JohnsHopkins",    "jhhcpassfile"),
}

# Per-client SQL server override for TAPE_LOAD_SOURCES (default TRGETL3).
TAPE_LOAD_SERVER = {
    "JHHCPassfile": "TRGINTP3",
}

# Per-client FileName LIKE filter for TAPE_LOAD_SOURCES — restricts a shared
# client DB to just the intended feed's rows.
TAPE_LOAD_NAME_FILTER = {
    "JHHCPassfile": "PassFile",
}

# Regex for state codes inside ESIPBMRx tape filenames (e.g. Rawlings_FL_, Rawlings_GA_)
ESIPBMRX_STATE_RE = re.compile(r"Rawlings_([A-Z]{2})_", re.I)

# Multi-week load detection: client → (TRGETL3 db, regex extracting week range
# from filename). The capture groups should be (start_yyyymmdd, end_yyyymmdd)
# or a single token uniquely identifying a week's worth of data. When recent
# loads contain >1 distinct week-key, the client label gets "(N weeks)".
MULTI_WEEK_CLIENTS = {}


def find_matching_jobs(client_id, jobs):
    primary = CLIENT_PRIMARY_KEY_OVERRIDE.get(client_id) or normalize(client_id)
    if not primary:
        return []
    targets = [primary] + [normalize(a) for a in CLIENT_ALIASES.get(client_id, [])]
    targets = [t for t in targets if t]
    matches = []
    for j in jobs:
        jn = normalize(j.get("JobName", "") or "")
        # Also test a digit-code-collapsed version of jn so JobNames with a
        # numeric step code between feed and sub-feed (e.g.
        # "CareFirst 0110 Facets Load" → "carefirst0110facetsload") still match
        # aliases like "carefirstfacets" via substring.
        jn_collapsed = re.sub(r"\d+", "", jn)
        fn = normalize((j.get("Feed")   or {}).get("FeedName", "") or "")
        cn = normalize((j.get("Client") or {}).get("ClientName", "") or "")
        if any(t in jn or t in jn_collapsed or t == fn or t == cn for t in targets):
            matches.append(j)
    real = [j for j in matches if not re.search(r"(logfile|sftp)", j.get("JobName", ""), re.I)]
    return real or matches


def build_snap_index(jobs, queue, snaps, tape_loads=None):
    """Return: date -> list of (normalized_source, datetime, destination, kind, job_name).

    `kind` ∈ {"snap", "load", "tape"} — distinguishes data source so a
    snap-only client doesn't get ✓ from a load-step completion.
    `job_name` — original RAMP JobName (queue entries only; empty for tape /
    RAMP snap-endpoint entries). Used by LOAD_NAME_REQUIRED to filter
    lookups so ancillary load jobs don't trigger L for cert clients.
    """
    by_date = defaultdict(list)

    if tape_loads:
        for src_key, rows in tape_loads.items():
            for row in rows:
                dt = row.get("FileLoadDate")
                if not dt or dt.year < 2026:
                    continue
                by_date[dt.date()].append((src_key, dt, "", "tape", ""))

    for s in snaps:
        # Accept: Success, Success/ManualFix, Success/NoWork — and Resolved
        # (user: "wait for the Ramp card to be marked Resolved").
        st = str(s.get("Status", ""))
        if not (st.startswith("Success") or st == "Resolved"):
            continue
        if s.get("TaskName") not in ("DeliverFlow", "DeliverSingle", "DeliverFlowSet"):
            continue
        end = parse_dt(s.get("End"))
        if not end or end.year < 2026:
            continue
        src = s.get("Source", "") or ""
        dest = s.get("Destination", "") or ""
        src_norm = normalize(re.sub(r"_(mine|snap|stage|load|rta)$", "", src, flags=re.I))
        if src_norm:
            by_date[end.date()].append((src_norm, end, dest, "snap", ""))

    job_by_id = {j.get("JobId"): j for j in jobs}
    for q in queue:
        if q.get("Status") not in ("Successful", "Resolved"):
            continue
        j = job_by_id.get(q.get("JobId"))
        if not j:
            continue
        jn = j.get("JobName", "") or ""
        jn_lower = jn.lower()
        # Index Load, Stage AND Snap jobs (the snap step is what triggers ✓
        # for SNAP_KIND_ONLY clients like ESIPBMRx, PrimePBMRx, Kaiser_GE).
        # EXTRA_INDEXED_JOBS lets through delivery-signal jobs that lack those
        # keywords (e.g. 'JHHC Passfile Email').
        if (not any(kw in jn_lower for kw in ("load", "stage", "snap", "mine"))
                and jn_lower.strip() not in EXTRA_INDEXED_JOBS):
            continue
        # skip log/sftp noise
        if re.search(r"(logfile|sftp|upload)", jn, re.I):
            continue
        # Index by START date — a load that begins 5/13 and finishes 5/14 is
        # for 5/13's data (per user: "AetnaHRP for 5/13 finished on 5/14, so
        # the 5/13 date should get a checkmark"). Keep EndDate as the
        # sort-time for latest-match purposes.
        start_dt = parse_dt(q.get("StartDate"))
        end_dt   = parse_dt(q.get("EndDate"))
        attribution_date = (start_dt.date() if start_dt
                            else (end_dt.date() if end_dt else None))
        if attribution_date is None:
            continue
        end = end_dt or start_dt
        # Index this completion under a JobName-derived key only. The Feed
        # name is unreliable because related-but-distinct clients can share
        # one Feed (e.g. "Kaiser WA 0110 Load" and "Kaiser WARX 0110 Load"
        # both have Feed "KaiserWA" → "kaiserwa", causing KaiserWARx loads
        # to falsely match Kaiser_WA). The JobName prefix uniquely identifies
        # which client the completion belongs to.
        kind = "snap" if re.search(r"\b(snap|mine)\b", jn, re.I) else "load"
        # Take everything before the first digit sequence in the JobName,
        # then strip trailing step words.
        m = re.match(r"^([^\d]+)", jn)
        prefix = m.group(1).strip() if m else jn
        prefix = re.sub(r"\s*(load|stage|snap|etl|daily|mine|start)\s*$", "",
                        prefix, flags=re.I).strip()
        k = normalize(prefix)
        if k and len(k) >= 4:
            by_date[attribution_date].append((k, end, "", kind, jn))
        # Also emit a "feed+sub-feed" key for JobNames matching
        # "<feed> <digits> <sub-feed> <step>" — e.g. "CareFirst 0110 Facets
        # Load" → "carefirstfacets". The base key (k) above captures only the
        # part before the digit code and loses sub-feed identity.
        sub_m = re.match(
            r"^([^\d]+?)\s+\d+\s+([^\d]+?)\s+"
            r"(post\s+snap|load|stage|snap|etl|daily|mine|start)\s*$",
            jn, re.I,
        )
        if sub_m:
            combined = normalize(sub_m.group(1) + sub_m.group(2))
            if combined and combined != k and len(combined) >= 4:
                by_date[attribution_date].append((combined, end, "", kind, jn))
    return by_date


def _tape_keys_only(client):
    """For tape-only clients, return the exact tape source key (no aliases)."""
    src_keys = {sk for (db, sk) in TAPE_LOAD_SOURCES.values() if client in TAPE_LOAD_SOURCES and TAPE_LOAD_SOURCES[client][1] == sk}
    return list(src_keys) if src_keys else []


def _src_matches_client(src_norm, client_keys, min_len=4):
    """Strict equality match between snap source and client keys.

    Substring matching produced too many false positives (e.g. snap source
    "aetnarx" was being attributed to clients like "AetnaRx_LegacyDMG", and
    "bcbsfl" to "BCBSFLEligibilityLoad"). Use explicit CLIENT_ALIASES to
    register variant source names per client.
    """
    if not src_norm or len(src_norm) < min_len:
        return False
    for k in client_keys:
        if k and k == src_norm:
            return True
    return False


def _kind_allowed(client, kind):
    """Return True if a snap_idx entry of this `kind` is allowed for `client`."""
    if client in LOAD_AS_DELIVERY_CLIENTS:
        return True   # accept anything (load/tape/snap)
    if client in SNAP_KIND_ONLY_CLIENTS:
        return kind == "snap"   # only real RAMP snap completions
    return True       # daily clients accept any kind


def _load_name_allowed(client, entry_jn, kind):
    """For clients in LOAD_NAME_REQUIRED, require the snap_idx entry's
    underlying JobName to contain one of the listed keywords. Entries
    without a JobName (RAMP snap-endpoint deliveries, tape rows) are
    rejected — they can't be verified against the whitelist.

    Exception: SNAP_ONLY clients use the snap-endpoint completion AS the
    delivery signal (no JobName available). For them, snap-kind entries
    bypass the filter. Per user 2026-05-18: MMOH was snapped 5/2 and
    should show ✓ via the RAMP snap-endpoint entry.

    Clients NOT in LOAD_NAME_REQUIRED bypass the filter entirely."""
    required = LOAD_NAME_REQUIRED.get(client)
    if not required:
        return True
    if kind == "snap" and client in SNAP_ONLY_CLIENTS:
        return True
    if not entry_jn:
        return False
    jn_lower = entry_jn.lower()
    return any(kw in jn_lower for kw in required)


def snap_on_day(client, day, snap_idx, window_days=0, forward_days=0):
    """Return latest snap/load completion datetime for `client` on `day`
    (or within `window_days` calendar days BEFORE `day`, or `forward_days`
    days AFTER for late completions).
    Applies SNAP_DESTINATION_FILTER and per-client kind filter.
    """
    if client in TAPE_ONLY_CLIENTS:
        if client not in TAPE_LOAD_SOURCES:
            return None
        wanted_src = TAPE_LOAD_SOURCES[client][1]
        # For SNAP_KIND_ONLY_CLIENTS like ESIPBMRx/MedImpactPBMRx the tape
        # entries represent LOADS — only count them as ✓ if the user actually
        # treats the load as delivery (LOAD_AS_DELIVERY_CLIENTS).
        if client in SNAP_KIND_ONLY_CLIENTS and client not in LOAD_AS_DELIVERY_CLIENTS:
            wanted_src = None
        best = None
        for offset in range(-forward_days, window_days + 1):
            check_day = day - timedelta(days=offset)
            for entry in snap_idx.get(check_day, ()):
                if wanted_src and entry[0] == wanted_src:
                    dt = entry[1]
                    if best is None or dt > best:
                        best = dt
        if best:
            return best
        # fall through to regular substring matching for SNAP_KIND_ONLY tape clients

    keys = [k for k in _keys_for_client(client) if k]
    if not keys:
        return None
    dest_filter = SNAP_DESTINATION_FILTER.get(client, "").lower()
    best = None
    for offset in range(-forward_days, window_days + 1):
        check_day = day - timedelta(days=offset)
        for entry in snap_idx.get(check_day, ()):
            src_norm, dt = entry[0], entry[1]
            dest = entry[2] if len(entry) > 2 else ""
            kind = entry[3] if len(entry) > 3 else "snap"
            jn   = entry[4] if len(entry) > 4 else ""
            if dest_filter and dest_filter not in dest.lower():
                continue
            if not _kind_allowed(client, kind):
                continue
            if not _load_name_allowed(client, jn, kind):
                continue
            if _src_matches_client(src_norm, keys):
                if best is None or dt > best:
                    best = dt
    return best


def snap_in_week(client, scheduled_day, snap_idx):
    """Return latest snap/load completion in the Mon-Fri week containing
    scheduled_day. A snap or load done anywhere this week (even before the
    scheduled day) counts as "completed this week" — e.g. Wellmark Tue cell
    picks up its Mon load, WellCare Fri cell picks up its Tue load.
    Monday-scheduled clients also pick up the prior weekend (Sat-Sun).
    """
    week_start = scheduled_day - timedelta(days=scheduled_day.weekday())
    week_end   = week_start + timedelta(days=4)
    # Always look back 2 days into the prior Sat+Sun so weekend loads
    # surface on any weekday cell, not just Monday-scheduled clients.
    # Per user 2026-05-18: Wellmark 0210 Claims Load on Sun 5/17 should
    # appear as L on the Tue 5/19 cell.
    window_start = week_start - timedelta(days=2)
    best = None
    d = window_start
    while d <= week_end:
        ts = snap_on_day(client, d, snap_idx)
        if ts and (best is None or ts > best):
            best = ts
        d += timedelta(days=1)
    return best


def latest_snap_this_month(client, snap_idx, year, month, on_or_before):
    """Latest snap completion datetime for `client` in (year,month), on/before today."""
    if client in TAPE_ONLY_CLIENTS:
        if client not in TAPE_LOAD_SOURCES:
            return None
        wanted_src = TAPE_LOAD_SOURCES[client][1]
        # For SNAP_KIND_ONLY tape clients (ESIPBMRx, MedImpactPBMRx) the tape
        # entries represent LOAD activity, not snap delivery — fall through to
        # regular kind-filtered matching so only actual snap completions count.
        if client in SNAP_KIND_ONLY_CLIENTS and client not in LOAD_AS_DELIVERY_CLIENTS:
            wanted_src = None
        if wanted_src:
            best = None
            for d, entries in snap_idx.items():
                if d.year != year or d.month != month or d > on_or_before:
                    continue
                for entry in entries:
                    if entry[0] == wanted_src:
                        dt = entry[1]
                        if best is None or dt > best:
                            best = dt
            if best:
                return best
        # fall through to regular matching for SNAP_KIND_ONLY tape clients

    keys = [k for k in _keys_for_client(client) if k]
    if not keys:
        return None
    dest_filter = SNAP_DESTINATION_FILTER.get(client, "").lower()
    best = None
    for d, entries in snap_idx.items():
        if d.year != year or d.month != month or d > on_or_before:
            continue
        for entry in entries:
            src_norm, dt = entry[0], entry[1]
            dest = entry[2] if len(entry) > 2 else ""
            kind = entry[3] if len(entry) > 3 else "snap"
            jn   = entry[4] if len(entry) > 4 else ""
            if dest_filter and dest_filter not in dest.lower():
                continue
            if not _kind_allowed(client, kind):
                continue
            if not _load_name_allowed(client, jn, kind):
                continue
            if _src_matches_client(src_norm, keys):
                if best is None or dt > best:
                    best = dt
    return best


def load_this_month(client, snap_idx, year, month, on_or_before):
    """Return latest LOAD-kind completion datetime for `client` in (year,month),
    on/before today. Considers kind ∈ {"load","tape"} only — not snap completions.
    Used so monthly cert-only clients don't show L from a stray
    /Ramp/Snap completion when the load job itself hasn't run yet.
    """
    if client in TAPE_ONLY_CLIENTS:
        if client not in TAPE_LOAD_SOURCES:
            return None
        wanted_src = TAPE_LOAD_SOURCES[client][1]
        best = None
        for d, entries in snap_idx.items():
            if d.year != year or d.month != month or d > on_or_before:
                continue
            for entry in entries:
                if entry[0] == wanted_src and (len(entry) <= 3 or entry[3] in ("load", "tape")):
                    dt = entry[1]
                    if best is None or dt > best:
                        best = dt
        return best

    keys = [k for k in _keys_for_client(client) if k]
    if not keys:
        return None
    best = None
    for d, entries in snap_idx.items():
        if d.year != year or d.month != month or d > on_or_before:
            continue
        for entry in entries:
            src_norm, dt = entry[0], entry[1]
            kind = entry[3] if len(entry) > 3 else "snap"
            jn   = entry[4] if len(entry) > 4 else ""
            if kind not in ("load", "tape"):
                continue
            if not _load_name_allowed(client, jn, kind):
                continue
            if _src_matches_client(src_norm, keys):
                if best is None or dt > best:
                    best = dt
    return best


def is_loading_today(client, queue, jobs):
    """True if a matching enabled job is currently Ready/Running.

    The set of "L"-triggering job types depends on client class:
      - LOAD_AS_DELIVERY clients (AetnaRx/AetnaHRP/etc.): only LOAD jobs.
        Once their load step finishes, ✓ takes over.
      - All other clients (CenteneRx/WellCareRx/etc.): LOAD or SNAP jobs.
        These need a cert to complete the cycle, so they stay L through
        both the load and snap steps until the cert lands.

    Stage / logfile / sftp / upload jobs never count as L.
    """
    matched = find_matching_jobs(client, jobs)
    load_only = client in LOAD_AS_DELIVERY_CLIENTS or client in L_ON_LOAD_ONLY_CLIENTS
    required_kwds = LOAD_NAME_REQUIRED.get(client)
    job_ids = set()
    for j in matched:
        if j.get("Enabled") != 1:
            continue
        jn = (j.get("JobName") or "").lower()
        if any(kw in jn for kw in ("stage", "logfile", "sftp", "upload")):
            continue
        is_load = "load" in jn and "snap" not in jn and "mine" not in jn
        is_snap = ("snap" in jn or "mine" in jn) and "load" not in jn
        # Per-client JobName whitelist (e.g. Rx clients where only MasterLoad
        # or Claims Load jobs should signal "loading"; COBC/IHP/ABII do not).
        if required_kwds and not any(kw in jn for kw in required_kwds):
            continue
        if load_only:
            if is_load:
                job_ids.add(j.get("JobId"))
        else:
            if is_load or is_snap:
                job_ids.add(j.get("JobId"))
    for q in queue:
        if q.get("JobId") in job_ids and q.get("Status") in ("Ready", "Running"):
            return True
    return False


def scan_adhoc_loads(queue, jobs, today, since, weekend_shift=True):
    """Scan RAMP queue for ad-hoc Load jobs that have no fixed schedule on
    the report. Each completed run shows up as a one-off row in the weekly
    section on the day it landed; in-flight runs surface today with 'L'.

    Per user 2026-06-04: MSPI Load jobs and HumanaRx Load are ad-hoc — they
    only appear when they actually run. Drop the row when nothing happened.

    Completions are attributed by **Start** date (matching build_snap_index) so
    a load that begins late and finishes after midnight still counts for the day
    it ran. `weekend_shift` (default True) surfaces Sat→Fri / Sun→Mon for ad-hoc
    rows; pass False when the caller wants the true (unshifted) load weekday.

    Returns a list of dicts: {"label", "day", "marker", "alert"}.
    """
    mspi_re   = re.compile(r"\bMSPI\b.*\bLoad\b", re.IGNORECASE)
    # HumanaRx returned to Monthly classification 2026-06-04 — kept the
    # function plumbing in case the user wants to add more ad-hoc patterns
    # later; for now only MSPI surfaces here.
    # Queue rows only carry JobId, not JobName — cross-reference via jobs.
    job_by_id = {j.get("JobId"): (j.get("JobName") or "") for j in jobs}
    # Walk newest → oldest (queue is QueueId DESC). For each (label, day, marker)
    # we keep the first occurrence; we also drop any Load Failure entry for a
    # label once a strictly newer non-failure entry (Ready/Running/Success/
    # Resolved) has been seen for the same label. Per user 2026-06-04: "if a
    # new MSPI load starts for the same client that has a failure, the failure
    # can be dropped."
    seen = set()
    rows = []
    nonfailure_seen = set()  # labels whose newest run is NOT a failure
    for q in queue:
        name = (job_by_id.get(q.get("JobId")) or "").strip()
        if not name:
            continue
        if not mspi_re.search(name):
            continue
        # Skip the snap-step variant; we only key off Load completion.
        if re.search(r"\b(snap|mine|logfile|sftp|stage)\b", name, re.I):
            continue
        status = (q.get("Status") or "").strip()
        start = parse_dt(q.get("StartDate"))
        end   = parse_dt(q.get("EndDate"))
        if not start or start.date() < since:
            continue
        base = re.sub(r"\s*\d+.*$", "", name).strip()
        base = re.sub(r"\s*MSPI\s*$", "", base, flags=re.IGNORECASE).strip()
        label = f"{base} MSPI"
        s_lower = status.lower()
        if s_lower.startswith("success") or s_lower == "resolved":
            day = start.date()
            marker, alert = "✓", False
            nonfailure_seen.add(label)
        elif s_lower in ("ready", "running"):
            day, marker, alert = today, "L", False
            nonfailure_seen.add(label)
        elif s_lower == "failed":
            # Drop the failure if a strictly newer (already-seen) non-failure
            # run exists for the same label.
            if label in nonfailure_seen:
                continue
            day = start.date()
            marker, alert = "Load Failure", True
        else:
            continue
        if weekend_shift:
            if day.weekday() == 5:
                day -= timedelta(days=1)
            elif day.weekday() == 6:
                day += timedelta(days=1)
        key = (label, day, marker)
        if key in seen:
            continue
        seen.add(key)
        rows.append({"label": label, "day": day, "marker": marker, "alert": alert})
    return rows


# The two jobs whose joint completion = a finished "Kaiser Submission" (per
# user 2026-06-09). Exact lowercased names — must NOT match the separate
# 'Kaiser Pareo Audit Submission Upload'. These are Logfile/Upload jobs, which
# the normal index/L logic excludes, so they're tracked here explicitly.
KAISER_SUBMISSION_JOBS = (
    "kaiser pareo submission logfile",
    "kaiser pareo submission upload",
)


def scan_kaiser_submission(queue, jobs, today, since):
    """Per-day status for the daily 'Kaiser Submission' row.

    ✓ on a day when BOTH 'Kaiser Pareo Submission Logfile' AND
    'Kaiser Pareo Submission Upload' completed (Successful/Resolved) that day;
    'L' on today while either is still Ready/Running. Weekend completions snap
    to the closest weekday (Sat→Fri, Sun→Mon).

    Returns (done_days: set[date], running_today: bool).
    """
    job_by_id = {j.get("JobId"): (j.get("JobName") or "").strip().lower() for j in jobs}

    def snap_wd(d):
        if d.weekday() == 5:
            return d - timedelta(days=1)
        if d.weekday() == 6:
            return d + timedelta(days=1)
        return d

    done = {name: set() for name in KAISER_SUBMISSION_JOBS}
    running_today = False
    for q in queue:
        name = job_by_id.get(q.get("JobId"))
        if name not in KAISER_SUBMISSION_JOBS:
            continue
        status = (q.get("Status") or "").strip().lower()
        start = parse_dt(q.get("StartDate"))
        end   = parse_dt(q.get("EndDate"))
        dt = end or start
        if status.startswith("success") or status == "resolved":
            if dt and dt.date() >= since:
                done[name].add(dt.date())
        elif status in ("ready", "running"):
            running_today = True
    log_days = done["kaiser pareo submission logfile"]
    up_days  = done["kaiser pareo submission upload"]
    done_days = {snap_wd(d) for d in (log_days & up_days)}
    return done_days, running_today


def find_unconfigured_masterload_clients(jobs):
    """Scan RAMP jobs for `<Client> MasterLoad 0110 Load` entries whose
    derived client name isn't recognised by any existing config dict.

    Returns list of dicts:
      {"raw": str, "normalized": str, "pbmrx": bool, "enabled": bool}

    Per user 2026-06-03: keep an eye on RAMP for these and add new
    implementations to the report. PBMRx hits get SNAP_KIND_ONLY behavior.
    """
    out = []
    seen = set()
    known = (set(DAILY_CLIENTS) | set(WEEKLY_CLIENTS.keys())
             | set(MONTHLY_CLIENTS) | {KAISER_PREPAY_CLIENT})
    known_keys = set()
    for c in known:
        for k in _keys_for_client(c):
            if k:
                known_keys.add(k)
    pat = re.compile(r"^(.+?)\s*MasterLoad\s+0110\s+Load\s*$", re.IGNORECASE)
    for j in jobs:
        name = (j.get("JobName") or "").strip()
        if not name:
            continue
        m = pat.match(name)
        if not m:
            continue
        client_raw = m.group(1).strip()
        norm = normalize(client_raw)
        if not norm or norm in seen:
            continue
        # Match against known client keys with substring (matches aliases too).
        matched = any(norm == k or norm in k or k in norm for k in known_keys)
        if matched:
            continue
        seen.add(norm)
        out.append({
            "raw": client_raw,
            "normalized": norm,
            "pbmrx": "PBMRx" in name,
            "enabled": j.get("Enabled") == 1,
        })
    return out


# Ancillary sub-pipeline modifiers that should NOT count toward primary-load
# inactivity (per user 2026-06-03: the inactive-label rule targets the main
# 0100/0110 cycle only, not COBC/RTA/etc. sub-jobs).
_ANCILLARY_JOB_TOKEN_RE = re.compile(
    r"\b(cobc|rta|abii|ihp|cms\s*referral|mmsea|covid|monthly\s+claim|adjustment)\b",
    re.IGNORECASE,
)


def auto_inactive_from_ramp(jobs):
    """Set of clients (canonical names from our config) whose primary
    0100 Stage / 0110 Load jobs in RAMP are ALL Inactive (Enabled=0).
    Excludes any client whose name starts with "Kaiser" (per user
    2026-06-03 — Kaiser feeds have their own snap/inactive semantics).
    Excludes ancillary job modifiers via _ANCILLARY_JOB_TOKEN_RE.
    """
    primary_state = defaultdict(lambda: {"active": 0, "inactive": 0})
    for j in jobs:
        name = (j.get("JobName") or "").strip()
        if not name:
            continue
        if "0100 Stage" not in name and "0110 Load" not in name:
            continue
        if _ANCILLARY_JOB_TOKEN_RE.search(name):
            continue
        jn = normalize(name)
        is_active = (j.get("Enabled") == 1)
        for client in (set(DAILY_CLIENTS) | set(WEEKLY_CLIENTS.keys())
                       | set(MONTHLY_CLIENTS) | {KAISER_PREPAY_CLIENT}):
            if client.startswith("Kaiser"):
                continue
            for k in _keys_for_client(client):
                if k and (k in jn or jn.startswith(k)):
                    if is_active:
                        primary_state[client]["active"] += 1
                    else:
                        primary_state[client]["inactive"] += 1
                    break
    out = set()
    for client, state in primary_state.items():
        if state["inactive"] > 0 and state["active"] == 0:
            out.add(client)
    return out


def has_inactive_jobs(client, jobs, cert_idx, snap_idx, today):
    """True if RAMP has no enabled jobs for the client AND there has been
    no DHT certification or snap completion in the last 30 days.
    Clients in FORCED_INACTIVE are always inactive.
    """
    if client in FORCED_INACTIVE:
        return True
    cutoff = today - timedelta(days=30)
    for k in _keys_for_client(client):
        for dt, _ in cert_idx.get(k, ()):
            if dt.date() >= cutoff:
                return False
    for d in list(snap_idx.keys()):
        if d < cutoff:
            continue
        for entry in snap_idx[d]:
            src_norm = entry[0]
            for k in _keys_for_client(client):
                if k and (k in src_norm or src_norm in k):
                    return False
    matched = find_matching_jobs(client, jobs)
    if not matched:
        return False
    enabled = [j for j in matched if j.get("Enabled") == 1]
    return len(enabled) == 0


def has_recent_failure(client, queue, jobs, today):
    """True if a LOAD job (not stage/snap/logfile) failed in the last 3 days
    with no Successful or Resolved run after it. Stage failures and
    intermediate-step failures don't trigger "Load Failure" — per user,
    only true load-step failures count.

    Clients in LOAD_NAME_REQUIRED also restrict to those keyword patterns,
    so an ancillary "RTA Load" / "COBC Load" failure doesn't trigger
    Load Failure for the main client cycle. Per user 2026-05-19.
    """
    matched = find_matching_jobs(client, jobs)
    required_kwds = LOAD_NAME_REQUIRED.get(client)
    job_ids = set()
    for j in matched:
        jn = (j.get("JobName") or "").lower()
        # Only LOAD-named jobs count for failure detection
        if "load" not in jn:
            continue
        # COBC and ABII are ancillary file pipelines — a COBC/ABII load failure
        # must NOT mark the client's claims delivery as a failure (per user
        # 2026-06-18: CareFirstRx's "Load Failure" was the COBC load, not the
        # 'CareFirstRx 0110 Load' claims job). Excluded globally for all clients.
        if any(kw in jn for kw in ("stage", "snap", "mine", "logfile",
                                   "sftp", "upload", "cobc", "abii")):
            continue
        if required_kwds and not any(kw in jn for kw in required_kwds):
            continue
        job_ids.add(j.get("JobId"))
    if not job_ids:
        return False
    cutoff = today - timedelta(days=3)
    latest_failed = {}
    latest_success = {}
    for q in queue:
        jid = q.get("JobId")
        if jid not in job_ids:
            continue
        end = parse_dt(q.get("EndDate") or q.get("StartDate"))
        if not end or end.date() < cutoff:
            continue
        if q.get("Status") == "Failed":
            if jid not in latest_failed or end > latest_failed[jid]:
                latest_failed[jid] = end
        elif q.get("Status") in ("Successful", "Resolved"):
            # Resolved counts as success (per user: wait for the queue card
            # to be marked Resolved as a recovery from a prior failure).
            if jid not in latest_success or end > latest_success[jid]:
                latest_success[jid] = end
    for jid, fail_dt in latest_failed.items():
        succ_dt = latest_success.get(jid)
        if succ_dt is None or fail_dt > succ_dt:
            return True
    return False


# ============================================================
#                       ADO matching per client
# ============================================================
def build_ticket_index(tickets, jobs):
    """client_name -> latest ticket dict + a list of all this-month tickets (for monthly placement)."""
    latest = {}
    placements = defaultdict(list)
    for t in tickets:
        cid = t["client"]
        if not cid:
            continue
        # for monthly placement, use ticket Created date (or snap completion for Load and Snap)
        created = parse_dt(t["created"])
        placed = created.date() if created else None
        if placed:
            placements[cid].append(placed)

        existing = latest.get(cid)
        if existing is None or (parse_dt(t["changed"]) or datetime.min) > (parse_dt(existing["changed"]) or datetime.min):
            latest[cid] = t
    return latest, placements


# ============================================================
#                       calendar planning
# ============================================================
def display_name(client, monthly=False, extra_suffix=""):
    # `monthly` no longer prepends an "M - " tag — removed per user 2026-06-16.
    base = CLIENT_DISPLAY_NAME.get(client, client)
    suffix = CLIENT_SUFFIXES.get(client, "")
    return f"{base}{suffix}{extra_suffix}"


def esipbmrx_states_for_week(week_start, week_end, esipbmrx_tape_rows):
    """Return sorted list of state codes loaded between week_start and week_end."""
    states = set()
    for row in esipbmrx_tape_rows or ():
        dt = row.get("FileLoadDate")
        if not dt or not (week_start <= dt.date() <= week_end):
            continue
        m = ESIPBMRX_STATE_RE.search(row.get("FileName", "") or "")
        if m:
            states.add(m.group(1).upper())
    return sorted(states)


def count_multi_week_loads(client, week_start, week_end, multi_week_loads):
    """Return number of distinct week-key tuples captured in filenames loaded
    in the [week_start, week_end] window for this client."""
    info = MULTI_WEEK_CLIENTS.get(client)
    if not info:
        return 0
    _, pattern = info
    rows = multi_week_loads.get(client, [])
    keys = set()
    for row in rows:
        dt = row.get("FileLoadDate")
        if not dt or not (week_start <= dt.date() <= week_end):
            continue
        m = pattern.search(row.get("FileName", "") or "")
        if m:
            keys.add(m.groups())
    return len(keys)


def month_weeks(year, month):
    """Return Mon-Fri weeks rendered on this month's tab.

    A week intersecting two months is claimed by whichever month has 3+
    weekdays in it (per user 2026-06-01: "whichever month has 3 or more
    weekdays, that week should be joined with the remaining days of the
    other month"). The claiming month keeps the full 5-day week, including
    carryover dates from the adjacent month. The losing month drops that
    partial week — it will appear on the claiming month's tab instead.

    Example: 6/29 Mon and 6/30 Tue (only 2 June weekdays) drop off June and
    join July's Week 1, which has 7/1, 7/2, 7/3 (3 July weekdays).
    """
    cal = calendar.Calendar(firstweekday=0)
    out = []
    for wk in cal.monthdatescalendar(year, month):
        weekdays = wk[:5]  # Mon-Fri only
        in_month = sum(1 for d in weekdays if d.month == month)
        if in_month >= 3:
            out.append(list(weekdays))
    return out


def next_monday_if_weekend(d):
    """If d is Sat/Sun, return the following Monday; otherwise d."""
    if d.weekday() == 5:
        return d + timedelta(days=2)
    if d.weekday() == 6:
        return d + timedelta(days=1)
    return d


def closest_weekday(d):
    """Snap d to the nearest Mon-Fri. Sat → previous Fri, Sun → next Mon."""
    if d.weekday() == 5:
        return d - timedelta(days=1)
    if d.weekday() == 6:
        return d + timedelta(days=1)
    return d


def us_federal_holidays(year):
    """Return {date: name} for the 11 US federal holidays in `year`.
    Fixed-date holidays that fall on a Saturday are observed the prior
    Friday; on Sunday, the following Monday."""
    out = {}

    def nth_weekday(month, weekday, n):
        cal = calendar.monthcalendar(year, month)
        days = [w[weekday] for w in cal if w[weekday] != 0]
        return date(year, month, days[n - 1])

    def last_weekday(month, weekday):
        cal = calendar.monthcalendar(year, month)
        days = [w[weekday] for w in cal if w[weekday] != 0]
        return date(year, month, days[-1])

    def observed(d):
        if d.weekday() == 5:
            return d - timedelta(days=1)
        if d.weekday() == 6:
            return d + timedelta(days=1)
        return d

    fixed = [
        (date(year,  1,  1), "New Year's Day"),
        (date(year,  6, 19), "Juneteenth"),
        (date(year,  7,  4), "Independence Day"),
        (date(year, 11, 11), "Veterans Day"),
        (date(year, 12, 25), "Christmas Day"),
    ]
    for d, name in fixed:
        d_obs = observed(d)
        out[d_obs] = name + (" (observed)" if d_obs != d else "")

    out[nth_weekday(1,  0, 3)]  = "MLK Day"               # 3rd Mon Jan
    out[nth_weekday(2,  0, 3)]  = "Presidents' Day"       # 3rd Mon Feb
    out[last_weekday(5, 0)]     = "Memorial Day"          # last Mon May
    out[nth_weekday(9,  0, 1)]  = "Labor Day"             # 1st Mon Sep
    out[nth_weekday(10, 0, 2)]  = "Columbus Day"          # 2nd Mon Oct
    out[nth_weekday(11, 3, 4)]  = "Thanksgiving"          # 4th Thu Nov
    return out


def nmsp_mmsea_date(year, month):
    return next_monday_if_weekend(date(year, month, 15))


def average_cert_day(client, cert_idx):
    """Return the most common day-of-month this client certified on (across history)."""
    days = []
    for key in _keys_for_client(client):
        for dt, status in cert_idx.get(key, ()):
            if status == "Certified":
                days.append(dt.day)
    if not days:
        return None
    return int(round(sum(days) / len(days)))


def is_friday_or_later_in_week(today, scheduled_day):
    """True if today is on/past Friday of the same week as scheduled_day."""
    week_start = scheduled_day - timedelta(days=scheduled_day.weekday())
    week_friday = week_start + timedelta(days=4)
    return today >= week_friday


def plan_calendar(year, month, cert_idx, snap_idx, latest_tickets, monthly_placements,
                  ramp_jobs, ramp_queue, esipbmrx_tape=None, multi_week_loads=None,
                  aetna_nmsp_loads=None, optumpbmrx_tape=None):
    """Return ((sections, weeks)) layout.

    sections: dict {kind: dict {date: [(label, marker, alert)] }}
      kind is one of 'daily', 'weekly', 'monthly', 'kaiser'.
    weeks: list of weeks (each = list of 5 dates Mon-Fri or None).
    """
    today = date.today()
    weeks = month_weeks(year, month)
    all_days = [d for wk in weeks for d in wk if d is not None]

    daily   = defaultdict(list)
    weekly  = defaultdict(list)
    monthly = defaultdict(list)
    kaiser  = defaultdict(list)

    def is_kaiser_feed(c):
        return c.startswith("Kaiser") and c != KAISER_PREPAY_CLIENT

    def alert_state(client, day, marker):
        """Pink-fill the Date cell when client is in a problem state.
        - Load Failure / Inactive / Failed → always shade
        - No Data → only shade when today is >7 days past the expected END day
        - Kaiser feeds → shade only when past Friday without cert that week
        """
        # Problem-state markers ALWAYS shade pink, regardless of client class.
        # Per user 2026-05-19: "Kaiser_HI does have a load failure, but is not
        # in pink." The Kaiser-feed branch below previously short-circuited
        # before this check, hiding the Load Failure shade for Kaiser feeds.
        # "Snap" added for snap-disabled clients (e.g. Kaiser_AmbM).
        if marker in ("Load Failure", "Inactive", "Failed", "Deployment", "Snap"):
            return True
        # Implementation phase — never pink (new client; lack of jobs/data
        # would otherwise trip has_inactive_jobs).
        if marker == "Implementation":
            return False
        if is_kaiser_feed(client):
            # Snap-only / load-as-delivery / forced-inactive Kaiser feeds
            # have their own per-client semantics — fall through to the
            # generic marker rules so a valid ✓ doesn't get shaded pink
            # just because it's Friday, and Inactive still shades correctly.
            if (client not in SNAP_KIND_ONLY_CLIENTS
                    and client not in LOAD_AS_DELIVERY_CLIENTS
                    and client not in FORCED_INACTIVE):
                if is_friday_or_later_in_week(today, day) and not isinstance(marker, date):
                    return True
                return False
        if marker == "No Data":
            if client in FORCE_SHADE_NO_DATA:
                return True
            rng = MONTHLY_EXPECTED_DAY_RANGE.get(client)
            if rng:
                try:
                    # Shade when today is >7 days past the START of the expected
                    # window for the CELL's month (not today's). Future-month
                    # templates must not pink-shade No Data.
                    start_day = date(day.year, day.month,
                                     min(rng[0], calendar.monthrange(day.year, day.month)[1]))
                    return (today - start_day).days > 7
                except ValueError:
                    return False
            return False
        # Weekly clients — pink-shade an empty past-scheduled cell when the
        # scheduled day is 3+ days in the past with no cert / activity.
        # Monthly clients keep the 7-day threshold above. Daily Aetnas are
        # exempt (per user 2026-06-03 — they're in DAILY_CLIENTS so this
        # guard is belt-and-suspenders). IMPLEMENTATION_LOAD_ONLY clients
        # are also exempt (per user 2026-06-05 — ElevanceMMMRx showed a
        # pink "!" because we're not actively working that client yet).
        # NYShip_Rx is included even though it's not in WEEKLY_CLIENTS — it
        # uses NYSHIP_DAYS placement but is weekly-equivalent (per user
        # 2026-06-05). Threshold changed from `> 3` to `>= 3` same day so
        # a Tuesday cell pinks on Friday.
        weekly_eligible = (client in WEEKLY_CLIENTS) or client == "NYShip_Rx"
        if (not marker
                and weekly_eligible
                and client not in {"AetnaHRP", "AetnaRCE", "AetnaRx",
                                   "NCStateAetna"}
                and client not in IMPLEMENTATION_LOAD_ONLY_CLIENTS
                and (today - day).days >= 3):
            return True
        # Past-day cells reflect historical activity — don't shade them with
        # the current Inactive/Failure state of the client. Only the marker
        # text itself (handled above) determines shading for past days.
        if day < today:
            return False
        if has_inactive_jobs(client, ramp_jobs, cert_idx, snap_idx, today):
            return True
        if day == today and has_recent_failure(client, ramp_queue, ramp_jobs, today):
            # Suppress the failure shade if the client is currently loading
            # again — the retry has superseded the stale failed state.
            if not is_loading_today(client, ramp_queue, ramp_jobs):
                return True
        return False

    today_week_start = today - timedelta(days=today.weekday())
    today_week_end   = today_week_start + timedelta(days=4)

    def resolve_marker(client, day, allow_checkmark, allow_week_window):
        """Return the Date-column marker for a client placed on `day`.
          - date           : DHT certified that day (or within the 7-day cycle for weekly)
          - "Load Failure" : recent failed load (precedence over L)
          - "L"            : currently loading and scheduled day is in current week
          - "✓"            : snap/load completed (only if allow_checkmark)
          - "Inactive"     : forced inactive client
          - ""             : nothing yet
        """
        # Forced-inactive clients show "Inactive" only for the current day
        # and future days — past-day cells keep their normal markers so a
        # newly-disabled client doesn't retroactively erase its prior cert
        # dates / ✓ from when it was active. Per user 2026-05-18: HealthNetCA
        # past Mondays should still show their cert dates. Past-day cells with
        # no activity fall back to "Inactive" at the end of this function.
        forced_inactive = client in FORCED_INACTIVE
        if forced_inactive and day >= today:
            return "Inactive"
        # cert on the exact day
        ts = cert_on_day(client, day, cert_idx)
        if ts:
            return ts.date()
        # cert anywhere in the 7-day cycle starting at this scheduled day
        if allow_week_window:
            ts = cert_in_week(client, day, cert_idx)
            if ts:
                return ts.date()
        in_current_week = today_week_start <= day <= today_week_end

        # Daily clients: on today's cell, L outranks Failure (active retry
        # is more useful than a stale failure). On past days, only ✓ applies.
        # Monday cells also look back Sat+Sun to catch weekend ETL loads
        # (e.g. NCStateAetna's Saturday 'Aetna RCE 310 ETL Load' → Monday ✓).
        if not allow_week_window:
            if day == today:
                if is_loading_today(client, ramp_queue, ramp_jobs):
                    return "L"
                if has_recent_failure(client, ramp_queue, ramp_jobs, today):
                    return "Load Failure"
            win_back = 2 if day.weekday() == 0 else 0     # Monday → look back Sat/Sun
            # KaiserPrePayCOB renders Sat/Sun loads on their own (Sat)/(Sun)
            # injected rows attached to Fri/Mon — don't let the Mon cell's
            # regular row pick up the Sun load too (would double-count).
            if client == KAISER_PREPAY_CLIENT:
                win_back = 0
            # Any past daily cell looks forward 1 day to catch a load
            # that crossed midnight or ran as a next-day catch-up
            # (e.g. AetnaHRP load for 5/13 finishing 5/14 → ✓ on 5/13).
            win_forward = 1 if day < today else 0
            if allow_checkmark and snap_on_day(
                    client, day, snap_idx,
                    window_days=win_back, forward_days=win_forward):
                return "✓"
            # Past-day "L" for SNAP_KIND_ONLY daily clients (e.g. AetnaHRP):
            # if the load ran on `day` but no snap completion yet AND a job
            # is currently active for the client, the cycle is still in
            # progress — keep that prior weekday's cell at "L". Per user
            # 2026-05-19: "AetnaHRP did not Snap yet from the 5/18/26 load …
            # The 5/18/26 HRP should still be an 'L'."
            if (client in SNAP_KIND_ONLY_CLIENTS
                    and day < today
                    and today_week_start <= day <= today_week_end):
                keys = list(_keys_for_client(client))
                load_on_d = any(
                    len(entry) > 3
                    and entry[3] in ("load", "tape")
                    and _src_matches_client(entry[0], keys)
                    for entry in snap_idx.get(day, ())
                )
                if load_on_d and is_loading_today(client, ramp_queue, ramp_jobs):
                    return "L"
            # Forced-inactive clients with no prior cert/✓ on this past day
            # fall back to "Inactive" rather than leaving the cell blank.
            # Past Mondays with real cert dates already returned above.
            if forced_inactive:
                return "Inactive"
            return ""

        # Implementation-load-only clients (e.g. ElevanceMMMRx): not being
        # actively worked yet — show L only while a job is currently running
        # in RAMP; clear to blank as soon as load+snap completes. Never ✓,
        # never a cert date. Per user 2026-06-03.
        if client in IMPLEMENTATION_LOAD_ONLY_CLIENTS:
            if in_current_week and is_loading_today(client, ramp_queue, ramp_jobs):
                return "L"
            return ""

        # Weekly clients: currently-loading L outranks past failure (active
        # retry is more useful than a stale Failed entry). Cert already
        # took priority above, so cert dates aren't displaced.
        if in_current_week:
            if is_loading_today(client, ramp_queue, ramp_jobs):
                return "L"
            if has_recent_failure(client, ramp_queue, ramp_jobs, today):
                return "Load Failure"
        if allow_checkmark:
            if snap_in_week(client, day, snap_idx):
                return "✓"
        elif in_current_week:
            # Weekly cert client (not snap-only) — if load/snap activity has
            # happened this week and the cert hasn't landed yet, stay L
            # (per user: "CenteneRx & WellCareRx should have an 'L' since
            # they have not been certified").
            if snap_in_week(client, day, snap_idx):
                return "L"
        # Forced-inactive weekly clients with no prior cert in this cycle
        # fall back to "Inactive" rather than leaving the cell blank.
        # Per user 2026-05-19: "HealthNetCA & TuftsRx for 5/18 is somehow
        # not marked 'Inactive' anymore."
        if forced_inactive:
            return "Inactive"
        return ""

    def expected_end_day(client):
        """Return the END day of the monthly client's expected delivery range.
        Falls back to historical avg, then 15th."""
        rng = MONTHLY_EXPECTED_DAY_RANGE.get(client)
        if rng:
            return rng[1]
        avg = average_cert_day(client, cert_idx)
        return avg if avg is not None else 15

    def _kaiser_amb_anchor():
        """Closest Thursday (weekday 3) to the 15th of (year, month)."""
        anchor = date(year, month, 15)
        diff = 3 - anchor.weekday()  # diff is in [-3, 3]
        return anchor + timedelta(days=diff)

    def _apply_weekday_spread(d, client):
        """If the client has a MONTHLY_PLACEMENT_WEEKDAY override, snap d to
        the matching weekday of d's Mon-Fri work-week. Clamps within month."""
        target_wd = MONTHLY_PLACEMENT_WEEKDAY.get(client)
        if target_wd is None:
            return d
        week_start = d - timedelta(days=d.weekday())
        candidate = week_start + timedelta(days=target_wd)
        if candidate.year == year and candidate.month == month:
            return candidate
        return d

    def determine_monthly(client):
        """Return (placement_date, marker) for a monthly client.
        Cert/snap dates remain on their actual date; all other markers are
        anchored to the client's expected delivery day (end of its range)."""
        # Kaiser_Amb feeds all anchor to the closest Thursday to the 15th
        # (per user 2026-05-26).
        if client.startswith("Kaiser_Amb"):
            placeholder = _kaiser_amb_anchor()
        else:
            # expected placement day (end of range; or avg if no range; fallback 15th)
            expected_d = expected_end_day(client)
            try:
                placeholder = date(year, month, min(expected_d, calendar.monthrange(year, month)[1]))
            except ValueError:
                placeholder = date(year, month, 15)
            if client in CLOSEST_WEEKDAY_CLIENTS:
                placeholder = closest_weekday(placeholder)
            else:
                placeholder = next_monday_if_weekend(placeholder)
            placeholder = _apply_weekday_spread(placeholder, client)
        # 0) Forced-inactive clients always show "Inactive" on expected day
        if client in FORCED_INACTIVE:
            return placeholder, "Inactive"
        # 0a) Explicit one-off placement override (per-client day + marker).
        # Highest precedence so EDW feeds can stay on 5/20 even though they
        # certified 5/21, or AetnaQNXT can show on 5/19 with L. The override
        # only applies when its date falls in the current calendar month.
        # Marker "AUTO" anchors placement to ov_day but resolves the marker
        # dynamically: cert this month → cert date; currently loading or
        # loaded this month → "L"; else "No Data".
        override = MONTHLY_PLACEMENT_OVERRIDES.get(client)
        if override:
            ov_day, ov_marker = override
            if ov_day.year == year and ov_day.month == month:
                if ov_marker == "AUTO":
                    c_latest = latest_cert(client, cert_idx, on_or_before=today)
                    if c_latest and c_latest.year == year and c_latest.month == month:
                        return ov_day, c_latest.date()
                    if is_loading_today(client, ramp_queue, ramp_jobs):
                        return ov_day, "L"
                    if has_recent_failure(client, ramp_queue, ramp_jobs, today):
                        return ov_day, "Load Failure"
                    if load_this_month(client, snap_idx, year, month, today):
                        return ov_day, "L"
                    return ov_day, "No Data"
                return ov_day, ov_marker
        # 0b) Snap-disabled clients (load runs but snap step is disabled in RAMP)
        # show marker "Snap" with pink shading on their expected day.
        if client in SNAP_DISABLED_CLIENTS:
            return placeholder, "Snap"
        # expected_date mirrors placeholder for Kaiser_Amb feeds and any
        # spread-adjusted client; otherwise recompute from the range end.
        if client.startswith("Kaiser_Amb"):
            expected_date = placeholder
        else:
            try:
                expected_date = date(year, month, min(expected_d, calendar.monthrange(year, month)[1]))
            except ValueError:
                expected_date = date(year, month, 15)
            if client in CLOSEST_WEEKDAY_CLIENTS:
                expected_date = closest_weekday(expected_date)
            else:
                expected_date = next_monday_if_weekend(expected_date)
            if expected_date.month != month:
                expected_date = date(year, month, calendar.monthrange(year, month)[1])
            expected_date = _apply_weekday_spread(expected_date, client)

        # 1) Already certified this month → place on actual cert date
        c_latest = latest_cert(client, cert_idx, on_or_before=today)
        if c_latest and c_latest.year == year and c_latest.month == month:
            d = c_latest.date()
            d = next_monday_if_weekend(d) if d.weekday() >= 5 else d
            return d, c_latest.date()

        # "Today" rules only fire when today actually falls within the target
        # month — future-month templates must not pull placements back into the
        # current week (per user 2026-05-26: BCBSFL Elig was missing from June
        # because is_loading_today returned today=5/26 and the row got dropped).
        today_in_month = (today.year == year and today.month == month)
        # Cert-only clients (BCBSKS/BCBSKSMedAdv/BCBSSCRx) stay on expected
        # day until DHT cert lands.
        if client in MONTHLY_CERT_ONLY_CLIENTS:
            # Blank-until-cert clients (BCBSKSMedAdv) ignore mid-process
            # activity entirely — they stay empty until a real cert arrives.
            if client in MONTHLY_BLANK_UNTIL_CERT:
                return expected_date, ""
            if today_in_month and is_loading_today(client, ramp_queue, ramp_jobs):
                return expected_date, "L"
            if has_recent_failure(client, ramp_queue, ramp_jobs, today):
                return expected_date, "Load Failure"
            # Show L only when the actual LOAD job has run this month (not
            # just a stray /Ramp/Snap completion). Per user 2026-05-15:
            # "BCBSSC RX 0110 Load has not started loading yet this month —
            # don't show as loading."
            ln = load_this_month(client, snap_idx, year, month, today)
            if ln:
                return expected_date, "L"
            return expected_date, "No Data"

        # 2) Currently loading right now → today + L (outranks past completion).
        if today_in_month and is_loading_today(client, ramp_queue, ramp_jobs):
            return today, "L"

        # 3) Recent failure today → today + Load Failure
        if today_in_month and has_recent_failure(client, ramp_queue, ramp_jobs, today):
            return today, "Load Failure"

        # 4) Successfully snapped this month → ✓ on the snap date.
        # Catches monthly clients whose snap completed but cert is pending.
        sn = latest_snap_this_month(client, snap_idx, year, month, today)
        if sn:
            d = sn.date()
            if d.weekday() >= 5:
                d = next_monday_if_weekend(d)
            return d, "✓"

        # 4b) No snap yet, but the LOAD job has run this month → keep at "L"
        # on the expected day. Per user 2026-05-15: Kaiser_AmbN has loaded
        # but not snapped yet — should stay "L" instead of "No Data".
        ln = load_this_month(client, snap_idx, year, month, today)
        if ln:
            return expected_date, "L"

        # 5) ADO delivery ticket changed this week → move up to that day with L.
        # Only valid in the current month's rendering.
        ticket = latest_tickets.get(client) if today_in_month else None
        if ticket:
            ch = parse_dt(ticket.get("changed", "")) or parse_dt(ticket.get("created", ""))
            if ch and today_week_start <= ch.date() <= today_week_end \
               and ch.date() <= today:
                state = (ticket.get("state") or "").lower()
                if state in ("active", "in progress", "new", "committed"):
                    d = ch.date()
                    if d.weekday() >= 5:
                        d = next_monday_if_weekend(d)
                    return d, "L"

        # 5) Inactive / No Data on expected day
        if has_inactive_jobs(client, ramp_jobs, cert_idx, snap_idx, today):
            return expected_date, "Inactive"
        return expected_date, "No Data"

    def place(bucket, kind, client, day, marker_override=None, label_prefix=""):
        # `from_manual` is True only when the marker comes from the
        # MANUAL_OVERRIDES dict — caller-passed marker_override values (e.g.
        # determine_monthly results) are still eligible for sticky-cert restore.
        from_manual = False
        if marker_override is None:
            mov = MANUAL_OVERRIDES.get((client, day))
            if mov is not None:
                marker_override = mov
                from_manual = True
        if marker_override is not None:
            marker = marker_override
        elif kind == "daily":
            marker = resolve_marker(client, day, allow_checkmark=True, allow_week_window=False)
        elif kind == "weekly":
            allow_check = client in SNAP_ONLY_CLIENTS
            marker = resolve_marker(client, day, allow_checkmark=allow_check, allow_week_window=True)
        elif kind == "kaiser":
            marker = resolve_marker(client, day, allow_checkmark=True, allow_week_window=False)
        else:  # monthly
            marker = resolve_marker(client, day, allow_checkmark=True, allow_week_window=False)
        alert  = alert_state(client, day, marker)
        marker, alert = apply_sticky_cert(client, day, marker, alert, from_manual)
        wk_start = day - timedelta(days=day.weekday())
        wk_end   = wk_start + timedelta(days=4)
        # ESIPBMRx state-list: per-state placement by round (load → snap → ✓).
        # Each tape row is a state-file load. A state is "covered" by the
        # next snap completion that lands AFTER its FileLoadDate — those
        # states get ✓ on the snap date. States loaded after the most recent
        # snap (or never snapped) show "L" on their load date.
        # Scan the WHOLE month — bucket[cell_day] routes each row to the
        # week-block where cell_day falls during rendering.
        if client == "ESIPBMRx" and esipbmrx_tape:
            tape_in_month = []
            for row in esipbmrx_tape:
                fdt = row.get("FileLoadDate")
                if not fdt or fdt.year != year or fdt.month != month:
                    continue
                sm = ESIPBMRX_STATE_RE.search(row.get("FileName", "") or "")
                if sm:
                    tape_in_month.append((sm.group(1).upper(), fdt))
            if tape_in_month:
                snap_completions = []
                for d_, entries in snap_idx.items():
                    if d_.year != year or d_.month != month:
                        continue
                    for entry in entries:
                        kind_e = entry[3] if len(entry) > 3 else "snap"
                        if entry[0] == "esipbmrx" and kind_e == "snap":
                            snap_completions.append(entry[1])
                snap_completions.sort()

                def covering_snap(load_dt):
                    for sdt in snap_completions:
                        if sdt >= load_dt:
                            return sdt
                    return None

                # Group state codes by (cell_day, marker). Dedupe states
                # within a group so a state with multiple files in one round
                # only appears once per cell.
                # Per user 2026-05-18: if a load/snap finishes over the
                # weekend, it should still appear on the closest weekday
                # (Sat → previous Friday, Sun → next Monday) instead of
                # disappearing because the calendar only renders Mon-Fri.
                def _closest_weekday(d):
                    if d.weekday() == 5:
                        return d - timedelta(days=1)   # Sat → Fri
                    if d.weekday() == 6:
                        return d + timedelta(days=1)   # Sun → Mon
                    return d
                state_groups = defaultdict(set)
                for state, load_dt in tape_in_month:
                    cov = covering_snap(load_dt)
                    if cov:
                        cell_day = _closest_weekday(cov.date())
                        marker_s = "✓"
                    else:
                        cell_day = _closest_weekday(load_dt.date())
                        marker_s = "L"
                    state_groups[(cell_day, marker_s)].add(state)

                highlight = "yellow" if client in YELLOW_HIGHLIGHT else None
                for (cell_day, marker_s) in sorted(state_groups,
                                                   key=lambda k: (k[0], 0 if k[1] == "✓" else 1)):
                    states_sorted = sorted(state_groups[(cell_day, marker_s)])
                    chunks = [states_sorted[i:i+4] for i in range(0, len(states_sorted), 4)]
                    alert_s = alert_state(client, cell_day, marker_s)
                    for chunk in chunks:
                        extra = f" ({', '.join(chunk)})"
                        label = f"{label_prefix}{display_name(client, monthly=(kind=='monthly'), extra_suffix=extra)}"
                        bucket[cell_day].append((label, marker_s, alert_s, highlight))
                return
        # Multi-week load flag for CenteneRx and similar
        extra = ""
        if client in MULTI_WEEK_CLIENTS and multi_week_loads:
            n = count_multi_week_loads(client, wk_start, wk_end, multi_week_loads)
            if n > 1:
                extra = f" ({n} weeks)"
        label = f"{label_prefix}{display_name(client, monthly=(kind=='monthly'), extra_suffix=extra)}"
        if client in BOLD_LABEL:
            highlight = "bold"
        elif client in YELLOW_HIGHLIGHT:
            highlight = "yellow"
        else:
            highlight = None
        # If this is a Load Failure with a registered ADO ticket, attach
        # the work-item URL so the marker cell becomes a clickable link.
        link = None
        if marker == "Load Failure":
            ticket = LOAD_FAILURE_ADO_LINKS.get((client, day))
            if ticket:
                link = ADO_LINK.format(ticket)
        bucket[day].append((label, marker, alert, highlight, link))

    # daily clients on every weekday (alphabetical)
    for d in all_days:
        for c in sorted(DAILY_CLIENTS):
            # Blank-before clients: skip cells before their go-live cutoff
            # (e.g. ElevanceMMMRx shows June 2026 forward only).
            bb = BLANK_BEFORE.get(c)
            if bb and d < bb:
                continue
            place(daily, "daily", c, d)

    # KaiserPrePayCOB — placed in the DAILY section (sorted alphabetically with
    # the other daily clients) per user 2026-06-16. (Was previously its own
    # bottom-of-week row in the `kaiser` bucket.)
    for d in all_days:
        place(daily, "daily", KAISER_PREPAY_CLIENT, d)

    # weekly clients on assigned weekday (alphabetical within column)
    for c in sorted(WEEKLY_CLIENTS):
        days = WEEKLY_CLIENTS[c]
        impl_start = IMPLEMENTATION_CLIENTS.get(c)
        blank_before = BLANK_BEFORE.get(c)
        for d in all_days:
            if d.strftime("%A") not in days:
                continue
            # Blank-before clients: render empty (skip the cell) for any day
            # before the cutoff — no marker, no pink "!".
            if blank_before and d < blank_before:
                continue
            # Implementation clients are suppressed before their start date.
            if impl_start and d < impl_start:
                continue
            # During implementation phase (before first cert), render
            # "Implementation" in the date cell as the manual marker.
            if impl_start:
                latest = latest_cert(c, cert_idx, on_or_before=d)
                if latest is None or latest.date() < impl_start:
                    place(weekly, "weekly", c, d, marker_override="Implementation")
                    continue
            place(weekly, "weekly", c, d)

    # NYShip_Rx rotates 4x/month
    for daynum in NYSHIP_DAYS:
        try:
            tgt = next_monday_if_weekend(date(year, month, daynum))
        except ValueError:
            continue
        if tgt.month != month:
            continue
        label = f"NYShip_Rx ({NYSHIP_LABEL[daynum]})"
        ov = NYSHIP_OVERRIDES.get((year, month, daynum))
        if ov is not None:
            # One-off override (e.g. 1st & 8th combined cycle): force the
            # marker. date(...) renders as the cert date; strings ("L") render
            # as-is. A date is never alerted; a string defers to alert_state.
            marker = ov
            alert  = False if isinstance(ov, date) else alert_state("NYShip_Rx", tgt, ov)
            marker, alert = apply_sticky_cert("NYShip_Rx", tgt, marker, alert, True)
        else:
            # Cert-style client — stay L when loaded this week until cert lands.
            marker = resolve_marker("NYShip_Rx", tgt, allow_checkmark=False, allow_week_window=True)
            alert  = alert_state("NYShip_Rx", tgt, marker)
            marker, alert = apply_sticky_cert("NYShip_Rx", tgt, marker, alert, False)
        weekly[tgt].append((label, marker, alert, None))

    # Monthly clients: state-driven placement (cert→cert date, loading→today,
    # snap→snap date, failure→today, else→expected day).
    for c in sorted(MONTHLY_CLIENTS):
        # OptumPBMRx is special — placed twice/month (early-month + end-month sets).
        if c == "OptumPBMRx":
            continue
        d, marker = determine_monthly(c)
        if d.month != month:
            continue
        place(monthly, "monthly", c, d, marker_override=marker)

    # OptumPBMRx: two delivery sets per month (RAW1/2/3 early + RAW5/6 late).
    # For each half of the month place separately based on tape-load activity.
    def first_friday(y, m):
        d = date(y, m, 1)
        return d + timedelta(days=(4 - d.weekday()) % 7)

    # Per user 2026-06-05: ✓ requires ALL of the half's RAW files to be
    # tape-loaded AND a matching number of 'Optum 0200 PBM Start Snap'
    # successful completions in the same window. Partial load → L.
    _OPTUM_RAW_RE = re.compile(r"RAW\s*0*(\d+)_", re.IGNORECASE)
    _OPTUM_SNAP_RE = re.compile(r"optum.*0200.*pbm.*start.*snap", re.IGNORECASE)
    _optum_job_by_id = {j.get("JobId"): (j.get("JobName") or "") for j in ramp_jobs}

    def place_optum_half(label_suffix, day_lo, day_hi, required_raws,
                         default_day=None):
        last_day = calendar.monthrange(year, month)[1]
        hi = min(day_hi, last_day)
        window_start = date(year, month, day_lo)
        window_end   = date(year, month, hi)

        # Scan TRGETL3 OptumPBMRx tape rows for filenames matching the
        # required RAW numbers within the window.
        loaded_raws = {}   # raw_n → latest FileLoadDate
        for tape in (optumpbmrx_tape or ()):
            fname = tape.get("FileName") or ""
            dt = tape.get("FileLoadDate")
            if not (dt and window_start <= dt.date() <= window_end):
                continue
            m = _OPTUM_RAW_RE.search(fname)
            if not m:
                continue
            raw_n = int(m.group(1))
            if raw_n not in required_raws:
                continue
            if raw_n not in loaded_raws or dt > loaded_raws[raw_n]:
                loaded_raws[raw_n] = dt

        # Count successful 'Optum 0200 PBM Start Snap' completions in window.
        snap_count = 0
        latest_snap_dt = None
        for q in ramp_queue:
            jn = _optum_job_by_id.get(q.get("JobId"), "")
            if not _OPTUM_SNAP_RE.search(jn):
                continue
            status = (q.get("Status") or "").lower()
            if not (status.startswith("success") or status == "resolved"):
                continue
            end_dt = parse_dt(q.get("EndDate"))
            if not end_dt or not (window_start <= end_dt.date() <= window_end):
                continue
            snap_count += 1
            if latest_snap_dt is None or end_dt > latest_snap_dt:
                latest_snap_dt = end_dt

        need = len(required_raws)
        all_loaded  = len(loaded_raws) >= need
        all_snapped = snap_count >= need

        if all_loaded and all_snapped:
            candidates = list(loaded_raws.values())
            if latest_snap_dt:
                candidates.append(latest_snap_dt)
            latest = max(candidates)
            placement = next_monday_if_weekend(latest.date())
            if placement.month != month:
                placement = date(year, month, last_day)
            marker = "✓"
        elif loaded_raws or (today.year == year and today.month == month
                             and window_start <= today <= window_end
                             and is_loading_today("OptumPBMRx", ramp_queue, ramp_jobs)):
            # Partial load (any RAW present but not all) OR actively loading
            # right now — stay L.
            if loaded_raws:
                latest = max(loaded_raws.values())
                placement = next_monday_if_weekend(latest.date())
                if placement.month != month:
                    placement = date(year, month, last_day)
            elif today.year == year and today.month == month and window_start <= today <= window_end:
                placement = today
            elif default_day is not None:
                placement = default_day
            else:
                mid = day_lo + (hi - day_lo) // 2
                placement = next_monday_if_weekend(date(year, month, mid))
            marker = "L"
        else:
            # Nothing yet — placeholder.
            if default_day is not None:
                placement = default_day
            else:
                mid = day_lo + (hi - day_lo) // 2
                placement = next_monday_if_weekend(date(year, month, mid))
            marker = "No Data"

        # Show what's pending so the team can see which RAW files are still
        # outstanding mid-cycle (only when partial-load L state).
        extra = ""
        if marker == "L":
            missing = sorted(set(required_raws) - set(loaded_raws.keys()))
            if missing:
                extra = " — RAW " + ",".join(str(n) for n in missing) + " pending"
            elif not all_snapped:
                extra = " — Snap pending"
        label = f"OptumPBMRx {label_suffix}{extra}"
        alert = False
        if marker == "No Data":
            ref_end = date(year, month, min(hi, last_day))
            alert = (today - ref_end).days > 7
        monthly[placement].append((label, marker, alert, None))

    place_optum_half("(RAW 1/2/3)", 1, 7, {1, 2, 3},
                     default_day=first_friday(year, month))
    place_optum_half("(RAW 5/6)",   24, 31, {5, 6})

    # Aetna NMSP - MMSEA: ✓ once SourceLog shows a NonMSP file fully imported
    # (ImportCompleteDate) this month, placed on the completion date. While a
    # file has only started loading in CMSE (ImportStartDate, no complete date)
    # the cell shows "L" on the start date. Otherwise the 15th rule (or next
    # Monday) with No Data.
    nmsp_day = nmsp_mmsea_date(year, month)
    if nmsp_day.month == month:
        done_dt = None   # latest completed load this month
        start_dt = None  # latest in-progress (started, not yet complete) load
        for rec in (aetna_nmsp_loads or ()):
            d, s = rec.get("done"), rec.get("start")
            if d and d.year == year and d.month == month and d.date() <= today:
                if done_dt is None or d > done_dt:
                    done_dt = d
            if s and not d and s.year == year and s.month == month and s.date() <= today:
                if start_dt is None or s > start_dt:
                    start_dt = s
        if done_dt:
            placement = done_dt.date()
            if placement.weekday() >= 5:
                placement = next_monday_if_weekend(placement)
            if placement.month != month:
                placement = nmsp_day
            marker = "✓"
            alert  = False
        elif start_dt:
            placement = start_dt.date()
            if placement.weekday() >= 5:
                placement = next_monday_if_weekend(placement)
            if placement.month != month:
                placement = nmsp_day
            marker = "L"
            alert  = False
        else:
            placement = nmsp_day
            marker = "No Data"
            alert = alert_state("AetnaMMSEA", nmsp_day, marker)
        monthly[placement].append(("Aetna NMSP - MMSEA", marker, alert, "bold"))

    # (Removed) loading-today extras pass — L is now surfaced on each client's
    # scheduled weekday cell via resolve_marker / determine_monthly directly.

    # One-off injected entries (catch-up loads, etc.)
    section_map = {"daily": daily, "weekly": weekly, "monthly": monthly, "kaiser": kaiser}
    for section, day, label, marker, alert, highlight in ADDITIONAL_ENTRIES:
        bucket = section_map.get(section)
        if bucket is None or day.month != month or day.year != year:
            continue
        bucket[day].append((label, marker, alert, highlight))

    # (KaiserPrePayCOB Sat/Sun weekend tracking removed 2026-06-16 per user —
    # no need to surface Saturday/Sunday loads anymore.)

    # Ad-hoc MSPI loads -> "Monthly Ad Hoc" (per user 2026-06-25): each MSPI
    # client (Aetna MSPI, BCBSNC MSPI, Aetna QNXT MSPI) appears ONCE per month,
    # on the date of its LATEST ad-hoc load that month, in the MONTHLY section.
    # No persistent daily row, no per-run weekly entries, never flagged missing
    # (matches HumanaRx, which is already monthly). Weekend-start loads shift to
    # an adjacent weekday for display.
    since_adhoc = date(year, month, 1) - timedelta(days=14)
    mspi_latest = {}   # label -> (orig_day, marker, alert)
    for ah in scan_adhoc_loads(ramp_queue, ramp_jobs, today, since_adhoc,
                               weekend_shift=False):
        d = ah["day"]
        if d.year != year or d.month != month:
            continue
        prev = mspi_latest.get(ah["label"])
        if prev is None or d >= prev[0]:
            mspi_latest[ah["label"]] = (d, ah["marker"], ah["alert"])
    for label, (d, mk, al) in mspi_latest.items():
        disp = d
        if disp.weekday() == 5:
            disp -= timedelta(days=1)
        elif disp.weekday() == 6:
            disp += timedelta(days=1)
        monthly[disp].append((label, mk, al, None, None))

    # Ad-hoc MONTHLY cert clients (per user 2026-06-25): surface ONCE per month
    # in the MONTHLY section on the DHT cert date, ONLY after certified. No cert
    # this month -> no row (never flagged missing). e.g. UnitedRx.
    for label, db in ADHOC_MONTHLY_CERT_CLIENTS.items():
        best = None
        for key in _keys_for_client(db):
            for dt, status in cert_idx.get(key, ()):
                if status == "Certified" and dt.year == year and dt.month == month:
                    if best is None or dt > best:
                        best = dt
        if best is not None:
            monthly[best.date()].append((label, best.date(), False, None, None))

    # CignaRx EOM/SOM injection — second CignaRx cycle closing out prior month
    # surfaces on the first Tuesday of each month. Per user 2026-06-03.
    cigna_target = None
    for d in all_days:
        if d.month == month and d.weekday() == 1:  # Tuesday
            cigna_target = d
            break
    if cigna_target is not None:
        cig_label = "CignaRx (EOM/SOM)(p)"
        override = CIGNARX_EOM_SOM_OVERRIDES.get((year, month))
        if override is not None:
            marker = override
        else:
            # Window: 7 days before SOM through 14 days into month — catches
            # the EOM load tail and the (often exception-cert) early-month
            # cert landing.
            win_start = date(year, month, 1) - timedelta(days=7)
            win_end   = date(year, month, 1) + timedelta(days=14)
            # Cert preference: earliest cert in the window.
            cig_cert = None
            for key in _keys_for_client("CignaRx"):
                for dt, status in cert_idx.get(key, ()):
                    if status != "Certified":
                        continue
                    if win_start <= dt.date() <= win_end:
                        if cig_cert is None or dt < cig_cert:
                            cig_cert = dt
            if cig_cert and cig_cert.date() <= today:
                marker = cig_cert.date()
            else:
                cig_keys = list(_keys_for_client("CignaRx"))
                loaded = False
                for d_scan, entries in snap_idx.items():
                    if not (win_start <= d_scan <= win_end):
                        continue
                    for entry in entries:
                        if _src_matches_client(entry[0], cig_keys):
                            jn = entry[4] if len(entry) > 4 else ""
                            if _load_name_allowed("CignaRx", jn, entry[3] if len(entry) > 3 else "load"):
                                loaded = True
                                break
                    if loaded:
                        break
                marker = "L" if loaded else "No Data"
        alert = alert_state("CignaRx", cigna_target, marker)
        weekly[cigna_target].append((cig_label, marker, alert, None))

    # One-off (per user 2026-06-11): Kaiser_AmbM runs a SECOND June cycle — it
    # certified 6/11, and a new monthly load lands ~6/18. determine_monthly only
    # places one row/month and the 6/11 cert takes precedence, so the new cycle
    # wouldn't surface until it re-certifies. Inject a 6/18 row that tracks the
    # new load (No Data → L) using only activity on/after the 6/12 cutoff (so the
    # prior 6/11 cert/load is ignored). Once the new cert lands, determine_monthly
    # shows it on the cert date and this injection is skipped (no duplicate).
    # Remove after the June 2026 cycle.
    if (year, month) == (2026, 6):
        ka_place  = date(2026, 6, 18)
        ka_cutoff = date(2026, 6, 12)        # ignore the prior 6/11 cert/load
        ka_keys   = list(_keys_for_client("Kaiser_AmbM"))
        ka_new_cert = any(
            status == "Certified" and ka_cutoff <= dt.date() <= today
            for k in ka_keys for dt, status in cert_idx.get(k, ())
        )
        if not ka_new_cert:
            ka_loaded = any(
                _src_matches_client(e[0], ka_keys)
                for d_scan, entries in snap_idx.items() if ka_cutoff <= d_scan <= today
                for e in entries
            )
            ka_marker = "L" if (is_loading_today("Kaiser_AmbM", ramp_queue, ramp_jobs)
                                or ka_loaded) else "No Data"
            monthly[ka_place].append(("Kaiser_AmbM", ka_marker, False, None))

    # Kaiser Submission daily row — ✓ when both 'Kaiser Pareo Submission Logfile'
    # AND '...Upload' finished that day; 'L' on today while in progress; blank
    # otherwise. Placed before the sort so DAILY_ORDER positions it (between
    # KaiserPrePayCOB and NCStateAetna).
    ks_done, ks_running = scan_kaiser_submission(
        ramp_queue, ramp_jobs, today, date(year, month, 1) - timedelta(days=14))
    for d in all_days:
        if d in ks_done:
            mk = "✓"
        elif d == today and ks_running:
            mk = "L"
        else:
            mk = ""
        daily[d].append(("Kaiser Submission", mk, False, None, None))

    # Sort: Daily by the fixed DAILY_ORDER; Weekly/Monthly/Kaiser alphabetical.
    for d in daily:
        daily[d].sort(key=lambda r: _daily_order_key(r[0]))
    for bucket in (weekly, monthly, kaiser):
        for d in bucket:
            bucket[d].sort(key=lambda r: r[0].lower())

    return {"daily": daily, "weekly": weekly, "monthly": monthly, "kaiser": kaiser}, weeks


# ============================================================
#                           rendering (no colour fills)
# ============================================================
HEADER_FILL = PatternFill("solid", fgColor="2C5F8A")
HEADER_FONT = Font(name="Segoe UI", bold=True, color="FFFFFF", size=11)
DAY_FILL    = PatternFill("solid", fgColor="E3EBF4")
DAY_FONT    = Font(name="Segoe UI", bold=True, size=10, color="1F3D5C")
TODAY_FILL  = PatternFill("solid", fgColor="FFD180")
ALERT_FILL  = PatternFill("solid", fgColor="FFC7CE")
ALERT_FONT  = Font(name="Segoe UI", size=9, bold=True, color="9C0006")
YELLOW_FILL = PatternFill("solid", fgColor="FFF2A8")
YELLOW_FONT = Font(name="Segoe UI", size=9, bold=True, color="7F6000")
CELL_FONT   = Font(name="Segoe UI", size=9)
THIN        = Side(style="thin", color="C8C8C8")
BORDER      = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)


def fmt_marker(m):
    """Pass through date objects for native Excel date formatting; stringify
    anything else (✓, L, No Data, Inactive, Load Failure, blank…)."""
    if isinstance(m, date):
        return m
    return str(m or "")


def _write_section_rows(ws, cur_row, wk, plan_section):
    """Write one section (daily/weekly/monthly) of stacked rows for a week.
    Returns the next free row after the section.
    Each row spans the 5 day-columns; cell content + date marker per column.
    """
    max_clients = max((len(plan_section.get(d, [])) for d in wk if d), default=0)
    for ci in range(max_clients):
        for i, d in enumerate(wk):
            col_day = i * 2 + 1
            col_dat = i * 2 + 2
            cell      = ws.cell(row=cur_row, column=col_day, value=None)
            date_cell = ws.cell(row=cur_row, column=col_dat, value=None)
            for c in (cell, date_cell):
                c.font = CELL_FONT
                c.border = BORDER
                c.alignment = Alignment(vertical="top")
            date_cell.alignment = Alignment(horizontal="center", vertical="top")
            if d is None:
                cell.fill = PatternFill("solid", fgColor="F5F5F5")
                date_cell.fill = PatternFill("solid", fgColor="F5F5F5")
                continue
            clients = plan_section.get(d, [])
            if ci < len(clients):
                row = clients[ci]
                name, marker, alert = row[0], row[1], row[2]
                highlight = row[3] if len(row) > 3 else None
                link      = row[4] if len(row) > 4 else None
                cell.value = name
                v = fmt_marker(marker)
                date_cell.value = v
                if isinstance(v, date):
                    date_cell.number_format = "mm/dd/yy"
                if highlight == "yellow":
                    cell.fill = YELLOW_FILL
                    cell.font = YELLOW_FONT
                    date_cell.fill = YELLOW_FILL
                elif highlight == "bold":
                    cell.font = Font(name="Segoe UI", size=9, bold=True)
                if alert:
                    date_cell.fill = ALERT_FILL
                    if not date_cell.value:
                        date_cell.value = "!"
                    date_cell.font = ALERT_FONT
                if link:
                    # Underline the alert font so the cell visibly reads
                    # as a clickable link (cursor changes on hover too).
                    date_cell.hyperlink = link
                    date_cell.font = Font(name="Segoe UI", size=9,
                                          bold=True, color="9C0006",
                                          underline="single")
        cur_row += 1
    return cur_row


def _blank_separator_row(ws, cur_row):
    for i in range(10):
        c = ws.cell(row=cur_row, column=i + 1, value=None)
        c.fill = PatternFill("solid", fgColor="FFFFFF")
        c.border = Border()
    ws.row_dimensions[cur_row].height = 6
    return cur_row + 1


# Section-header band (Daily / Weekly / Monthly). Doubles as the visual
# separator between sections, so no blank row is needed alongside it.
SECTION_LABEL_FILL = PatternFill("solid", fgColor="E8EDF3")
SECTION_LABEL_FONT = Font(name="Segoe UI", bold=True, size=9, color="2C5F8A")


def _section_label_row(ws, cur_row, text):
    """Write a thin labeled band spanning all 10 columns as a section header."""
    for i in range(10):
        c = ws.cell(row=cur_row, column=i + 1, value=None)
        c.fill = SECTION_LABEL_FILL
        c.border = Border()
    c0 = ws.cell(row=cur_row, column=1, value=text)
    c0.font = SECTION_LABEL_FONT
    c0.alignment = Alignment(horizontal="left", vertical="center")
    ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=10)
    ws.row_dimensions[cur_row].height = 14
    return cur_row + 1


def write_weekly_stacked(ws, year, month, sections, weeks, today):
    holidays = us_federal_holidays(year)
    cur_row = 1
    week_no = 0
    for wk in weeks:
        week_no += 1
        first_d = next((d for d in wk if d), None)
        last_d  = next((d for d in reversed(wk) if d), None)
        label = f"Week {week_no}: {first_d:%m/%d} – {last_d:%m/%d}" if first_d and last_d else f"Week {week_no}"
        ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=10)
        c = ws.cell(row=cur_row, column=1, value=label)
        c.font = Font(name="Segoe UI", bold=True, size=12, color="1F3D5C")
        cur_row += 1

        for i, day in enumerate(WEEKDAYS):
            col_day = i * 2 + 1
            col_dat = i * 2 + 2
            for col, val in ((col_day, day), (col_dat, "Date")):
                cell = ws.cell(row=cur_row, column=col, value=val)
                cell.fill = HEADER_FILL
                cell.font = HEADER_FONT
                cell.alignment = Alignment(horizontal="center")
                cell.border = BORDER
        cur_row += 1

        # date strip row (real Excel dates, formatted mm/dd/yy).
        # Federal-holiday labels are written into the day-name sub-column
        # of this same row so the label sits directly under "Monday" and
        # right next to the date — not as a separate banner row.
        holiday_fill = PatternFill("solid", fgColor="FFE4B5")
        holiday_font = Font(name="Segoe UI", italic=True, size=9,
                            bold=True, color="7C4A00")
        for i, d in enumerate(wk):
            col_day = i * 2 + 1
            col_dat = i * 2 + 2
            if d:
                cell = ws.cell(row=cur_row, column=col_dat, value=d)
                cell.number_format = "mm/dd/yy"
                cell.fill = DAY_FILL if d != today else TODAY_FILL
                cell.font = DAY_FONT
                cell.alignment = Alignment(horizontal="center")
                cell.border = BORDER
                hname = holidays.get(d)
                if hname:
                    hc = ws.cell(row=cur_row, column=col_day, value=hname)
                    hc.fill = holiday_fill
                    hc.font = holiday_font
                    hc.alignment = Alignment(horizontal="center", vertical="center")
                    hc.border = BORDER
        cur_row += 1

        # Each section is introduced by its header band (which also separates
        # it from the section above — no blank row needed). KaiserPrePayCOB now
        # lives in the Daily section; the `kaiser` bucket is only populated for
        # historical snapshot months and is rendered after Monthly when present.
        cur_row = _section_label_row(ws, cur_row, "Daily")
        cur_row = _write_section_rows(ws, cur_row, wk, sections["daily"])
        cur_row = _section_label_row(ws, cur_row, "Weekly")
        cur_row = _write_section_rows(ws, cur_row, wk, sections["weekly"])
        cur_row = _section_label_row(ws, cur_row, "Monthly")
        cur_row = _write_section_rows(ws, cur_row, wk, sections["monthly"])
        if any(sections["kaiser"].get(d) for d in wk if d):
            cur_row = _write_section_rows(ws, cur_row, wk, sections["kaiser"])

        # per-week key block
        ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=10)
        kc = ws.cell(row=cur_row, column=1,
                     value="Key:  Date = Certified  |  ✓ = Loaded/Snapped  |  L = Loading"
                           "  |  pink = Failure/Inactive  |  (s) SLA  |  (p) Rx Post Snap"
                           "  |  (n) Not Delivered  |  -  = No load that day")
        kc.font = Font(name="Segoe UI", italic=True, size=9, color="555555")
        kc.alignment = Alignment(horizontal="left")
        cur_row += 2

    # Client-name columns set to ≈190 px (width 27.07) per user 2026-05-20.
    # Excel pixels ≈ 7 * width + 0.5.
    # Column D (Tuesday date) was widened to ≈95 px for BCBSAR's "Implementation"
    # text; reverted to the standard 11 once BCBSAR certified (2026-06-12).
    client_w = (190 - 0.5) / 7
    for i in range(10):
        w = client_w if i % 2 == 0 else 11
        ws.column_dimensions[get_column_letter(i + 1)].width = w
    ws.sheet_view.showGridLines = False


def _write_month_section(ws, cur_row, wk, plan_section):
    """Single-column-per-day version for the Month sheet."""
    max_clients = max((len(plan_section.get(d, [])) for d in wk if d), default=0)
    for ci in range(max_clients):
        for col, d in enumerate(wk, start=1):
            cell = ws.cell(row=cur_row, column=col, value=None)
            cell.font = CELL_FONT
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = BORDER
            if d is None:
                cell.fill = PatternFill("solid", fgColor="F5F5F5")
                continue
            clients = plan_section.get(d, [])
            if ci < len(clients):
                row = clients[ci]
                name, marker, alert = row[0], row[1], row[2]
                highlight = row[3] if len(row) > 3 else None
                m_str = fmt_marker(marker)
                cell.value = f"{name}  [{m_str}]" if m_str else name
                if highlight == "yellow":
                    cell.fill = YELLOW_FILL
                    cell.font = YELLOW_FONT
                if alert:
                    cell.fill = ALERT_FILL
                    cell.font = ALERT_FONT
        cur_row += 1
    return cur_row


def write_month_sheet(ws, year, month, sections, weeks, today):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5)
    tc = ws.cell(row=1, column=1, value=f"{calendar.month_name[month]} {year} — Delivery Calendar")
    tc.font = Font(name="Segoe UI", bold=True, size=14, color="1F3D5C")
    tc.alignment = Alignment(horizontal="center")
    ws.row_dimensions[1].height = 26

    for i, d in enumerate(WEEKDAYS, start=1):
        c = ws.cell(row=2, column=i, value=d)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = Alignment(horizontal="center")
        c.border = BORDER
        ws.column_dimensions[get_column_letter(i)].width = 32

    cur_row = 3
    for wk in weeks:
        for col, d in enumerate(wk, start=1):
            label = d.strftime("%a %m/%d") if d else ""
            cell = ws.cell(row=cur_row, column=col, value=label)
            cell.fill = TODAY_FILL if d == today else DAY_FILL
            cell.font = DAY_FONT
            cell.alignment = Alignment(horizontal="center")
            cell.border = BORDER
        cur_row += 1

        cur_row = _write_month_section(ws, cur_row, wk, sections["daily"])
        cur_row += 1  # blank
        cur_row = _write_month_section(ws, cur_row, wk, sections["weekly"])
        cur_row += 1
        cur_row = _write_month_section(ws, cur_row, wk, sections["monthly"])
        cur_row += 1
        cur_row = _write_month_section(ws, cur_row, wk, sections["kaiser"])

        # per-week key
        ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=5)
        kc = ws.cell(row=cur_row, column=1,
                     value="Key:  Date = Certified  |  ✓ = Loaded/Snapped  |  L = Loading"
                           "  |  pink Date = Failure/Inactive  |  (s) SLA  |  (p) Rx Post Snap"
                           "  |  (n) Not Delivered  |  -  = No load that day")
        kc.font = Font(name="Segoe UI", italic=True, size=9, color="555555")
        cur_row += 2

    ws.sheet_view.showGridLines = False


def write_tickets_sheet(ws, latest_tickets, ramp_jobs):
    headers = ["Client", "Title Kind", "ADO Ticket", "ADO State", "Tags",
               "RAMP Jobs (enabled)", "Assigned To", "Changed"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="left")
        cell.border = BORDER

    rows = sorted(latest_tickets.values(), key=lambda t: (t["state"], t["client"].lower()))
    for r_i, t in enumerate(rows, start=2):
        matched = find_matching_jobs(t["client"], ramp_jobs)
        enabled = sum(1 for j in matched if j.get("Enabled") == 1)
        vals = [
            t["client"], t["kind"], f"#{t['id']}", t["state"], t["tags"],
            enabled, t["assigned"], (t["changed"][:10] if t["changed"] else ""),
        ]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(row=r_i, column=c, value=v)
            cell.font = CELL_FONT
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = BORDER

    ws.freeze_panes = "A2"
    for i, w in enumerate([22, 18, 12, 14, 28, 14, 24, 14], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def write_key_sheet(ws):
    rows = [
        ("Marker / Suffix",  "Meaning"),
        ("[Date]",           "Client was certified that day in DHT.TableList (CertTimestamp)."),
        ("Snap",             "Snap completed that day in RAMP /api/Ramp/Snap/SnapQueueStatus."),
        ("L",                "Load or snap currently in progress for the client today."),
        ("(s)",              "SLA Client — tight delivery window."),
        ("(p)",              "Rx Client Post Snap."),
        ("(n)",              "Special handling — historically Not Delivered."),
        ("M -",              "Monthly client — placed on the day its delivery ticket fired."),
    ]
    for c, h in enumerate(rows[0], start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.border = BORDER
    for r, row in enumerate(rows[1:], start=2):
        for c, val in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=val)
            cell.font = CELL_FONT
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    notes_row = len(rows) + 3
    notes = [
        "Data sources",
        f"  • SQL  TRGUTIL10.DHTStats [DHT].[TableList]  (CertTimestamp + CurrentStatus)",
        f"  • RAMP {RAMP_BASE}/api/Ramp/Snap/SnapQueueStatus  (snap completions)",
        f"  • RAMP {RAMP_BASE}/api/Ramp/Queue/List           (load-job completions)",
        f"  • RAMP {RAMP_BASE}/api/Ramp/Job/List             (enabled-job detection)",
        f"  • ADO  {ADO_BASE}  (User Stories tagged 'Delivery Ticket')",
        "",
        "Title formats recognised on ADO tickets:",
        "  • 'Snap and Mine - <Client> - ...'",
        "  • 'Load and Snap - <Client> - ...'",
        "  • 'Kaiser - SNAP/MINE - <Client> - ...'",
    ]
    for i, n in enumerate(notes):
        cell = ws.cell(row=notes_row + i, column=1, value=n)
        cell.font = Font(name="Segoe UI", bold=(i == 0 or n.startswith("Title")), size=10)

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 90
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False


def write_client_owner_sheet(ws):
    """Render the Client Owner tab — four owner groups side-by-side, each
    a (Owner, Client, Priority) sub-block. The Owner sub-column is merged
    vertically across the whole block with the name centered both axes.
    Each owner's 3-column header is a distinct color. A blank row
    separates each owner's upper list from their lower list, matching the
    source layout.
    """
    # Distinct header colors per owner. White HEADER_FONT remains readable
    # on each. Per user 2026-05-20.
    owner_header_colors = {
        "Dave":     "2C5F8A",   # blue (matches main report header)
        "Emmanuel": "4A7C59",   # green
        "Holly":    "7C4A6E",   # mauve
        "Adam":     "B86F2E",   # orange
    }

    owners = list(CLIENT_OWNERS.items())
    upper_max = max(len(d["upper"]) for _, d in owners)
    lower_max = max(len(d["lower"]) for _, d in owners)

    blank_row = 2 + upper_max               # row index where the gap sits
    bottom_end_row = blank_row + lower_max  # last data row

    owner_font   = Font(name="Segoe UI", bold=True, size=14, color="1F3D5C")
    owner_align  = Alignment(horizontal="center", vertical="center")
    center       = Alignment(horizontal="center", vertical="center")
    plain        = Alignment(vertical="top", wrap_text=False)

    for col_idx, (owner, data) in enumerate(owners):
        col_o = col_idx * 3 + 1
        col_c = col_idx * 3 + 2
        col_p = col_idx * 3 + 3

        # ---- per-owner header row (row 1) ----
        fill = PatternFill("solid", fgColor=owner_header_colors.get(owner, "2C5F8A"))
        for col, label in ((col_o, "Owner"), (col_c, "Client"), (col_p, "Priority")):
            cell = ws.cell(row=1, column=col, value=label)
            cell.fill = fill
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal="center")
            cell.border = BORDER

        # ---- pre-border the Client + Priority columns ----
        for row in range(2, bottom_end_row + 1):
            for col in (col_c, col_p):
                ws.cell(row=row, column=col).border = BORDER

        # ---- merge Owner column vertically and center the name ----
        ws.merge_cells(start_row=2, start_column=col_o,
                       end_row=bottom_end_row, end_column=col_o)
        oc = ws.cell(row=2, column=col_o, value=owner)
        oc.font = owner_font
        oc.alignment = owner_align
        # Apply BORDER to every cell within the merged Owner range so the
        # bottom edge actually renders in Excel — without this, a merged
        # cell only borders the top-left occurrence and the bottom edge of
        # the merged block disappears. Per user 2026-05-20.
        for r in range(2, bottom_end_row + 1):
            ws.cell(row=r, column=col_o).border = BORDER

        # ---- upper entries ----
        for r_off, entry in enumerate(data["upper"]):
            row = 2 + r_off
            client, priority = entry
            cc = ws.cell(row=row, column=col_c, value=client)
            cc.font = CELL_FONT
            cc.alignment = plain
            pc = ws.cell(row=row, column=col_p, value=priority)
            pc.font = CELL_FONT
            pc.alignment = center

        # ---- lower entries (after the blank separator row) ----
        for r_off, entry in enumerate(data["lower"]):
            row = blank_row + 1 + r_off
            client, priority = entry
            cc = ws.cell(row=row, column=col_c, value=client)
            cc.font = CELL_FONT
            cc.alignment = plain
            pc = ws.cell(row=row, column=col_p, value=priority)
            pc.font = CELL_FONT
            pc.alignment = center

    # Column widths — user-requested pixel targets (105 px Owner, 185 px Client).
    # Excel width-to-pixels: pixels ≈ 7 * width + 0.5 (Calibri 11 baseline).
    # Owner widened 2026-05-20 from 91→105 px to give the centered name more room.
    OWNER_WIDTH    = (105 - 0.5) / 7   # ≈ 14.93
    CLIENT_WIDTH   = (185 - 0.5) / 7   # ≈ 26.36
    PRIORITY_WIDTH = 9
    for col_idx in range(len(owners)):
        ws.column_dimensions[get_column_letter(col_idx * 3 + 1)].width = OWNER_WIDTH
        ws.column_dimensions[get_column_letter(col_idx * 3 + 2)].width = CLIENT_WIDTH
        ws.column_dimensions[get_column_letter(col_idx * 3 + 3)].width = PRIORITY_WIDTH

    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False


# ============================================================
#                            main
# ============================================================
def parse_all_clients_xlsx(file_path, year):
    """Read an ExpectedClientDates_<MMM>.xlsx 'All Clients' tab into the
    same `(sections, weeks)` shape that `plan_calendar` returns.

    Two header styles are accepted:
      Jan: `All Clients - Week N - M/D-M/D` (Monday date encoded in header)
      Feb-Apr (and later): `<MonthName> - Week N - All Clients` (no date in
        header — the Monday date is inferred by scanning the first real
        date cell in the block).

    Daily rows: same client name across all populated weekday cells. In the
    Feb+ format, non-cert weekday cells contain the boolean `True` (a load-
    verification checkbox); these become blank markers, not strings.
    """
    from openpyxl import load_workbook
    wb = load_workbook(file_path, data_only=True)
    if "All Clients" not in wb.sheetnames:
        return None, None
    ws = wb["All Clients"]

    sections = {"daily": defaultdict(list), "weekly": defaultdict(list),
                "monthly": defaultdict(list), "kaiser": defaultdict(list)}
    weeks = []

    header_re_jan = re.compile(
        r"All Clients\s*-\s*Week\s*(\d+)\s*-\s*(\d+)/(\d+)-(\d+)/(\d+)",
        re.IGNORECASE,
    )
    header_re_feb = re.compile(
        r"([A-Za-z]+)\s*-\s*Week\s*(\d+)\s*-\s*All\s*Clients",
        re.IGNORECASE,
    )

    # Pending block state when the Monday date hasn't been resolved yet.
    pending_week = None  # dict: {'block_rows': [(row_idx, cells)], ...}
    current_week_dates = None
    in_data_block = False

    def to_marker(raw):
        # Feb+ files put boolean True/False in non-cert date cells as a
        # load-verification flag; treat these as blank, not "True"/"False".
        if isinstance(raw, bool):
            return ""
        if isinstance(raw, datetime):
            return raw.date()
        if isinstance(raw, date):
            return raw
        if raw is None:
            return ""
        return str(raw).strip()

    def _name_matches_fill(display):
        """True when `display` is one of the ALL_CLIENTS_FILL_CHECKMARK
        clients (allowing common suffix conventions like (s), (p), (n),
        " - ", and the leading `M - ` monthly tag)."""
        base = display.strip()
        base = re.sub(r"^M\s*-\s*", "", base).strip()
        base = base.lstrip("* ").strip()
        for target in ALL_CLIENTS_FILL_CHECKMARK:
            if (base == target
                    or base.startswith(target + "(")
                    or base.startswith(target + " ")
                    or base.startswith(target + "-")):
                return True
        return False

    def flush_block(block_rows, week_days):
        """Commit a block of pending rows once week_days has been resolved."""
        for cells in block_rows:
            populated = []
            for i in range(5):
                n = cells[i * 2]
                if n is None:
                    continue
                ns = str(n).strip()
                if not ns:
                    continue
                populated.append((i, ns))
            is_daily_row = False
            daily_name = None
            if populated and len(populated) >= 2:
                base = populated[0][1].lstrip("* ").strip()
                if all(p[1].lstrip("* ").strip() == base for p in populated):
                    is_daily_row = True
                    daily_name = base

            for i in range(5):
                name = cells[i * 2]
                mark = cells[i * 2 + 1]
                if name is None and (mark is None or isinstance(mark, bool)):
                    continue
                name_s = "" if name is None else str(name).strip()
                if not name_s:
                    continue
                display = name_s
                if display.startswith("*"):
                    display = display.lstrip("* ").strip()

                kind = "weekly"
                if is_daily_row:
                    kind = "daily"
                    display = daily_name
                elif display.startswith("M -") or display.startswith("M-"):
                    kind = "monthly"
                    # Strip the "M - " tag from historical snapshot labels too,
                    # so it's gone everywhere (per user 2026-06-16).
                    display = re.sub(r"^M\s*-\s*", "", display).strip()
                elif "KaiserPrePayCOB" in display:
                    kind = "kaiser"

                marker = to_marker(mark)
                # Empty / True placeholder cells for the daily-Aetna and
                # PBMRx group: render as ✓ (loaded/snapped) instead of blank.
                if (isinstance(marker, str) and marker == ""
                        and _name_matches_fill(display)):
                    marker = "✓"
                alert = (isinstance(marker, str) and marker
                         and marker.strip().lower() in ALL_CLIENTS_ALERT_MARKERS)

                cell_day = week_days[i]
                sections[kind][cell_day].append((display, marker, alert))

    def commit_pending():
        nonlocal pending_week
        if pending_week is None:
            return
        # Anchor the Mon-Fri cycle Monday from a date cell whose weekday
        # matches its column position (typical case — most cert dates equal
        # the scheduled cycle day). Falls back to any date's own-week Monday.
        anchor_mon = None
        for i in range(5):
            for cells in pending_week["block_rows"]:
                v = cells[i * 2 + 1]
                d_val = None
                if isinstance(v, datetime):
                    d_val = v.date()
                elif isinstance(v, date) and not isinstance(v, bool):
                    d_val = v
                if d_val is None or d_val.weekday() != i:
                    continue
                anchor_mon = d_val - timedelta(days=i)
                break
            if anchor_mon:
                break
        if anchor_mon is None:
            # Fallback: take the first real date and snap to its own Monday.
            for cells in pending_week["block_rows"]:
                if anchor_mon:
                    break
                for i in range(5):
                    v = cells[i * 2 + 1]
                    d_val = None
                    if isinstance(v, datetime):
                        d_val = v.date()
                    elif isinstance(v, date) and not isinstance(v, bool):
                        d_val = v
                    if d_val is None:
                        continue
                    anchor_mon = d_val - timedelta(days=d_val.weekday())
                    break
        if anchor_mon is None and pending_week.get("month_hint"):
            # Fallback: compute Mon by week number within the month.
            yr = pending_week.get("year_hint", year)
            mn = pending_week["month_hint"]
            wn = pending_week["week_no"]
            cal_obj = calendar.Calendar(firstweekday=0)
            month_mons = [d for d in cal_obj.itermonthdates(yr, mn)
                          if d.weekday() == 0 and (d.month == mn
                                                   or (d - timedelta(days=4)).month == mn)]
            if 1 <= wn <= len(month_mons):
                anchor_mon = month_mons[wn - 1]
        if anchor_mon is None:
            pending_week = None
            return
        week_days = [anchor_mon + timedelta(days=i) for i in range(5)]
        weeks.append(week_days)
        flush_block(pending_week["block_rows"], week_days)
        pending_week = None

    month_name_to_num = {
        m.lower(): i for i, m in enumerate(
            ["", "January", "February", "March", "April", "May", "June",
             "July", "August", "September", "October", "November", "December"]
        ) if m
    }

    for row_idx in range(1, ws.max_row + 1):
        cells = [ws.cell(row=row_idx, column=c).value for c in range(1, 11)]
        first = cells[0]
        first_s = "" if first is None else str(first).strip()

        m1 = header_re_jan.match(first_s)
        m2 = header_re_feb.match(first_s) if not m1 else None
        if m1:
            commit_pending()
            mon_m, mon_d = int(m1.group(2)), int(m1.group(3))
            try:
                mon_date = date(year, mon_m, mon_d)
            except ValueError:
                continue
            week_days = [mon_date + timedelta(days=i) for i in range(5)]
            current_week_dates = week_days
            weeks.append(week_days)
            in_data_block = False
            pending_week = None
            continue
        if m2:
            commit_pending()
            mname = m2.group(1).lower()
            wnum = int(m2.group(2))
            month_hint = month_name_to_num.get(mname)
            pending_week = {
                "block_rows": [],
                "month_hint": month_hint,
                "year_hint": year,
                "week_no": wnum,
            }
            current_week_dates = None
            in_data_block = False
            continue

        if first_s == "Monday" and cells[2] in ("Tuesday", " Tuesday"):
            in_data_block = True
            continue

        if (first_s.startswith("Verify SNAP") or first_s.startswith("Verify Load")
                or first_s.startswith("SNAP pattern") or first_s.startswith("The three way")
                or (in_data_block and not any(c is not None and not isinstance(c, bool)
                                              and str(c).strip() for c in cells))):
            in_data_block = False
            continue

        if not in_data_block:
            continue

        if pending_week is not None:
            pending_week["block_rows"].append(cells)
            continue
        if current_week_dates is None:
            continue
        flush_block([cells], current_week_dates)

    # End-of-file: flush any open pending block.
    commit_pending()
    return sections, weeks


def _html_escape(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def _fmt_marker_html(m):
    """Marker rendering for HTML cells."""
    if isinstance(m, date):
        return m.strftime("%m/%d/%y")
    return _html_escape(m or "")


def _render_section_rows_html(wk, plan_section, today):
    """Return one HTML <tbody> for a section. Each row has 10 cells.
    Per user 2026-06-03: today highlighting is restricted to the date-strip
    row only — data-row cells do NOT get a today-column accent, so the table
    body reads cleanly without a vertical line down every row."""
    max_clients = max((len(plan_section.get(d, [])) for d in wk if d), default=0)
    if max_clients == 0:
        return ""
    rows_html = []
    for ci in range(max_clients):
        cells = []
        for i, d in enumerate(wk):
            if d is None:
                cells.append(f'<td class="name dim-month"></td>')
                cells.append(f'<td class="marker dim-month"></td>')
                continue
            entries = plan_section.get(d, [])
            if ci >= len(entries):
                cells.append(f'<td class="name"></td>')
                cells.append(f'<td class="marker"></td>')
                continue
            row = entries[ci]
            name, marker, alert = row[0], row[1], row[2]
            highlight = row[3] if len(row) > 3 else None
            link      = row[4] if len(row) > 4 else None
            name_classes = ["name", "client-cell"]
            marker_classes = ["marker"]
            if highlight == "yellow":
                name_classes.append("hl-yellow")
                marker_classes.append("hl-yellow")
            elif highlight == "bold":
                name_classes.append("hl-bold")
            if alert:
                marker_classes.append("alert")
            v = _fmt_marker_html(marker)
            if alert and not v:
                v = "!"
            data_client = _html_escape(name or "")
            name_html = data_client
            marker_html = v
            if link:
                marker_classes.append("link")
                marker_html = (f'<a href="{_html_escape(link)}" target="_blank" '
                               f'rel="noopener">{v}</a>')
            cells.append(f'<td class="{" ".join(name_classes)}" '
                         f'data-client="{data_client}">{name_html}</td>')
            cells.append(f'<td class="{" ".join(marker_classes)}">{marker_html}</td>')
        rows_html.append(f'<tr class="data-row">{"".join(cells)}</tr>')
    return "\n".join(rows_html)


def _render_week_card_html(wk, week_no, sections, today, holidays):
    first_d = next((d for d in wk if d), None)
    last_d  = next((d for d in reversed(wk) if d), None)
    if first_d and last_d:
        label = (f"Week {week_no}: {first_d.strftime('%m/%d')} – "
                 f"{last_d.strftime('%m/%d')}")
    else:
        label = f"Week {week_no}"

    # Header + date-strip rows.
    # Per user 2026-06-03: only the actual date cell is highlighted as
    # "today" — no vertical column accent. Header row stays uniform.
    header_cells = []
    date_cells = []
    for i, d in enumerate(wk):
        day_name = WEEKDAYS[i]
        header_cells.append(
            f'<th class="hdr-day">{day_name}</th>'
            f'<th class="hdr-date">Date</th>')
        if d is None:
            date_cells.append(
                '<td class="strip-name dim-month"></td>'
                '<td class="strip-date dim-month"></td>')
            continue
        hname = holidays.get(d, "")
        hname_cls = " holiday" if hname else ""
        is_today = " is-today" if d == today else ""
        date_str = d.strftime("%m/%d/%y")
        date_cells.append(
            f'<td class="strip-name{hname_cls}">{_html_escape(hname)}</td>'
            f'<td class="strip-date{is_today}">{date_str}</td>')

    sec_labels = {"daily": "Daily", "weekly": "Weekly", "monthly": "Monthly"}
    sec_html_parts = []
    for sec_key in ("daily", "weekly", "monthly", "kaiser"):
        body = _render_section_rows_html(wk, sections[sec_key], today)
        if body:
            lbl = sec_labels.get(sec_key)
            label_row = (f'<tr class="sec-label"><td colspan="10">{lbl}</td></tr>'
                         if lbl else "")
            sec_html_parts.append(
                f'<tbody class="sec sec-{sec_key}">{label_row}{body}</tbody>'
            )

    key_line = ("Key:  Date = Certified  |  ✓ = Loaded/Snapped  |  L = Loading"
                "  |  pink = Failure/Inactive  |  (s) SLA  |  (p) Rx Post Snap"
                "  |  (n) Not Delivered  |  -  = No load that day")

    # Explicit column widths so the table sizes to content, not the page.
    # Client-name columns ~165 px, date columns ~70 px, Tuesday date 95 px.
    colgroup = (
        '<colgroup>'
        '<col class="cn"><col class="dt">'
        '<col class="cn"><col class="dt-wide">'
        '<col class="cn"><col class="dt">'
        '<col class="cn"><col class="dt">'
        '<col class="cn"><col class="dt">'
        '</colgroup>'
    )

    return (
        f'<section class="week-card">'
        f'  <div class="week-label">{_html_escape(label)}</div>'
        f'  <table class="grid">'
        f'    {colgroup}'
        f'    <thead class="grid-head"><tr>{"".join(header_cells)}</tr></thead>'
        f'    <tbody class="strip"><tr>{"".join(date_cells)}</tr></tbody>'
        f'    {"".join(sec_html_parts)}'
        f'  </table>'
        f'  <div class="key">{_html_escape(key_line)}</div>'
        f'</section>'
    )


def build_dashboard_html(month_packs, today, current_month_name):
    """Render the dashboard HTML.
    month_packs: list of dicts {name, year, month, sections, weeks, holidays}
    """
    import calendar as _cal_mod

    tabs_html = []
    panels_html = []
    for mp in month_packs:
        is_current = (mp["name"] == current_month_name)
        active = " active" if is_current else ""
        tab_id = f'tab-{mp["name"].replace(" ", "-")}'
        tabs_html.append(
            f'<button class="tab{active}" data-target="{tab_id}">'
            f'{_html_escape(mp["name"])}</button>')

        week_blocks = []
        for week_no, wk in enumerate(mp["weeks"], start=1):
            week_blocks.append(_render_week_card_html(
                wk, week_no, mp["sections"], today, mp["holidays"]))
        panels_html.append(
            f'<section class="month-panel{active}" id="{tab_id}">'
            f'  {"".join(week_blocks)}'
            f'</section>')

    generated = datetime.now().strftime("%Y-%m-%d %H:%M")
    title = "Client Delivery Status"

    css = """
:root {
  --bg: #f4f6f9; --card: #ffffff; --border: #c8c8c8; --text: #1f2a37;
  --muted: #5b6776; --accent: #2C5F8A; --accent-dark: #1F3D5C;
  --day-fill: #E3EBF4; --today-fill: #FFD180; --today-strong: #f08c00;
  --alert: #FFC7CE; --alert-dark: #9C0006;
  --yellow: #FFF2A8; --yellow-dark: #7F6000;
  --holiday: #FFE4B5; --holiday-dark: #7C4A00;
}
* { box-sizing: border-box; }
body {
  margin: 0;
  font-family: "Segoe UI", -apple-system, BlinkMacSystemFont, sans-serif;
  background: var(--bg); color: var(--text); font-size: 13px;
}
header.bar {
  background: var(--accent-dark); color: #fff;
  padding: 10px 18px; position: sticky; top: 0; z-index: 5;
  display: flex; align-items: center; gap: 16px; flex-wrap: wrap;
}
header.bar h1 { margin: 0; font-size: 16px; font-weight: 600; }
header.bar .meta { font-size: 11px; opacity: 0.85; }
header.bar .grow { flex: 1; }
header.bar input[type=search] {
  border: 0; border-radius: 4px; padding: 6px 10px; font-size: 13px; min-width: 220px;
}
header.bar button {
  background: var(--accent); color: #fff; border: 0; border-radius: 4px;
  padding: 6px 12px; cursor: pointer; font-size: 12px;
}
header.bar .legend {
  font-size: 11px; opacity: 0.92; display: flex; align-items: center; gap: 6px;
}
header.bar .legend b {
  display: inline-block; background: var(--alert); color: var(--alert-dark);
  font-weight: 700; border-radius: 3px; padding: 0 6px; line-height: 16px;
}
nav.tabs {
  background: #fff; border-bottom: 1px solid var(--border);
  padding: 4px 12px; position: sticky; top: 44px; z-index: 4;
  display: flex; gap: 4px; flex-wrap: wrap;
}
nav.tabs button.tab {
  background: transparent; color: var(--accent-dark);
  border: 1px solid transparent; border-bottom: 0;
  border-radius: 4px 4px 0 0;
  padding: 6px 12px; cursor: pointer; font-size: 12px; font-weight: 600;
}
nav.tabs button.tab.active {
  background: var(--accent); color: #fff;
}
main { padding: 12px 18px 40px; }
.month-panel { display: none; }
.month-panel.active { display: block; }
.week-card {
  margin: 0 0 14px;
}
.week-label {
  font-weight: 700; font-size: 14px; color: var(--accent-dark);
  margin: 12px 0 4px;
}
/* Tightened 2026-06-03: table sizes to content (width: auto), not the page.
   Explicit column widths via <colgroup>: name ≈ 165 px, date ≈ 70 px,
   Tuesday date 95 px while BCBSAR is in Implementation. */
table.grid {
  border-collapse: separate; border-spacing: 0; table-layout: fixed;
  background: var(--card); width: auto;
}
table.grid col.cn       { width: 165px; }
table.grid col.dt       { width:  70px; }
table.grid col.dt-wide  { width:  95px; }
table.grid th, table.grid td {
  border: 1px solid var(--border); padding: 3px 6px;
  font-size: 12px; vertical-align: top;
  overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
}
table.grid th {
  background: var(--accent); color: #fff;
  font-weight: 600; font-size: 11px; text-align: center; letter-spacing: 0.3px;
}
.strip td { background: var(--day-fill); text-align: center; font-weight: 600; color: var(--accent-dark); }
.strip td.is-today {
  background: var(--today-fill); color: var(--accent-dark);
  box-shadow: inset 0 0 0 2px var(--today-strong);
}
.strip td.holiday { background: var(--holiday); color: var(--holiday-dark); font-style: italic; font-weight: 700; }
.strip td.dim-month, td.dim-month { background: #F5F5F5 !important; color: #aaa; }
td.name { white-space: nowrap; }
td.marker { text-align: center; font-variant-numeric: tabular-nums; }
td.alert { background: var(--alert); color: var(--alert-dark); font-weight: 700; }
td.hl-yellow { background: var(--yellow); color: var(--yellow-dark); font-weight: 700; }
.hl-bold { font-weight: 700; }
td.marker.link a { color: var(--alert-dark); font-weight: 700; text-decoration: underline; }
.sec-gap td { background: #fff; border: 0; height: 4px; padding: 0; }
.sec-label td { background: #E8EDF3; color: #2C5F8A; font-weight: 700; font-size: 10px; text-transform: uppercase; letter-spacing: 0.5px; padding: 2px 6px; border: 0; }
.key {
  font-style: italic; font-size: 11px; color: #555;
  margin: 4px 0 12px;
}
/* Search: dim all client-name cells, then re-emphasize the matched ones
   only. Per user 2026-06-03: highlight just the client name, not the
   whole row. */
body.search-active td.client-cell { opacity: 0.25; }
body.search-active td.client-cell.match-cell { opacity: 1; background: #fffbe8; font-weight: 600; }
"""

    js = """
(function() {
  function show(tabBtn) {
    document.querySelectorAll('nav.tabs button.tab').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.month-panel').forEach(p => p.classList.remove('active'));
    tabBtn.classList.add('active');
    const target = document.getElementById(tabBtn.dataset.target);
    if (target) target.classList.add('active');
    window.scrollTo({ top: 0 });
  }
  document.querySelectorAll('nav.tabs button.tab').forEach(btn => {
    btn.addEventListener('click', () => show(btn));
  });
  // Search
  const search = document.getElementById('search');
  function applySearch() {
    const q = search.value.trim().toLowerCase();
    document.body.classList.toggle('search-active', q.length > 0);
    // Per-cell highlight: only the matching client-name cell pops; the rest
    // of the row stays at normal/dimmed weight.
    document.querySelectorAll('td.client-cell').forEach(td => {
      const c = (td.dataset.client || '').toLowerCase();
      const isMatch = q && c && c.indexOf(q) !== -1;
      td.classList.toggle('match-cell', isMatch);
    });
  }
  search.addEventListener('input', applySearch);
  // Today jump
  const todayBtn = document.getElementById('today-jump');
  if (todayBtn) {
    todayBtn.addEventListener('click', () => {
      const el = document.querySelector('.strip td.is-today');
      if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
    });
  }
})();
"""

    head = (
        "<!doctype html><html lang='en'><head>"
        "<meta charset='utf-8'>"
        f"<title>{_html_escape(title)}</title>"
        f"<style>{css}</style>"
        "</head>"
    )
    body = (
        "<body>"
        "<header class='bar'>"
        f"  <h1>{_html_escape(title)}</h1>"
        f"  <div class='meta'>Generated {generated}</div>"
        "  <div class='legend'><b>!</b> = flagged / needs attention "
        "(past due or problem, no delivery recorded yet)</div>"
        "  <div class='grow'></div>"
        "  <input type='search' id='search' placeholder='Filter by client name…' autocomplete='off'>"
        "  <button id='today-jump'>Jump to today</button>"
        "</header>"
        f"<nav class='tabs'>{''.join(tabs_html)}</nav>"
        f"<main>{''.join(panels_html)}</main>"
        f"<script>{js}</script>"
        "</body></html>"
    )
    return head + body


def main():
    today = date.today()
    year, month = today.year, today.month
    month_start = date(year, month, 1)

    print(f"[info] Today: {today}  Month: {year}-{month:02d}")

    print("[info] Querying DHT cert table…")
    # pull 3 months back for monthly clients' avg-day calc
    certs = fetch_dht_certs(since=date(year, month, 1) - timedelta(days=90))
    cert_idx = build_cert_index(certs)
    global CERT_WEEK_IDX
    CERT_WEEK_IDX = build_cert_week_index(certs)
    print(f"[info]   {len(certs)} cert rows / {len(cert_idx)} distinct DatabaseNames")

    print("[info] Fetching ADO tickets…")
    tickets = fetch_ado_tickets(min_changed_date=month_start - timedelta(days=14))
    print(f"[info]   {len(tickets)} delivery tickets in window")

    print("[info] Fetching RAMP jobs…")
    jobs = fetch_ramp_jobs()
    enabled_n = sum(1 for j in jobs if j.get("Enabled") == 1)
    print(f"[info]   {len(jobs)} total ({enabled_n} enabled)")

    # ----- Auto-discover new MasterLoad 0110 Load implementations (per user
    # 2026-06-03, updated 2026-06-08). New clients default to Weekly/Monday
    # placement as CERT-STYLE clients: they stay "L" until the cert date lands
    # and never show a ✓ (no SNAP_ONLY membership). Per user 2026-06-08: "All
    # new 'MasterLoad 0110 Load' will also be in this format." The old
    # PBMRx→SNAP_KIND_ONLY (✓ after snap) and non-PBMRx→IMPLEMENTATION_LOAD_ONLY
    # (blank after snap) branches were removed. (ElevanceMMMRx was promoted to a
    # DAILY_CLIENTS client 2026-06-16 and is now "known", so it's no longer
    # rediscovered here.)
    new_impls = find_unconfigured_masterload_clients(jobs)
    if new_impls:
        for entry in new_impls:
            client = entry["raw"]
            print(f"[info]   NEW MasterLoad implementation: {client} "
                  f"(pbmrx={entry['pbmrx']}, enabled={entry['enabled']})")
            if client not in WEEKLY_CLIENTS:
                WEEKLY_CLIENTS[client] = ["Monday"]
            # Add alias so RAMP job lookups find the new client.
            aliases = CLIENT_ALIASES.setdefault(client, [])
            if entry["normalized"] not in aliases:
                aliases.append(entry["normalized"])

    # ----- Auto-Inactive clients whose primary 0100/0110 jobs are all
    # Inactive in RAMP (Enabled=0). Kaiser feeds excluded per user.
    auto_inactive = auto_inactive_from_ramp(jobs)
    if auto_inactive:
        print(f"[info]   Auto-Inactive (0100/0110 disabled, non-Kaiser): "
              f"{sorted(auto_inactive)}")
        FORCED_INACTIVE.update(auto_inactive)

    print("[info] Fetching RAMP queue + snap history…")
    queue = fetch_ramp_queue()
    snaps = fetch_ramp_snaps()
    print(f"[info]   queue={len(queue)}, snaps={len(snaps)}")

    print("[info] Fetching tape loads…")
    since_dt = date(year, month, 1) - timedelta(days=14)
    tape_loads = {}
    for client, (db, src_key) in TAPE_LOAD_SOURCES.items():
        server = TAPE_LOAD_SERVER.get(client, "TRGETL3")
        name_like = TAPE_LOAD_NAME_FILTER.get(client)
        rows = fetch_tape_loads(db, since_dt, server=server, name_like=name_like)
        tape_loads[src_key] = rows
        print(f"[info]   {server}.{db}: {len(rows)} rows")

    print("[info] Fetching multi-week client tape loads…")
    multi_week_loads = {}
    for client, (db, _pattern) in MULTI_WEEK_CLIENTS.items():
        rows = fetch_tape_loads(db, since_dt)
        multi_week_loads[client] = rows
        print(f"[info]   {client}@{db}: {len(rows)} rows")

    snap_idx = build_snap_index(jobs, queue, snaps, tape_loads=tape_loads)
    print(f"[info] snap index dates: {len(snap_idx)}")

    print("[info] Querying SQLUtilAudit for Aetna NMSP NonMSP loads…")
    aetna_nmsp_loads = fetch_aetna_nmsp_loads(since=date(year, month, 1) - timedelta(days=30))
    print(f"[info]   {len(aetna_nmsp_loads)} Aetna NonMSP entries")

    latest_tickets, monthly_placements = build_ticket_index(tickets, jobs)
    print(f"[info] latest tickets indexed for {len(latest_tickets)} clients")

    wb = Workbook()
    # Drop the auto-created default sheet so we control ordering.
    del wb[wb.sheetnames[0]]

    # Client Owner sits first.
    ws_owner = wb.create_sheet("Client Owner")
    write_client_owner_sheet(ws_owner)

    import calendar as _cal_mod
    current_tab_name = f"{_cal_mod.month_name[month]} {year}"

    # Generate every month of the year so prior months don't drop off as they
    # conclude (per user 2026-06-03). Compute current month last in the loop
    # iteration but build tabs in Jan→Dec order for natural tab ordering.
    load_sticky_certs()
    print(f"[info] sticky cert cache: {len(STICKY_CERTS)} remembered cells")

    month_packs = []
    for m in range(1, 13):
        sec_m = wk_m = None
        # For closed months Jan–Apr 2026 the manually-maintained
        # ExpectedClientDates_*.xlsx files are the source of truth. Live DHT
        # / RAMP history rolls off too quickly to reconstruct.
        snapshot_file = EXPECTED_DATES_FILES.get(m) if year == 2026 else None
        if snapshot_file:
            try:
                snap_path = os.path.join(EXPECTED_DATES_DIR, snapshot_file)
                if os.path.exists(snap_path):
                    sec_m, wk_m = parse_all_clients_xlsx(snap_path, year)
                    if sec_m is not None and wk_m:
                        print(f"[info]   {_cal_mod.month_name[m]}: loaded from {snapshot_file}")
                    else:
                        sec_m = wk_m = None
            except Exception as e:
                print(f"[warn] failed to parse {snapshot_file}: {e}")
                sec_m = wk_m = None

        if sec_m is None or wk_m is None:
            sec_m, wk_m = plan_calendar(year, m, cert_idx, snap_idx,
                                        latest_tickets, monthly_placements,
                                        jobs, queue,
                                        esipbmrx_tape=tape_loads.get("esipbmrx"),
                                        multi_week_loads=multi_week_loads,
                                        aetna_nmsp_loads=aetna_nmsp_loads,
                                        optumpbmrx_tape=tape_loads.get("optumpbmrx"))

        tab_name = f"{_cal_mod.month_name[m]} {year}"
        ws_m = wb.create_sheet(tab_name)
        write_weekly_stacked(ws_m, year, m, sec_m, wk_m, today)
        month_packs.append({
            "name": tab_name, "year": year, "month": m,
            "sections": sec_m, "weeks": wk_m,
            "holidays": us_federal_holidays(year),
        })
        if m == month:
            sections, weeks = sec_m, wk_m

    # Persist any newly-seen cert dates so future runs can't regress them to "!".
    save_sticky_certs()
    print(f"[info] sticky cert cache saved: {len(STICKY_CERTS)} cells")

    # Open the workbook to the current-month tab by default.
    wb.active = wb.sheetnames.index(current_tab_name)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    out_path = os.path.join(OUTPUT_DIR, f"ClientDeliveryStatus_{today:%Y-%m-%d}.xlsx")
    try:
        wb.save(out_path)
    except PermissionError:
        # file is open in Excel — fall back to a timestamped name so we can still verify
        fallback = os.path.join(OUTPUT_DIR,
                                f"ClientDeliveryStatus_{today:%Y-%m-%d}_{datetime.now():%H%M%S}.xlsx")
        wb.save(fallback)
        out_path = fallback
        print(f"[warn] primary file locked; wrote {fallback} instead")
    print(f"[done] Wrote {out_path}")

    # Project-folder copy (dated filename for archival inspection).
    try:
        os.makedirs(LOCAL_COPY_DIR, exist_ok=True)
        local_path = os.path.join(LOCAL_COPY_DIR,
                                  f"ClientDeliveryStatus_{today:%Y-%m-%d}.xlsx")
        import shutil
        shutil.copyfile(out_path, local_path)
        print(f"[done] Local copy: {local_path}")
    except Exception as e:
        print(f"[warn] Local-copy failed: {e}")

    # OneDrive copy with a FIXED filename so a single Notion link stays valid
    # across runs (OneDrive auto-syncs to SharePoint; same URL serves latest).
    try:
        os.makedirs(os.path.dirname(ONEDRIVE_COPY_PATH), exist_ok=True)
        import shutil
        shutil.copyfile(out_path, ONEDRIVE_COPY_PATH)
        print(f"[done] OneDrive copy: {ONEDRIVE_COPY_PATH}")
    except Exception as e:
        print(f"[warn] OneDrive-copy failed: {e}")

    # Dashboard HTML — self-contained file with search + today emphasis +
    # sticky toolbar/tabs. Same color codes and weekly-stacked layout as the
    # .xlsx, just rendered in the browser. Sits alongside the .xlsx.
    try:
        html_str = build_dashboard_html(month_packs, today, current_tab_name)
        html_paths = [
            os.path.join(OUTPUT_DIR, "ClientDeliveryStatus.html"),
            os.path.join(LOCAL_COPY_DIR, "ClientDeliveryStatus.html"),
            os.path.join(os.path.dirname(ONEDRIVE_COPY_PATH),
                         "ClientDeliveryStatus.html"),
        ]
        wrote_one = False
        for p in html_paths:
            try:
                os.makedirs(os.path.dirname(p), exist_ok=True)
                with open(p, "w", encoding="utf-8") as f:
                    f.write(html_str)
                print(f"[done] HTML dashboard: {p}")
                wrote_one = True
            except Exception as e:
                print(f"[warn] HTML write failed for {p}: {e}")
        if not wrote_one:
            print("[warn] HTML dashboard not written to any path")
    except Exception as e:
        print(f"[warn] HTML dashboard build failed: {e}")

    return out_path


if __name__ == "__main__":
    main()
