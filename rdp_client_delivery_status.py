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
    "BCBSNorthCarolinaFEP": ["bcbsncfep"],
    # NCStateAetna's daily load runs through 'Aetna RCE 310 ETL Load' (Feed
    # "RCE Medical" → key "aetnarce" after stripping). Include aetnarce alias
    # so NCStateAetna gets ✓ from those same daily completions.
    "NCStateAetna":         ["ncstateaetna", "aetnarce"],
    "WPRxDMGCOBMining":     ["wellpointedwardrxdmgcobmining", "wpwedmgcobmining"],
    "HumanaRx":             ["humanarx"],
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
    # MMOH (no Rx): only 'MMOH Claims 0110 Load' counts as the load
    # indicator. JobName "MMOH Claims 0110 Load" → snap_idx key "mmohclaims".
    "MMOH":                 ["mmoh", "mmohclaims"],
    # MMOHRx weekly Tue: only 'MMOHRx Weekly Claim 0110 Load' counts. COBC
    # alias dropped 2026-05-18 so MMOHRx COBC Successful loads don't trip L.
    "MMOHRx":               ["mmohrx", "mmohrxweeklyclaim"],
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
# Rendered at top of each week, alphabetical. KaiserPrePayCOB is also daily but
# is rendered separately at the bottom of each week per user preference.
DAILY_CLIENTS = ["AetnaHRP", "AetnaRCE", "AetnaRx", "NCStateAetna"]
KAISER_PREPAY_CLIENT = "KaiserPrePayCOB"  # rendered last in each week

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
    "UPMC":                  ["Monday"],
    # === TUESDAY ===
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
    "Wellmark":              ["Tuesday"],
    # === WEDNESDAY ===
    "CareSource":            ["Wednesday"],
    "CareSourceRx":          ["Wednesday"],
    "CenteneFidelis":        ["Wednesday"],
    "CenteneFidelisRx":      ["Wednesday"],
    "EmblemRx":              ["Wednesday"],
    "ExcellusRx":            ["Wednesday"],
    "HarvardPilgrim":        ["Wednesday"],
    "Medica":                ["Wednesday"],
    "WellpointEdwardRx":     ["Wednesday"],
    # === THURSDAY ===
    "HealthNewEngland":      ["Thursday"],
    "Kaiser_CO":             ["Thursday"],
    "Kaiser_GA":             ["Thursday"],
    "Kaiser_HI":             ["Thursday"],
    "Kaiser_MASTapestry":    ["Thursday"],
    "Kaiser_NW":             ["Thursday"],
    "KaiserNCPareo":         ["Thursday"],
    "KaiserSCPareo":         ["Thursday"],
    "Oscar":                 ["Thursday"],
    "OscarRx":               ["Thursday"],
    "Premera":               ["Thursday"],
    "PrimePBMRx":            ["Thursday"],
    # === FRIDAY ===
    "CenteneRx":             ["Friday"],
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
# Kaiser_AmbM removed 2026-05-19: no longer inactive, but snap is disabled —
# handled separately via SNAP_DISABLED_CLIENTS below.
FORCED_INACTIVE = {"Tufts_PublicPlan", "TuftsRx", "HealthNetCA"}

# Clients whose load is running but snap step is disabled in RAMP — show
# marker "Snap" with pink shading on the expected delivery day.
# Per user 2026-05-19: "Kaiser_AmbM is no longer inactive, but the Snap is
# disabled. Keep shaded in pink, but put 'Snap' in cell."
SNAP_DISABLED_CLIENTS = {"Kaiser_AmbM"}

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
    # MMOH (monthly): only 'MMOH Claims 0110 Load' is the load indicator.
    # Tightened 2026-05-19 per user: prior pattern "claim" caught ancillary
    # jobs (Start Snap step, Claim 0150 RTA Load) and falsely L'd the cell
    # on days like 5/19. The narrower "claims 0110 load" pattern only matches
    # the actual monthly load step.
    "MMOH":              ("claims 0110 load",),
    # MMOHRx weekly Tue: only 'MMOHRx Weekly Claim 0110 Load' counts —
    # filter excludes the Monthly Claim Stage and Weekly Claim Stage.
    "MMOHRx":            ("weekly claim 0110 load",),
    # CignaRx (weekly Tue): only 'Cigna RX 0110 Load' counts. COBC Load,
    # Daily PassFile, and other ancillary jobs share the "cignarx" matching
    # prefix and would otherwise trip the L indicator. Per user 2026-05-19:
    # "'Cigna RX 0110 Load' is not running, so CignaRx should not have an 'L'."
    # Pattern "rx 0110 load" matches "Cigna RX 0110 Load" but not
    # "Cigna RX COBC 0110 Load" (the "COBC" between "RX" and "0110" breaks
    # the substring).
    "CignaRx":           ("rx 0110 load",),
    "WellCareRx":        ("masterload", "claim"),
    "OscarRx":           ("masterload", "claim"),
    "CenteneRx":         ("masterload", "claim"),
    "CenteneFidelisRx":  ("masterload", "claim"),
    "AetnaQNXTRx":       ("masterload", "claim"),
    "AetnaQNXT":         ("masterload", "claim"),
    "WellpointEdwardRx": ("masterload", "claim"),
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
    # 2026-05-19: 'Medica 0110 Load' has started (Q=1357993 Ready). Pin Wed
    # 5/20 to L so the regular weekly cycle surfaces. Without this, the
    # 5/18 catch-up cert (the "Medica (5/1/26)" row) is in the same Mon-Fri
    # week and cert_in_week picks it up for the regular cell. Remove this
    # entry once the regular Wed cert lands.
    ("Medica",    date(2026, 5, 20)): "L",
    # 2026-05-19: 'Centene Medical 0110 Claims Load' failed (ADO 954657).
    # has_recent_failure may miss this if a stage/snap step succeeded after
    # the load failure — pin it explicitly.
    ("Centene",   date(2026, 5, 19)): "Load Failure",
    # 2026-05-19: Kaiser_NW Thu cell — user reports load failure not yet
    # surfaced automatically. Pin it explicitly.
    ("Kaiser_NW", date(2026, 5, 21)): "Load Failure",
    # 2026-05-20: CignaFacets 5/12 Tue cycle certified 5/19 (Mon, outside the
    # default Mon-Fri 5/11-5/15 backward window). Per user: "missing past
    # dates … CignaFacets on 5/12/26." Pin the late cert explicitly.
    ("CignaFacets", date(2026, 5, 12)): date(2026, 5, 19),
    # 2026-05-20: MMOHRx Weekly has a failure ticket but no queued/enabled
    # "Weekly Claim 0110 Load" run in the queue — has_recent_failure can't
    # detect it. Pin explicitly. ADO #955575 linked via LOAD_FAILURE_ADO_LINKS.
    ("MMOHRx",    date(2026, 5, 19)): "Load Failure",
}

# ADO ticket IDs to hyperlink onto specific Load-Failure cells. Keyed by
# (client, day) — same convention as MANUAL_OVERRIDES. When a cell renders
# "Load Failure" AND has an entry here, the marker text becomes a clickable
# link to the TFS work item. Per user 2026-05-20: "For Load failures that
# have an ADO, like 954657 for Centene Medical, added as a link to the
# 'Load Failure' comment."
LOAD_FAILURE_ADO_LINKS = {
    ("Centene",    date(2026, 5, 19)): 954657,  # 'Centene Medical 0110 Claims Load'
    ("ExcellusRx", date(2026, 5, 20)): 955578,  # 'Excellus - Rx - ExcellusRx 0110 Load'
    ("MMOHRx",     date(2026, 5, 19)): 955575,  # 'MMOH - Rx - MMOHRx Weekly Claim 0110 Load'
    # CignaRx removed 2026-05-20 — 'Cigna RX 0110 Load' is now Ready in the
    # queue (no longer a stage failure to surface). Re-add if/when needed.
}

# Extra rows injected into the calendar after standard placement runs. Use for
# one-off catch-up entries that don't fit the regular weekly/monthly cadence.
# Tuple: (section, day, label, marker, alert, highlight)
#   section ∈ {"daily", "weekly", "monthly", "kaiser"}
ADDITIONAL_ENTRIES = [
    # Medica catch-up for 5/1/26 claims — certified 2026-05-18 09:13:42
    # (DHT). Display cert date in the Mon cell.
    ("weekly", date(2026, 5, 18), "Medica (5/1/26)", date(2026, 5, 18), False, None),
]

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
# anchored to 5/21 (cert day) even while loading. Kaiser_AmbM is handled
# via SNAP_DISABLED_CLIENTS (load runs, snap disabled — marker "Snap").
MONTHLY_CERT_ONLY_CLIENTS = {
    "BCBSKS", "BCBSKSMedAdv", "BCBSSCRx", "CareFirstRx",
    "Kaiser_AmbCO", "Kaiser_AmbGA", "Kaiser_AmbHI",
    "Kaiser_AmbN", "Kaiser_AmbNW", "Kaiser_AmbS",
    # Kaiser_WA: per user 2026-05-18, load completion alone is not delivery —
    # the cell should stay L on the expected day until the cycle truly
    # completes. Previously in LOAD_AS_DELIVERY, which auto-✓'d on a Successful
    # load even when the actual data was empty/incomplete.
    "Kaiser_WA",
}

# Monthly clients that should show an empty Date cell (rather than "No Data")
# until the cert lands. Per-user: "M - BCBSKSMedAdv had data, so have it blank
# each month until the ticket, Snap, Certification process finish on/near 15th".
MONTHLY_BLANK_UNTIL_CERT = {"BCBSKSMedAdv"}

# Monthly clients whose "No Data" should always be shaded regardless of the
# 7-day grace window (e.g. ElixirRx hasn't received data in a long time).
FORCE_SHADE_NO_DATA = {"ElixirRx"}

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
    "Kaiser_AmbCO", "Kaiser_AmbGA", "Kaiser_AmbHI", "Kaiser_AmbM", "Kaiser_AmbN",
    "Kaiser_AmbNW", "Kaiser_AmbS",
    "Kaiser_GE",
    "Kaiser_WA", "Kaiser_WARx",
    "MedicalMutualMHS", "MedicalMutualOH", "MedImpactPBMRx",
    "MMOH", "NCState", "NCStateRx",
    "ESIPBMRx",                         # monthly snap-only (RAMP snap-driven)
    "OptumPBMRx",                       # monthly, tape-driven
    "PremeraMedAdvRx", "PremeraMedAdvVIS",
    "SamaritanHealth", "Tufts_PublicPlan", "TuftsRx",
}

# Override display name for a client (the label only; client_key stays the same).
CLIENT_DISPLAY_NAME = {
    "BCBSFLEligibilityLoad": "BCBSFL Elig",
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

# Clients whose "is delivered" signal is exclusively from TRGETL3 tape loads.
# Lookups for these clients ignore RAMP snap entries entirely.
TAPE_ONLY_CLIENTS = {"OptumPBMRx", "ESIPBMRx", "MedImpactPBMRx"}

# Snap destination filter — when a client uses a specific snap destination,
# only count snap entries matching that destination string.
SNAP_DESTINATION_FILTER = {
    "MMOH": "Pharmacy",
}

# Clients that show ✓ when snapped (not just blank when no cert).
# Combines:  daily clients (✓ on snap-only days), PBMRx clients, and the
# small set of "select" snap-deliverable clients the user named.
SNAP_ONLY_CLIENTS = {
    "OptumPBMRx",
    "ESIPBMRx", "CVSPBMRx",
    "MedImpactPBMRx", "PrimePBMRx",
    "MMOH",
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
    "AetnaSubro", "MMOH", "TuftsRx", "NCState", "WPRxDMGCOBMining",
    # Kaiser_GE needs snap-step completion (0120 Snap).
    "Kaiser_GE",
    # Kaiser ambulance feeds: per user 2026-05-15, must wait for an actual
    # snap step (Kaiser Ambulance NC/CO/GA/HI/NW/S 0120 Snap) — a load-step
    # completion alone leaves the cell in "L" (load done, snap pending).
    "Kaiser_AmbCO", "Kaiser_AmbGA", "Kaiser_AmbHI",
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

# Override the auto-derived primary key for clients whose name is a substring
# of another client's name (causing spurious substring matches in
# find_matching_jobs). Per user 2026-05-18: WellCare jobs were detected from
# WellCareRx Ready entries because 'wellcare' ⊂ 'wellcarerx'.
CLIENT_PRIMARY_KEY_OVERRIDE = {
    "WellCare": "wellcaremedical",
}

# NYShip_Rx fires four times per month — on the 1st, 8th, 16th, 24th
# (or the next Monday if that date is a weekend).
NYSHIP_DAYS = [1, 8, 16, 24]
NYSHIP_LABEL = {1: "1st", 8: "8th", 16: "16th", 24: "24th"}

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
        "SELECT DatabaseName, [Name], PCN, CertTimestamp, CurrentStatus "
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
    """Latest cert in the lookup window for this client+scheduled_day.

    Default (backward): Mon-Fri week containing scheduled_day. A cert that
    lands any weekday of the same calendar week belongs to that week's cell.

    `CERT_DIRECTION[client] = "forward"`: 7-day window starting at scheduled
    day. Used for clients like Premera where a Mon 5/11 cert is the LATE
    completion of the previous Thu 5/7 cycle.
    """
    if CERT_DIRECTION.get(client) == "forward":
        cycle_start = scheduled_day
        cycle_end   = scheduled_day + timedelta(days=6)
    else:
        cycle_start = scheduled_day - timedelta(days=scheduled_day.weekday())
        cycle_end   = cycle_start + timedelta(days=4)
    best = None
    for key in _keys_for_client(client):
        for dt, status in cert_idx.get(key, ()):
            if status != "Certified":
                continue
            d = dt.date()
            if cycle_start <= d <= cycle_end:
                if best is None or dt > best:
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
    return data[0] if data and isinstance(data[0], list) else data


def fetch_aetna_nmsp_loads(since):
    r"""Query SQLUtilAudit.cmse_new.SourceLog for Aetna NonMSP file loads
    since `since`. ✓ for "M - Aetna NMSP - MMSEA" fires once an entry exists
    for the file at \\trgdatacap2\MMSEA\Aetna\<year>\NonMSP.
    Returns list of completion datetimes.
    """
    q = (
        "SET NOCOUNT ON; "
        "SELECT CONVERT(varchar(23), ImportCompleteDate, 121) AS Done "
        "FROM [cmse_new].[dbo].[SourceLog] "
        "WHERE EntryName LIKE '%MMSEA\\Aetna\\2026\\NonMSP%' "
        f"AND ImportCompleteDate >= '{since.isoformat()}' "
        "ORDER BY ImportCompleteDate"
    )
    r = subprocess.run(
        ["sqlcmd", "-S", "SQLUtilAudit", "-d", "cmse_new", "-E", "-Q", q,
         "-W", "-s", "\t", "-h", "-1"],
        capture_output=True, text=True, check=False,
    )
    out = []
    for line in r.stdout.splitlines():
        line = line.strip()
        if not line or line.startswith("---") or "rows affected" in line:
            continue
        dt = parse_dt(line)
        if dt:
            out.append(dt)
    return out


def fetch_tape_loads(db, since):
    """Query TRGETL3.<db>.etl.Tape for recent successful loads (ProcessStatus=50).
    Returns list of dicts: {FileName, FileLoadDate (datetime)}.
    """
    q = (
        "SET NOCOUNT ON; "
        "SELECT FileName, FileLoadDate FROM [etl].[Tape] "
        f"WHERE ProcessStatus = 50 AND FileLoadDate >= '{since.isoformat()}' "
        "ORDER BY FileLoadDate"
    )
    r = subprocess.run(
        ["sqlcmd", "-S", "TRGETL3", "-d", db, "-E", "-Q", q,
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


# Map of client → (TRGETL3 database name, snap-index source key).
TAPE_LOAD_SOURCES = {
    "OptumPBMRx":      ("OptumPBMRx",      "optumpbmrx"),
    "ESIPBMRx":        ("ESIPBMRx",        "esipbmrx"),
    "MedImpactPBMRx":  ("MedImpactPBMRx",  "medimpactpbmrx"),
}

# Regex for state codes inside ESIPBMRx tape filenames (e.g. Rawlings_FL_, Rawlings_GA_)
ESIPBMRX_STATE_RE = re.compile(r"Rawlings_([A-Z]{2})_", re.I)

# Multi-week load detection: client → (TRGETL3 db, regex extracting week range
# from filename). The capture groups should be (start_yyyymmdd, end_yyyymmdd)
# or a single token uniquely identifying a week's worth of data. When recent
# loads contain >1 distinct week-key, the client label gets "(N weeks)".
MULTI_WEEK_CLIENTS = {
    "CenteneRx": ("CenteneRx", re.compile(r"_(\d{8})_(\d{8})\.txt$", re.I)),
}


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
        if not any(kw in jn_lower for kw in ("load", "stage", "snap", "mine")):
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
    load_only = client in LOAD_AS_DELIVERY_CLIENTS
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
        if any(kw in jn for kw in ("stage", "snap", "mine", "logfile", "sftp", "upload")):
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
    base = CLIENT_DISPLAY_NAME.get(client, client)
    suffix = CLIENT_SUFFIXES.get(client, "")
    label = f"{base}{suffix}{extra_suffix}"
    return f"M - {label}" if monthly else label


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
    cal = calendar.Calendar(firstweekday=0)
    out = []
    for wk in cal.monthdatescalendar(year, month):
        row = [d if d.month == month else None for d in wk[:5]]
        if any(d is not None for d in row):
            out.append(row)
    return out


def next_monday_if_weekend(d):
    """If d is Sat/Sun, return the following Monday; otherwise d."""
    if d.weekday() == 5:
        return d + timedelta(days=2)
    if d.weekday() == 6:
        return d + timedelta(days=1)
    return d


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
                  aetna_nmsp_loads=None):
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
                    # window (e.g. MedImpactPBMRx range 5-10 → shade after 5/12).
                    start_day = date(today.year, today.month,
                                     min(rng[0], calendar.monthrange(today.year, today.month)[1]))
                    return (today - start_day).days > 7
                except ValueError:
                    return False
            return False
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

    def determine_monthly(client):
        """Return (placement_date, marker) for a monthly client.
        Cert/snap dates remain on their actual date; all other markers are
        anchored to the client's expected delivery day (end of its range)."""
        # expected placement day (end of range; or avg if no range; fallback 15th)
        expected_d = expected_end_day(client)
        try:
            placeholder = date(year, month, min(expected_d, calendar.monthrange(year, month)[1]))
        except ValueError:
            placeholder = date(year, month, 15)
        placeholder = next_monday_if_weekend(placeholder)
        # 0) Forced-inactive clients always show "Inactive" on expected day
        if client in FORCED_INACTIVE:
            return placeholder, "Inactive"
        # 0b) Snap-disabled clients (load runs but snap step is disabled in RAMP)
        # show marker "Snap" with pink shading on their expected day.
        if client in SNAP_DISABLED_CLIENTS:
            return placeholder, "Snap"
        try:
            expected_date = date(year, month, min(expected_d, calendar.monthrange(year, month)[1]))
        except ValueError:
            expected_date = date(year, month, 15)
        expected_date = next_monday_if_weekend(expected_date)
        if expected_date.month != month:
            expected_date = date(year, month, calendar.monthrange(year, month)[1])

        # Kaiser_Amb* feeds cert on the Thursday following their load date.
        # Override expected_date dynamically when a load has happened this
        # month so the cell tracks the actual cycle (per user 2026-05-15).
        if client.startswith("Kaiser_Amb") and client != "Kaiser_AmbM":
            ln = load_this_month(client, snap_idx, year, month, today)
            if ln:
                load_d = ln.date()
                # First Thursday STRICTLY after load_d (weekday 3 = Thursday)
                days_until = (3 - load_d.weekday()) % 7
                if days_until == 0:
                    days_until = 7
                candidate = load_d + timedelta(days=days_until)
                if candidate.year == year and candidate.month == month:
                    expected_date = candidate

        # 1) Already certified this month → place on actual cert date
        c_latest = latest_cert(client, cert_idx, on_or_before=today)
        if c_latest and c_latest.year == year and c_latest.month == month:
            d = c_latest.date()
            d = next_monday_if_weekend(d) if d.weekday() >= 5 else d
            return d, c_latest.date()

        # Cert-only clients (BCBSKS/BCBSKSMedAdv/BCBSSCRx) stay on expected
        # day until DHT cert lands.
        if client in MONTHLY_CERT_ONLY_CLIENTS:
            # Blank-until-cert clients (BCBSKSMedAdv) ignore mid-process
            # activity entirely — they stay empty until a real cert arrives.
            if client in MONTHLY_BLANK_UNTIL_CERT:
                return expected_date, ""
            if is_loading_today(client, ramp_queue, ramp_jobs):
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
        if is_loading_today(client, ramp_queue, ramp_jobs):
            return today, "L"

        # 3) Recent failure today → today + Load Failure
        if has_recent_failure(client, ramp_queue, ramp_jobs, today):
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
        ticket = latest_tickets.get(client)
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
        if marker_override is None:
            mov = MANUAL_OVERRIDES.get((client, day))
            if mov is not None:
                marker_override = mov
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
            place(daily, "daily", c, d)

    # KaiserPrePayCOB at the very bottom of each week (one row, all 5 columns)
    for d in all_days:
        place(kaiser, "kaiser", KAISER_PREPAY_CLIENT, d)

    # weekly clients on assigned weekday (alphabetical within column)
    for c in sorted(WEEKLY_CLIENTS):
        days = WEEKLY_CLIENTS[c]
        for d in all_days:
            if d.strftime("%A") in days:
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
        # Cert-style client — stay L when loaded this week until cert lands.
        marker = resolve_marker("NYShip_Rx", tgt, allow_checkmark=False, allow_week_window=True)
        alert  = alert_state("NYShip_Rx", tgt, marker)
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
    def place_optum_half(label_suffix, day_lo, day_hi):
        last_day = calendar.monthrange(year, month)[1]
        hi = min(day_hi, last_day)
        window_start = date(year, month, day_lo)
        window_end   = date(year, month, hi)
        # find any OptumPBMRx tape load in this window
        wanted = TAPE_LOAD_SOURCES["OptumPBMRx"][1]
        latest = None
        for d in sorted(snap_idx.keys()):
            if not (window_start <= d <= window_end):
                continue
            for entry in snap_idx[d]:
                if entry[0] == wanted and (latest is None or entry[1] > latest):
                    latest = entry[1]
        if latest:
            placement = latest.date()
            # If the latest tape load lands on a weekend, push to next Monday
            placement = next_monday_if_weekend(placement)
            if placement.month != month:
                placement = date(year, month, calendar.monthrange(year, month)[1])
            marker = "✓"
        else:
            # No tape this half yet — place on mid-window, marker depends on grace
            mid = day_lo + (hi - day_lo) // 2
            placement = next_monday_if_weekend(date(year, month, mid))
            marker = "No Data"
        label = f"M - OptumPBMRx {label_suffix}"
        # alert if "No Data" and past 7 days past END
        alert = False
        if marker == "No Data":
            ref_end = date(year, month, min(hi, last_day))
            alert = (today - ref_end).days > 7
        monthly[placement].append((label, marker, alert, None))

    place_optum_half("(RAW 1/2/3)", 1, 7)
    place_optum_half("(RAW 5/6)",   24, 31)

    # Aetna NMSP - MMSEA: ✓ once SourceLog shows a NonMSP file imported this
    # month; otherwise the 15th rule (or next Monday) with No Data.
    nmsp_day = nmsp_mmsea_date(year, month)
    if nmsp_day.month == month:
        nmsp_dt = None
        for dt in (aetna_nmsp_loads or ()):
            if dt.year == year and dt.month == month and dt.date() <= today:
                if nmsp_dt is None or dt > nmsp_dt:
                    nmsp_dt = dt
        if nmsp_dt:
            placement = nmsp_dt.date()
            if placement.weekday() >= 5:
                placement = next_monday_if_weekend(placement)
            if placement.month != month:
                placement = nmsp_day
            marker = "✓"
            alert  = False
        else:
            placement = nmsp_day
            marker = "No Data"
            alert = alert_state("AetnaMMSEA", nmsp_day, marker)
        monthly[placement].append(("M - Aetna NMSP - MMSEA", marker, alert, "bold"))

    # (Removed) loading-today extras pass — L is now surfaced on each client's
    # scheduled weekday cell via resolve_marker / determine_monthly directly.

    # One-off injected entries (catch-up loads, etc.)
    section_map = {"daily": daily, "weekly": weekly, "monthly": monthly, "kaiser": kaiser}
    for section, day, label, marker, alert, highlight in ADDITIONAL_ENTRIES:
        bucket = section_map.get(section)
        if bucket is None or day.month != month or day.year != year:
            continue
        bucket[day].append((label, marker, alert, highlight))

    # KaiserPrePayCOB weekend handling: Saturday loads surface on the prior
    # Friday with a " (Sat)" suffix; Sunday loads surface on the next Monday
    # with a " (Sun)" suffix. Per user 2026-05-18.
    kpp_keys = {k for k in _keys_for_client(KAISER_PREPAY_CLIENT)}
    kpp_seen = set()   # dedupe by (target_day, suffix)
    for d in sorted(snap_idx.keys()):
        if d.year != year or d.month != month:
            continue
        if d.weekday() == 5:        # Saturday → previous Friday
            target = d - timedelta(days=1)
            suffix = " (Sat)"
        elif d.weekday() == 6:      # Sunday → next Monday
            target = d + timedelta(days=1)
            suffix = " (Sun)"
        else:
            continue
        if target.month != month or target.year != year:
            continue
        for entry in snap_idx.get(d, ()):
            src = entry[0]
            if src not in kpp_keys:
                continue
            key = (target, suffix)
            if key in kpp_seen:
                break
            kpp_seen.add(key)
            label = f"{display_name(KAISER_PREPAY_CLIENT)}{suffix}"
            kaiser[target].append((label, "✓", False, None))
            break

    # Sort each cell alphabetically within each section
    for bucket in (daily, weekly, monthly, kaiser):
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


def write_weekly_stacked(ws, year, month, sections, weeks, today):
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

        # date strip row (real Excel dates, formatted mm/dd/yy)
        for i, d in enumerate(wk):
            col_dat = i * 2 + 2
            if d:
                cell = ws.cell(row=cur_row, column=col_dat, value=d)
                cell.number_format = "mm/dd/yy"
                cell.fill = DAY_FILL if d != today else TODAY_FILL
                cell.font = DAY_FONT
                cell.alignment = Alignment(horizontal="center")
                cell.border = BORDER
        cur_row += 1

        # Daily → blank → Weekly → blank → Monthly → blank → KaiserPrePayCOB
        cur_row = _write_section_rows(ws, cur_row, wk, sections["daily"])
        cur_row = _blank_separator_row(ws, cur_row)
        cur_row = _write_section_rows(ws, cur_row, wk, sections["weekly"])
        cur_row = _blank_separator_row(ws, cur_row)
        cur_row = _write_section_rows(ws, cur_row, wk, sections["monthly"])
        cur_row = _blank_separator_row(ws, cur_row)
        cur_row = _write_section_rows(ws, cur_row, wk, sections["kaiser"])

        # per-week key block
        ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=10)
        kc = ws.cell(row=cur_row, column=1,
                     value="Key:  Date = Certified  |  ✓ = Loaded/Snapped  |  L = Loading"
                           "  |  pink = Failure/Inactive  |  (s) SLA  |  (p) Rx Post Snap"
                           "  |  (n) Not Delivered  |  M - Monthly")
        kc.font = Font(name="Segoe UI", italic=True, size=9, color="555555")
        kc.alignment = Alignment(horizontal="left")
        cur_row += 2

    # Client-name columns set to ≈190 px (width 27.07) per user 2026-05-20.
    # Excel pixels ≈ 7 * width + 0.5.
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
                           "  |  (n) Not Delivered  |  M - Monthly")
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
def main():
    today = date.today()
    year, month = today.year, today.month
    month_start = date(year, month, 1)

    print(f"[info] Today: {today}  Month: {year}-{month:02d}")

    print("[info] Querying DHT cert table…")
    # pull 6 months back so monthly clients have enough history for avg-day calc
    certs = fetch_dht_certs(since=date(year, month, 1) - timedelta(days=185))
    cert_idx = build_cert_index(certs)
    print(f"[info]   {len(certs)} cert rows / {len(cert_idx)} distinct DatabaseNames")

    print("[info] Fetching ADO tickets…")
    tickets = fetch_ado_tickets(min_changed_date=month_start - timedelta(days=14))
    print(f"[info]   {len(tickets)} delivery tickets in window")

    print("[info] Fetching RAMP jobs…")
    jobs = fetch_ramp_jobs()
    enabled_n = sum(1 for j in jobs if j.get("Enabled") == 1)
    print(f"[info]   {len(jobs)} total ({enabled_n} enabled)")

    print("[info] Fetching RAMP queue + snap history…")
    queue = fetch_ramp_queue()
    snaps = fetch_ramp_snaps()
    print(f"[info]   queue={len(queue)}, snaps={len(snaps)}")

    print("[info] Fetching TRGETL3 PBMRx tape loads…")
    since_dt = date(year, month, 1) - timedelta(days=14)
    tape_loads = {}
    for client, (db, src_key) in TAPE_LOAD_SOURCES.items():
        rows = fetch_tape_loads(db, since_dt)
        tape_loads[src_key] = rows
        print(f"[info]   {db}: {len(rows)} rows")

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

    sections, weeks = plan_calendar(year, month, cert_idx, snap_idx,
                                    latest_tickets, monthly_placements, jobs, queue,
                                    esipbmrx_tape=tape_loads.get("esipbmrx"),
                                    multi_week_loads=multi_week_loads,
                                    aetna_nmsp_loads=aetna_nmsp_loads)

    wb = Workbook()
    ws_weekly = wb.active
    ws_weekly.title = "Weekly Stacked"
    write_weekly_stacked(ws_weekly, year, month, sections, weeks, today)

    ws_owner = wb.create_sheet("Client Owner")
    write_client_owner_sheet(ws_owner)

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

    return out_path


if __name__ == "__main__":
    main()
