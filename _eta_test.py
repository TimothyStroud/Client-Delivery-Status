import ramp_aetnahrp_status_digest as hrp
import ramp_aetnarx_status_digest as rx

for tag, m, srv, job in (
    ("HRP", hrp, "TRGETL2", "ETL AetnaHRP MasterLoad"),
    ("RX",  rx,  "TRGETL2", "ETL AetnaRx MasterLoad Claims And Eligibility"),
):
    print(f"=== {tag} ===")
    print("  milestone:", m._ssis_milestone())
    print("  final_stage_eta (None expected, nothing running):", m._ssis_final_stage_eta())
    print("  recent full durations (s):", sorted(m._recent_full_durations(srv, job)))
    print("  current_run_start (None expected):", m._current_run_start(srv, job))
    print("  eta_detail line:", m.eta_detail(srv, job))
