import rdp_client_delivery_status as r
jobs = r.fetch_ramp_jobs()
queue = r.fetch_ramp_queue()
snaps = r.fetch_ramp_snaps()
snap_idx = r.build_snap_index(jobs, queue, snaps, tape_loads={})
for d_, entries in sorted(snap_idx.items()):
    for e in entries:
        if e and e[0] == "jhhcpassfileemail":
            print(d_, "->", e[1], "| kind=", e[3], "| jn=", e[4])
