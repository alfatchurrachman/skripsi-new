#!/usr/bin/env python3
import os
import glob
import sys
import numpy as np

DOSE_LIMIT   = 0.5
OUTPUT_DIR   = "outputs"
PHSP_PATTERN = "DNADamage_*.phsp"
COL_DOSE     = 1
COL_DSB      = 20

def parse_phsp(filepath):
    total_dsb   = 0.0
    n_histories = 0
    with open(filepath, "r") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            parts = line.split()
            if len(parts) <= max(COL_DOSE, COL_DSB):
                continue
            try:
                dose_cumulative = float(parts[COL_DOSE])
                dsb_value       = float(parts[COL_DSB])
            except ValueError:
                continue
            if dose_cumulative >= DOSE_LIMIT:
                break
            total_dsb   += dsb_value
            n_histories += 1
    return total_dsb, n_histories

def main():
    pattern = os.path.join(OUTPUT_DIR, PHSP_PATTERN)
    files   = sorted(glob.glob(pattern))

    if not files:
        print("ERROR: Tidak ada file .phsp ditemukan.", file=sys.stderr)
        sys.exit(1)

    dsb_per_run = []
    for filepath in files:
        total_dsb, n_hist = parse_phsp(filepath)
        dsb_per_run.append(total_dsb)
        print(
            f"  {os.path.basename(filepath)}: "
            f"total SB = {total_dsb:.4f} (dari {n_hist} histori)",
            file=sys.stderr
        )

    if len(dsb_per_run) < 2:
        print("inf")
        return

    arr       = np.array(dsb_per_run)
    mean      = arr.mean()
    std       = arr.std(ddof=1)
    rel_error = std / mean if mean != 0 else float("inf")

    # ==========================================================
    # Plateau detection
    # ==========================================================
    delta_dsb = None

    if len(arr) >= 2:
        prev = arr[-2]
        curr = arr[-1]

        if prev != 0:
            delta_dsb = abs(curr - prev) / abs(prev)
        else:
            delta_dsb = float("inf")

    print(
        f"  Across {len(arr)} runs: mean = {mean:.4f}, "
        f"std = {std:.4f}, relative error = {rel_error:.6f}",
        file=sys.stderr
    )

    if delta_dsb is not None:
        print(
            f"  Relative DSB change (last run vs previous) = {delta_dsb:.6e}",
            file=sys.stderr
        )

    # Output format:
    # rel_error delta_dsb
    if delta_dsb is None:
        print(f"{rel_error:.6f} inf")
    else:
        print(f"{rel_error:.6f} {delta_dsb:.10f}")

if __name__ == "__main__":
    main()
