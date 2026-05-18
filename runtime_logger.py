#!/usr/bin/env python3
import re
import json
import os

LOG_FILE = "outputs/runtime_log.json"

PATTERNS = {
    "Parameter Reading": r"Parameter Reading\s*:\s*User=([0-9.]+)s\s*Real=([0-9.]+)s\s*Sys=([0-9.]+)s",
    "Initialization": r"Initialization:\s*User=([0-9.]+)s\s*Real=([0-9.]+)s\s*Sys=([0-9.]+)s",
    "Execution": r"Execution:\s*User=([0-9.]+)s\s*Real=([0-9.]+)s\s*Sys=([0-9.]+)s",
    "Finalization": r"Finalization:\s*User=([0-9.]+)s\s*Real=([0-9.]+)s\s*Sys=([0-9.]+)s",
    "Total": r"Total:\s*User=([0-9.]+)s\s*Real=([0-9.]+)s\s*Sys=([0-9.]+)s",
}


def parse_runtime(log_text):
    result = {}

    for section, pattern in PATTERNS.items():
        m = re.search(pattern, log_text)

        if m:
            result[section] = {
                "User": float(m.group(1)),
                "Real": float(m.group(2)),
                "Sys": float(m.group(3)),
            }
        else:
            result[section] = {
                "User": None,
                "Real": None,
                "Sys": None,
            }

    return result


def append_runtime(run_number, runtime_data):
    os.makedirs("outputs", exist_ok=True)

    if os.path.exists(LOG_FILE):
        with open(LOG_FILE, "r") as f:
            data = json.load(f)
    else:
        data = []

    data.append({
        "run": run_number,
        "runtime": runtime_data
    })

    with open(LOG_FILE, "w") as f:
        json.dump(data, f, indent=2)


if __name__ == "__main__":
    import sys

    if len(sys.argv) != 3:
        print("Usage: python runtime_logger.py <run_number> <log_file>")
        sys.exit(1)

    run_number = int(sys.argv[1])
    logfile = sys.argv[2]

    with open(logfile, "r") as f:
        text = f.read()

    parsed = parse_runtime(text)
    append_runtime(run_number, parsed)