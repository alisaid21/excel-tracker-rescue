"""
Measure the before/after workbooks and write reports/results.json.

Timings:
  * openpyxl_load_s  - time for Python's openpyxl to load the workbook (median)
  * libreoffice_recalc_s - LibreOffice (headless) opening the file, recalculating
    every formula and saving a copy. Run with --libreoffice; it needs the
    LibreOffice recalc helper path via --recalc-script.

These are not Excel timings. Excel is generally faster than LibreOffice at
whole-column lookups, so record your own Excel timings for the case study
(see LOOM_SCRIPT.md) and pass them to make_cover.py.
"""

import argparse
import json
import shutil
import statistics
import subprocess
import sys
import tempfile
import time
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import tracker_doctor  # noqa: E402

FILES = {"before": ROOT / "before" / "WO_Tracker_bloated.xlsx",
         "after": ROOT / "after" / "WO_Tracker_fixed.xlsx"}


def time_openpyxl(path, runs=3):
    out = []
    for _ in range(runs):
        t = time.perf_counter()
        openpyxl.load_workbook(path)
        out.append(time.perf_counter() - t)
    return round(statistics.median(out), 2)


def time_libreoffice(path, script, runs):
    out = []
    for _ in range(runs):
        with tempfile.TemporaryDirectory() as d:
            tmp = Path(d) / path.name
            shutil.copy(path, tmp)
            t = time.perf_counter()
            subprocess.run([sys.executable, script, str(tmp), "1200"], check=True, capture_output=True)
            out.append(time.perf_counter() - t)
    return round(statistics.median(out), 1), runs


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--libreoffice", action="store_true")
    ap.add_argument("--recalc-script", default="/mnt/skills/public/xlsx/scripts/recalc.py")
    ap.add_argument("--only", choices=list(FILES))
    a = ap.parse_args()
    out_path = ROOT / "reports" / "results.json"
    results = json.loads(out_path.read_text()) if out_path.exists() else {}
    for key, path in FILES.items():
        if a.only and key != a.only:
            continue
        d = tracker_doctor.diagnose(str(path))
        sheet = next(iter(d["sheets"].values()))
        r = results.get(key, {})
        r.update({
            "file": path.name,
            "size_mb": round(d["size_bytes"] / 1_048_576, 2),
            "cell_formats": d["styles"]["cell_formats"],
            "used_range_last_row": max(s["last_cell_row"] for s in d["sheets"].values()),
            "cells_stored": sum(s["cells"] for s in d["sheets"].values()),
            "formulas": sum(s["formulas"] for s in d["sheets"].values()),
            "volatile_formulas": sum(s["volatile_formulas"] for s in d["sheets"].values()),
            "whole_column_formulas": sum(s["whole_column_formulas"] for s in d["sheets"].values()),
            "cf_rules": sum(s["cf_rules"] for s in d["sheets"].values()),
            "findings": len(d["findings"]),
            "data_rows": sheet["last_value_row"] - 1,
            "openpyxl_load_s": time_openpyxl(path),
        })
        if a.libreoffice:
            secs, runs = time_libreoffice(path, a.recalc_script, runs=1 if key == "before" else 3)
            r["libreoffice_recalc_s"], r["libreoffice_runs"] = secs, runs
        results[key] = r
        print(key, json.dumps(r))
    out_path.write_text(json.dumps(results, indent=2))
