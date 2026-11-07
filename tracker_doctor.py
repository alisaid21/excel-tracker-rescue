#!/usr/bin/env python3
"""
tracker_doctor.py - diagnose why an Excel workbook is slow or oversized.

Reads the .xlsx file directly (no Excel needed, nothing is modified) and
reports the usual causes of tracker bloat:

  * used range far larger than the real data (phantom formatted rows)
  * formulas copied far past the last row of data
  * volatile functions (TODAY, NOW, OFFSET, INDIRECT, RAND, ...) that force
    recalculation on every edit
  * whole-column references (A:A) inside formulas repeated on every row
  * conditional formatting that is fragmented, duplicated or whole-column
  * an exploded number of cell formats / styles

Usage:
    python tracker_doctor.py MyTracker.xlsx
    python tracker_doctor.py MyTracker.xlsx --json report.json

Requires only the Python 3 standard library.
"""

import argparse
import json
import os
import posixpath
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RNS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PKG = "{http://schemas.openxmlformats.org/package/2006/relationships}"

VOLATILE = re.compile(r"\b(NOW|TODAY|RAND|RANDBETWEEN|RANDARRAY|OFFSET|INDIRECT|CELL|INFO)\s*\(", re.I)
WHOLE_COL = re.compile(r"""(?<![A-Za-z0-9_.$])\$?[A-Z]{1,3}:\$?[A-Z]{1,3}(?![A-Za-z0-9_(])""")
CELL_REF = re.compile(r"([A-Z]{1,3})(\d+)")
EXCEL_MAX_ROWS = 1_048_576


def row_of(ref):
    m = CELL_REF.match(ref or "")
    return int(m.group(2)) if m else 0


def range_rows(sqref):
    """Total rows covered by a space-separated list of ranges (A1:B9 C3)."""
    total = 0
    for part in sqref.split():
        if ":" in part:
            a, b = part.split(":")
            ra, rb = row_of(a), row_of(b)
            if ra == 0 or rb == 0:          # A:A style whole column
                total += EXCEL_MAX_ROWS
            else:
                total += rb - ra + 1
        else:
            total += 1
    return total


def sheet_paths(z):
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    target = {r.get("Id"): r.get("Target") for r in rels.iter(PKG + "Relationship")}
    out = []
    for s in wb.iter(NS + "sheet"):
        t = target[s.get(RNS + "id")]
        path = t.lstrip("/") if t.startswith("/") else posixpath.normpath(posixpath.join("xl", t))
        out.append((s.get("name"), path))
    names = wb.find(NS + "definedNames")
    return out, (len(list(names)) if names is not None else 0)


def style_stats(z):
    if "xl/styles.xml" not in z.namelist():
        return {}
    st = ET.fromstring(z.read("xl/styles.xml"))

    def count(tag):
        el = st.find(NS + tag)
        return len(list(el)) if el is not None else 0

    return {"cell_formats": count("cellXfs"), "named_style_formats": count("cellStyleXfs"),
            "named_styles": count("cellStyles"), "fonts": count("fonts"), "fills": count("fills"),
            "borders": count("borders"), "cf_formats": count("dxfs")}


def scan_sheet(z, path):
    s = dict(dimension="", cells=0, value_cells=0, last_value_row=0, last_formula_row=0,
             last_cell_row=0, formulas=0, volatile_formulas=0, whole_column_formulas=0,
             cf_blocks=0, cf_rules=0, cf_rules_whole_column=0, cf_rows_covered=0, tables=0)
    shared = {}  # si -> (volatile, whole_col)
    with z.open(path) as fh:
        for _, el in ET.iterparse(fh, events=("end",)):
            tag = el.tag
            if tag == NS + "c":
                s["cells"] += 1
                r = row_of(el.get("r"))
                s["last_cell_row"] = max(s["last_cell_row"], r)
                f = el.find(NS + "f")
                has_value = el.find(NS + "v") is not None or el.find(NS + "is") is not None
                if f is not None:
                    s["formulas"] += 1
                    s["last_formula_row"] = max(s["last_formula_row"], r)
                    text = f.text or ""
                    if f.get("t") == "shared":
                        si = f.get("si")
                        if text:
                            shared[si] = (bool(VOLATILE.search(text)), bool(WHOLE_COL.search(text)))
                        vol, whole = shared.get(si, (False, False))
                    else:
                        vol, whole = bool(VOLATILE.search(text)), bool(WHOLE_COL.search(text))
                    s["volatile_formulas"] += vol
                    s["whole_column_formulas"] += whole
                elif has_value:
                    s["value_cells"] += 1
                    s["last_value_row"] = max(s["last_value_row"], r)
                el.clear()
            elif tag == NS + "row":
                el.clear()
            elif tag == NS + "dimension":
                s["dimension"] = el.get("ref", "")
            elif tag == NS + "conditionalFormatting":
                s["cf_blocks"] += 1
                rules = el.findall(NS + "cfRule")
                rows = range_rows(el.get("sqref", ""))
                s["cf_rules"] += len(rules)
                s["cf_rows_covered"] += rows * len(rules)
                if rows >= EXCEL_MAX_ROWS:
                    s["cf_rules_whole_column"] += len(rules)
            elif tag == NS + "tablePart":
                s["tables"] += 1
    return s


def findings(sheet_name, s):
    out = []
    data_rows = s["last_value_row"]
    used = max(s["last_cell_row"], row_of(s["dimension"].split(":")[-1]))
    if data_rows and used > data_rows * 1.5 and used - data_rows > 1000:
        out.append(("HIGH", f"Used range runs to row {used:,} but data stops at row {data_rows:,}.",
                    "Delete the empty formatted rows below the data, then save so Excel resets the used range."))
    if data_rows and s["last_formula_row"] > data_rows + 100:
        out.append(("HIGH", f"Formulas continue to row {s['last_formula_row']:,}, "
                            f"{s['last_formula_row'] - data_rows:,} rows past the data.",
                    "Convert the data to an Excel Table so formulas fill down only as rows are added."))
    if s["volatile_formulas"] > 50:
        out.append(("HIGH", f"{s['volatile_formulas']:,} formulas use volatile functions "
                            "(TODAY, NOW, OFFSET, INDIRECT...), so they recalculate on every edit.",
                    "Put TODAY() in one settings cell and reference it; replace OFFSET/INDIRECT with direct references."))
    if s["whole_column_formulas"] > 50:
        out.append(("HIGH", f"{s['whole_column_formulas']:,} formulas scan whole columns (A:A).",
                    "Calculate per-site or per-person counts once in a summary sheet, over the table's columns only."))
    if s["cf_rules"] > 20:
        out.append(("MEDIUM", f"{s['cf_rules']:,} conditional formatting rules in {s['cf_blocks']:,} separate ranges.",
                    "Delete them all and recreate a handful of rules scoped to the data range."))
    if s["cf_rules_whole_column"]:
        out.append(("MEDIUM", f"{s['cf_rules_whole_column']} conditional formatting rules cover entire columns.",
                    "Limit each rule to the rows that actually hold data."))
    return [(sheet_name, *f) for f in out]


def diagnose(path):
    report = {"file": os.path.basename(path), "size_bytes": os.path.getsize(path), "sheets": {}, "findings": []}
    with zipfile.ZipFile(path) as z:
        sheets, n_names = sheet_paths(z)
        report["defined_names"] = n_names
        report["styles"] = style_stats(z)
        for name, p in sheets:
            s = scan_sheet(z, p)
            report["sheets"][name] = s
            report["findings"] += findings(name, s)
    cf = report["styles"].get("cell_formats", 0)
    if cf > 4000:
        report["findings"].append(("Workbook", "HIGH", f"{cf:,} distinct cell formats (Excel's limit is about 64,000).",
                                   "Clear formats on the data and reapply a small, consistent set of styles."))
    elif cf > 500:
        report["findings"].append(("Workbook", "MEDIUM", f"{cf:,} distinct cell formats.",
                                   "Standardise fonts, fills and date formats across the sheet."))
    return report


def print_report(r):
    mb = r["size_bytes"] / 1_048_576
    print(f"\n=== tracker_doctor: {r['file']} ===")
    print(f"File size          {mb:,.2f} MB")
    st = r["styles"]
    if st:
        print(f"Cell formats       {st['cell_formats']:,}   (fonts {st['fonts']:,}, fills {st['fills']:,}, "
              f"CF formats {st['cf_formats']:,})")
    for name, s in r["sheets"].items():
        print(f"\n[{name}]")
        print(f"  Data ends at row        {s['last_value_row']:,}")
        print(f"  Used range ends at row  {max(s['last_cell_row'], row_of(s['dimension'].split(':')[-1])):,}")
        print(f"  Cells stored            {s['cells']:,}")
        print(f"  Formulas                {s['formulas']:,}  (volatile {s['volatile_formulas']:,}, "
              f"whole-column {s['whole_column_formulas']:,})")
        print(f"  Conditional formatting  {s['cf_rules']:,} rules in {s['cf_blocks']:,} ranges "
              f"({s['cf_rules_whole_column']} whole-column)")
        print(f"  Excel Tables            {s['tables']}")
    print("\nFindings")
    if not r["findings"]:
        print("  None. This workbook is in good shape.")
    for sheet, sev, what, fix in r["findings"]:
        print(f"  [{sev}] {sheet}: {what}\n         Fix: {fix}")
    print()


def main(argv=None):
    ap = argparse.ArgumentParser(description="Diagnose why an Excel workbook is slow or oversized.")
    ap.add_argument("xlsx", help="path to the .xlsx or .xlsm file")
    ap.add_argument("--json", metavar="PATH", help="also write the full report as JSON")
    a = ap.parse_args(argv)
    if not zipfile.is_zipfile(a.xlsx):
        sys.exit("Not an .xlsx/.xlsm file (older .xls files must be re-saved as .xlsx first).")
    r = diagnose(a.xlsx)
    print_report(r)
    if a.json:
        with open(a.json, "w") as fh:
            json.dump(r, fh, indent=2)
    return r


if __name__ == "__main__":
    main()
