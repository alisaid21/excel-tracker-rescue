"""
Build the two demo workbooks for the "Bloated Excel tracker, repaired" case study.

Both workbooks hold the SAME 3,000 synthetic work orders (seeded, reproducible).
No real company, site, asset or person appears anywhere in the data.

  before/WO_Tracker_bloated.xlsx  - the tracker as it typically ends up after
                                     a few years of copy-paste and quick fixes
  after/WO_Tracker_fixed.xlsx     - the same tracker, rebuilt properly

Usage:
    python demo/build_workbooks.py            # writes both files
"""

import datetime as dt
import random
from pathlib import Path

import xlsxwriter

ROOT = Path(__file__).resolve().parents[1]
SEED = 2026
N_ROWS = 3000
REPORT_DATE = dt.date(2026, 10, 24)   # data is generated relative to this date

# ---------------------------------------------------------------- synthetic data
TRADES = {
    "HVAC": ["AHU", "FCU", "CHL", "SPL"],
    "Electrical": ["DB", "GEN", "LTG", "UPS"],
    "Plumbing": ["PMP", "WH", "VLV", "TNK"],
    "Carpentry": ["DR", "WIN", "FRN", "CEI"],
}
TASKS = {
    "HVAC": ["No cooling in room", "Unit tripping on high pressure", "Noisy fan motor",
             "Water leak from indoor unit", "Replace air filter", "Thermostat not responding"],
    "Electrical": ["Lights not working", "Breaker tripping", "Socket damaged",
                   "Generator failed to start", "UPS alarm active", "Replace burnt fitting"],
    "Plumbing": ["Blocked drain", "Leaking tap", "No hot water", "Pump not running",
                 "Toilet flush faulty", "Pipe leak in ceiling"],
    "Carpentry": ["Door lock broken", "Window not closing", "Damaged ceiling tile",
                  "Repair office desk", "Door closer faulty", "Replace cabinet hinge"],
}
PRIORITY = [("Emergency", 1, 0.08), ("Urgent", 3, 0.27), ("Routine", 14, 0.65)]
HEADERS = ["WO Number", "Date Raised", "Site", "Asset Tag", "Trade", "Priority",
           "Description", "Assigned Tech", "Target Date", "Status", "Date Completed"]


def make_rows():
    rnd = random.Random(SEED)
    rows = []
    start = REPORT_DATE - dt.timedelta(days=365)
    for i in range(1, N_ROWS + 1):
        trade = rnd.choice(list(TRADES))
        prio, sla_days, _ = rnd.choices(PRIORITY, weights=[p[2] for p in PRIORITY])[0]
        raised = start + dt.timedelta(days=int(365 * (i / N_ROWS) ** 0.8) + rnd.randint(-3, 3))
        raised = min(raised, REPORT_DATE)
        target = raised + dt.timedelta(days=sla_days)
        age = (REPORT_DATE - raised).days
        # older jobs are more likely to be finished
        p_done = 0.97 if age > 60 else 0.85 if age > 14 else 0.45
        if rnd.random() < p_done:
            late = rnd.random() < 0.12
            dur = rnd.randint(sla_days + 1, sla_days * 3 + 2) if late else rnd.randint(0, sla_days)
            completed = min(raised + dt.timedelta(days=dur), REPORT_DATE)
            status = "Closed" if age > 30 and rnd.random() < 0.7 else "Completed"
        else:
            completed = None
            status = rnd.choices(["Open", "In Progress", "On Hold"], weights=[5, 4, 1])[0]
        rows.append([
            f"WO-26-{i:05d}",
            raised,
            f"Site {rnd.randint(1, 70):02d}",
            f"{rnd.choice(TRADES[trade])}-{rnd.randint(1, 400):04d}",
            trade,
            prio,
            rnd.choice(TASKS[trade]),
            f"TECH-{rnd.randint(1, 200):03d}",
            target,
            status,
            completed,
        ])
    return rows


# ------------------------------------------------------------- bloated tracker
def build_bloated(rows, path):
    """Reproduce the four usual causes of tracker bloat, at realistic scale."""
    rnd = random.Random(SEED + 1)
    wb = xlsxwriter.Workbook(str(path), {"constant_memory": True})
    ws = wb.add_worksheet("WO TRACKER")

    head = wb.add_format({"bold": True, "bg_color": "#1F4E78", "font_color": "#FFFFFF",
                          "border": 1, "font_name": "Arial"})
    headers = HEADERS + ["Days Open", "SLA Status", "Age Bucket",
                         "Site Open Jobs", "Tech Open Jobs", "Month"]
    ws.write_row(0, 0, headers, head)

    # Cause 1: per-cell style explosion. Formatting pasted in from many sources
    # leaves thousands of near-identical but distinct cell formats.
    palette = [f"#{r:02X}{g:02X}{b:02X}" for r in range(200, 256, 3)
               for g in range(200, 256, 7) for b in range(200, 256, 9)]
    fmt_cache = {}

    def messy_format(kind):
        key = (kind, rnd.choice(palette), rnd.choice([9, 9.5, 10, 10.5, 11]),
               rnd.choice(["Arial", "Calibri", "Tahoma"]), rnd.random() < 0.15)
        if key not in fmt_cache:
            props = {"bg_color": key[1], "font_size": key[2], "font_name": key[3],
                     "bold": key[4], "border": 1}
            if kind == "date":
                props["num_format"] = rnd.choice(["dd/mm/yyyy", "d-mmm-yy", "dd-mm-yyyy"])
            fmt_cache[key] = wb.add_format(props)
        return fmt_cache[key]

    last_formula_row = 10000   # Cause 2a: formulas dragged far past the data
    for r in range(1, last_formula_row + 1):
        x = r + 1  # Excel row number
        if r <= len(rows):
            row = rows[r - 1]
            for c, v in enumerate(row):
                if isinstance(v, dt.date):
                    ws.write_datetime(r, c, dt.datetime.combine(v, dt.time()), messy_format("date"))
                elif v is None:
                    ws.write_blank(r, c, None, messy_format("date"))
                else:
                    ws.write(r, c, v, messy_format("text"))
        # Cause 3: volatile functions in every row (OFFSET, INDIRECT, TODAY)
        # Cause 4: whole-column COUNTIFS repeated in every row
        f = messy_format("text")
        ws.write_formula(r, 11, f'=IF(B{x}="","",IF(OFFSET(K{x},0,0)="",TODAY()-OFFSET(B{x},0,0),'
                                f'OFFSET(K{x},0,0)-OFFSET(B{x},0,0)))', f)
        ws.write_formula(r, 12, f'=IF(B{x}="","",IF(INDIRECT("K"&ROW())="",'
                                f'IF(TODAY()>INDIRECT("I"&ROW()),"Overdue","On track"),'
                                f'IF(INDIRECT("K"&ROW())>INDIRECT("I"&ROW()),"Late","Met")))', f)
        ws.write_formula(r, 13, f'=IF(L{x}="","",IF(L{x}<=7,"0-7 days",IF(L{x}<=30,"8-30 days",'
                                f'IF(L{x}<=90,"31-90 days","90+ days"))))', f)
        ws.write_formula(r, 14, f'=IF(C{x}="","",COUNTIFS(INDIRECT("C:C"),C{x},INDIRECT("J:J"),"Open")'
                                f'+COUNTIFS(INDIRECT("C:C"),C{x},INDIRECT("J:J"),"In Progress")'
                                f'+COUNTIFS(INDIRECT("C:C"),C{x},INDIRECT("J:J"),"On Hold"))', f)
        ws.write_formula(r, 15, f'=IF(H{x}="","",COUNTIFS(H:H,H{x},J:J,"Open")'
                                f'+COUNTIFS(H:H,H{x},J:J,"In Progress")'
                                f'+COUNTIFS(H:H,H{x},J:J,"On Hold"))', f)
        ws.write_formula(r, 16, f'=IF(B{x}="","",TEXT(B{x},"yyyy-mm"))', f)

    # Cause 2b: phantom formatted rows - borders and fills applied to whole
    # blocks "for future data", which pushes the used range far past the data.
    phantom = wb.add_format({"border": 1, "bg_color": "#FFFFFF", "font_name": "Calibri"})
    for r in range(last_formula_row + 1, 60001):
        for c in range(len(headers)):
            ws.write_blank(r, c, None, phantom)

    # Cause 5: fragmented, duplicated conditional formatting. Every copy-paste
    # of a formatted block splits the rule into another copy.
    red = wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})
    amber = wb.add_format({"bg_color": "#FFEB9C"})
    grey = wb.add_format({"font_color": "#808080"})
    for start in range(2, len(rows) + 2, 8):
        rng = f"A{start}:Q{start + 7}"
        ws.conditional_format(rng, {"type": "formula", "criteria": f'=$M{start}="Overdue"', "format": red})
        ws.conditional_format(rng, {"type": "formula", "criteria": f'=$F{start}="Emergency"', "format": amber})
    # ...plus the same rules again on entire columns
    ws.conditional_format("A1:Q1048576", {"type": "formula", "criteria": '=$M1="Overdue"', "format": red})
    ws.conditional_format("A1:Q1048576", {"type": "formula", "criteria": '=$J1="Closed"', "format": grey})
    ws.conditional_format("L1:L1048576", {"type": "cell", "criteria": ">", "value": 30, "format": amber})

    ws.freeze_panes(1, 0)
    ws.set_column(0, 16, 14)
    wb.close()
    return len(fmt_cache)


# --------------------------------------------------------------- fixed tracker
def build_fixed(rows, path):
    wb = xlsxwriter.Workbook(str(path))
    base = {"font_name": "Arial", "font_size": 10}
    date_f = wb.add_format({**base, "num_format": "dd-mmm-yyyy"})
    int_f = wb.add_format({**base, "num_format": "0"})
    text_f = wb.add_format(base)
    title_f = wb.add_format({**base, "bold": True, "font_size": 14})
    label_f = wb.add_format({**base, "bold": True})
    note_f = wb.add_format({**base, "italic": True, "font_color": "#595959"})

    ws = wb.add_worksheet("WO Log")
    sm = wb.add_worksheet("Summary")
    st = wb.add_worksheet("Settings")
    # Settings: the ONLY volatile function in the workbook lives here, once.
    st.write("A1", "Tracker settings", title_f)
    st.write("A3", "Report date", label_f)
    st.write_formula("B3", "=TODAY()", date_f)
    st.write("A4", "Every age and SLA calculation uses this one cell. "
                   "Overwrite it with a fixed date to re-run the tracker as of that day.", note_f)
    wb.define_name("ReportDate", "=Settings!$B$3")
    st.set_column("A:A", 14)
    st.set_column("B:B", 14)

    data = []
    for row in rows:
        data.append([dt.datetime.combine(v, dt.time()) if isinstance(v, dt.date) else v for v in row])
    n = len(rows)
    columns = [{"header": h, "format": date_f if h in ("Date Raised", "Target Date", "Date Completed")
                else text_f} for h in HEADERS]
    # Calculated table columns: Excel fills them down automatically when a new
    # work order is added. References are written fully qualified
    # (tblWO[[#This Row],[Col]]) so LibreOffice evaluates them too.
    columns += [
        {"header": "Days Open", "format": int_f,
         "formula": '=IF(tblWO[[#This Row],[Date Completed]]="",ReportDate-tblWO[[#This Row],[Date Raised]],tblWO[[#This Row],[Date Completed]]-tblWO[[#This Row],[Date Raised]])'},
        {"header": "SLA Status", "format": text_f,
         "formula": '=IF(tblWO[[#This Row],[Date Completed]]="",IF(ReportDate>tblWO[[#This Row],[Target Date]],"Overdue","On track"),IF(tblWO[[#This Row],[Date Completed]]>tblWO[[#This Row],[Target Date]],"Late","Met"))'},
        {"header": "Age Bucket", "format": text_f,
         "formula": '=IF(tblWO[[#This Row],[Days Open]]<=7,"0-7 days",IF(tblWO[[#This Row],[Days Open]]<=30,"8-30 days",IF(tblWO[[#This Row],[Days Open]]<=90,"31-90 days","90+ days")))'},
        {"header": "Month", "format": text_f, "formula": '=TEXT(tblWO[[#This Row],[Date Raised]],"yyyy-mm")'},
    ]
    ws.add_table(0, 0, n, len(columns) - 1, {
        "name": "tblWO", "data": data, "columns": columns,
        "style": "Table Style Light 9",
    })
    last = n + 1
    # Three conditional formatting rules, scoped to the data only.
    red = wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})
    amber = wb.add_format({"bg_color": "#FFEB9C"})
    grey = wb.add_format({"font_color": "#808080"})
    ws.conditional_format(f"A2:O{last}", {"type": "formula", "criteria": '=$M2="Overdue"', "format": red})
    ws.conditional_format(f"F2:F{last}", {"type": "cell", "criteria": "==", "value": '"Emergency"', "format": amber})
    ws.conditional_format(f"A2:O{last}", {"type": "formula", "criteria": '=$J2="Closed"', "format": grey})
    ws.freeze_panes(1, 0)
    ws.set_column("A:A", 13); ws.set_column("B:B", 14); ws.set_column("C:C", 10); ws.set_column("D:F", 12)
    ws.set_column("G:G", 30); ws.set_column("H:K", 13); ws.set_column("L:O", 11)

    # Summary: open-job counts per site and per technician, calculated ONCE
    # per site/tech instead of once per row over whole columns.
    sm.write("A1", "Open jobs summary", title_f)
    sm.write("A2", "Open = status Open, In Progress or On Hold. Counts read from tblWO.", note_f)
    sm.write_row(3, 0, ["Site", "Open jobs"], label_f)
    sites = sorted({r[2] for r in rows})
    open_sum = ('SUM(COUNTIFS(tblWO[{col}],{ref},tblWO[Status],{{"Open","In Progress","On Hold"}}))')
    for i, s in enumerate(sites):
        x = i + 5
        sm.write(x - 1, 0, s, text_f)
        sm.write_formula(x - 1, 1, "=" + open_sum.format(col="Site", ref=f"A{x}"), int_f)
    sm.write_row(3, 3, ["Technician", "Open jobs"], label_f)
    techs = sorted({r[7] for r in rows})
    for i, t in enumerate(techs):
        x = i + 5
        sm.write(x - 1, 3, t, text_f)
        sm.write_formula(x - 1, 4, "=" + open_sum.format(col="Assigned Tech", ref=f"D{x}"), int_f)
    sm.set_column("A:A", 10); sm.set_column("B:B", 10); sm.set_column("D:D", 12); sm.set_column("E:E", 10)
    sm.freeze_panes(4, 0)

    wb.close()
    return len(sites), len(techs)


if __name__ == "__main__":
    rows = make_rows()
    (ROOT / "before").mkdir(exist_ok=True)
    (ROOT / "after").mkdir(exist_ok=True)
    n_fmt = build_bloated(rows, ROOT / "before" / "WO_Tracker_bloated.xlsx")
    n_sites, n_techs = build_fixed(rows, ROOT / "after" / "WO_Tracker_fixed.xlsx")
    print(f"rows={len(rows)} distinct messy formats={n_fmt} sites={n_sites} techs={n_techs}")
