# Loom walkthrough: 90 seconds

Record on your own laptop, in Excel, with both files open side by side.
Do the Excel timing first so the numbers you say out loud are your own.

## Step 0: time both files in Excel (before recording)

1. Open `before/WO_Tracker_bloated.xlsx`. Wait until the status bar says **Ready**.
2. Press **Alt+F11** to open the VBA editor, then **Ctrl+G** for the Immediate window.
3. Paste this line and press Enter. It prints the full recalculation time in seconds:

   ```
   t = Timer: Application.CalculateFull: Debug.Print Format(Timer - t, "0.00")
   ```

4. Run it three times and keep the middle number. Repeat for `after/WO_Tracker_fixed.xlsx`.
5. Close the VBA editor without saving anything (no macro is stored in the files).
6. Regenerate the cover with your numbers, for example:

   ```
   python demo/make_cover.py --recalc-before 42 --recalc-after 0.3 --recalc-label "Excel 365"
   ```

## Script

**0:00–0:10 · Hook** (bloated file on screen)
"This work order tracker holds 3,000 jobs. It's almost 3 MB, and every time someone
changes a status, Excel freezes for [X] seconds."

**0:10–0:35 · Diagnose** (show each thing as you say it)
- Change one Status cell. Let the "Calculating" freeze show.
- **Ctrl+End**: "The data stops at row 3,001, but Excel thinks the sheet runs to row 60,001."
- Click a Days Open cell: "Every row uses OFFSET, INDIRECT and TODAY, so all of them
  recalculate on every edit, and the open-job counts scan entire columns, 3,000 times over."
- Home > Conditional Formatting > Manage Rules: "753 rules, from years of copy-paste."

**0:35–0:45 · The tool**
Run `python tracker_doctor.py before/WO_Tracker_bloated.xlsx` in a terminal.
"I wrote a script that finds all of this in seconds, without opening Excel."

**0:45–1:15 · The fix** (fixed file on screen)
- Change a Status cell: no freeze.
- **Ctrl+End**: lands on row 3,001.
- "It's now an Excel Table, so formulas fill down automatically when you add a job."
- Settings sheet: "TODAY() lives in one cell instead of 30,000."
- Summary sheet: "Open jobs per site and per technician are calculated once here,
  not on every row."
- "Same 3,000 jobs, and I checked every row: identical results."

**1:15–1:30 · Close**
"[size] down to [size], recalculation from [X] seconds to under a second.
If your tracker is slowing your team down, send me a copy with dummy data and I'll send
back a free diagnosis."
