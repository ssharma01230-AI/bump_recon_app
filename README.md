# BuMP Monthly Reconciliation Tool

A reusable local app for the monthly reconciliation that Dan (WPP Account
Director, Western Europe) runs against the BuMP finance/commercial report.

Upload Dan's overview workbook and this month's BuMP export, click one
button, and download:

1. **An updated copy of Dan's overview workbook** — *Latest BuMP* and
   *Difference Since Jan 1st* columns refreshed, all other column names,
   row labels, formulas, styling, sheet names and structure left exactly
   as they were.
2. **A plain-English markdown summary** describing what changed, what's
   new, what's missing, and any data-quality flags.

## What it does

For each scope row in Dan's overview:

1. Reads the **Scope ID** (column D).
2. Finds all matching rows in the BuMP report where
   `version = CURRENT` and `tracker = CURRENT`.
3. Sums `project_fee_gbp` across those rows — the **Latest BuMP** value.
4. Writes that into Dan's **Latest BuMP** column (auto-detected, default
   column H).
5. Compares it against Dan's **Post Transition** baseline
   (`Sum of Total BAT Budget`, column F by default).
6. Writes `Latest BuMP − Baseline` into the **Difference Since Jan 1st**
   column (auto-detected, default column I).

Scopes that aren't in this month's BuMP report are **left unchanged** by
default — there is an opt-in "Clear unmatched output cells" toggle in
advanced settings if you ever want to wipe them.

Market subtotals, sub-group totals (e.g. *Romania Total* under
*WESTERN EUROPE*) and the Grand Total are recomputed as values (not
formulas) only when every scope row in the group has a numeric Latest
BuMP — otherwise the total is left alone and the markdown summary
flags it.

## Running locally

```bash
# 1. Install
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

# 2. Run
streamlit run app.py
```

The app opens in your browser. No command-line use required after that
— upload two files, click *Run reconciliation*, download outputs.

## Project layout

```
bump_recon_app/
├── app.py                       # Streamlit UI
├── reconciliation_engine.py     # Pure logic (UI-agnostic, reusable)
├── requirements.txt
├── README.md
└── tests/
    └── test_reconciliation.py   # 31 unit & acceptance tests
```

The engine is deliberately separated from the UI. To reuse it from a
scheduled job, an API, or a CLI:

```python
from openpyxl import load_workbook
from reconciliation_engine import update_overview_workbook, generate_markdown_summary

ovr = load_workbook("Dan_Overview.xlsx")
bmp = load_workbook("BuMP_2026_04_14.xlsx")

result = update_overview_workbook(ovr, bmp)
ovr.save("Dan_Reconciliation_Updated_2026-05-15.xlsx")

with open("summary.md", "w") as f:
    f.write(generate_markdown_summary(result))
```

## Running the tests

```bash
pip install pytest
python -m pytest tests/ -v
```

The suite includes synthetic-fixture tests (column detection, scope
normalisation, total-row grouping, end-to-end on a tiny in-memory
workbook) and two acceptance tests that run against the real workbooks
if they're available in `/home/claude/`.

## Auto-detection — and how to override it

The app auto-detects:

| Thing                       | How                                                                                                                    | Override                          |
| --------------------------- | ---------------------------------------------------------------------------------------------------------------------- | --------------------------------- |
| Overview sheet              | First sheet named "Overview", else first sheet with `scope_id` in row 1–3.                                             | Dropdown in *Advanced settings*.  |
| BuMP sheet                  | The sheet whose row 1 contains `scope_id` and has the most columns (BuMP exports are very wide).                        | Dropdown in *Advanced settings*.  |
| Data header row             | The row in the overview that contains both `scope_id` and the baseline column name. (Dan's file: row 3.)                | Auto-derived from baseline name.  |
| Baseline column             | Looks for `Sum of Total BAT Budget` on the data header row.                                                            | Text input.                       |
| Latest BuMP output column   | Looks on row 1 for headers containing "Latest BuMP", or "Total BAT" + "Budget" + a date.                               | Text input (column letter).       |
| Difference column           | Looks on row 1 for "Difference Since" or "+ / -".                                                                       | Text input (column letter).       |
| `version` / `tracker` filter| Defaults to `CURRENT` / `CURRENT`. Leave blank in Advanced settings to disable.                                         | Text input.                       |

If a column can't be auto-located, the app raises a clear error telling
you exactly which column was missing and how to specify it.

## Calculation behaviour — by example

Dan's row 15 (Scope ID `4733`, Denmark, VUSE):

- Baseline (Sum of Total BAT Budget): **£9,236.71**
- Sum of `project_fee_gbp` in BuMP where `version=CURRENT` and
  `tracker=CURRENT` for scope 4733: **£9,242.87**
- Difference written: `9242.87 − 9236.71 = +£6.16`

This matches the figure Dan already had in his file for the
`2026-04-14` cut — confirming the calculation matches the existing
manual process exactly.

## Preservation guarantees

The app **only writes values to specific target cells**:

- `Latest BuMP` cells of matched scope rows
- `Difference Since Jan 1st` cells of matched scope rows
- (optionally, when toggled) those same cells for unmatched scopes
- `Latest BuMP` and `Difference` cells of total rows that can be
  unambiguously recalculated

It never touches:

- Column header text on row 1 or row 3
- Row labels in columns A–C
- Other sheets in the workbook
- Number formats (the GBP accounting format is preserved on every cell)
- Fonts, fills, borders, alignment, row heights, column widths
- Formulas in cells outside the four target columns
- Macros (`.xlsm` files keep their VBA archive when re-saved)

## Outputs

After clicking *Run reconciliation* the app shows:

- 4 KPI tiles (total scopes, found, not found, new BuMP scopes)
- Total Latest BuMP value & net movement vs baseline
- A sortable per-scope table
- A table of new/unmatched BuMP scopes
- A collapsible warnings panel
- A preview of the markdown summary

Then two download buttons:

- `Dan_Reconciliation_Updated_YYYY-MM-DD.xlsx`
- `Dan_Reconciliation_Summary_YYYY-MM-DD.md`

The original uploaded files are never modified on disk.

## Data-quality checks surfaced in the markdown

- Duplicate Scope IDs in Dan's overview
- Missing baseline values
- BuMP rows excluded by the version/tracker filter
- Any scope where bat_market / country / brand disagree between the two
  files
- Any total row that couldn't be safely recalculated

## Currency handling

Calculations use `decimal.Decimal` and are rounded to 2 d.p. (half-up)
only when written into Excel. Scope IDs are compared as strings so
`4705`, `"4705"`, `4705.0` and `" 4705 "` all match, and leading zeros
(if they ever occur) are preserved.

## Limitations / future enhancements

- The unique row-grouping rule for "X Total" assumes the convention
  Dan's file already uses (market in col A, optional sub-group in col B,
  blank scope_id). If the layout changes substantially the total rows
  are skipped and flagged, never silently miscomputed.
- If BuMP ever ships a row without `project_fee_gbp` (truly missing,
  not just zero) that row is silently skipped in the sum; the markdown
  summary's data-quality section will flag it via row-count
  discrepancies. A future version could surface this more directly.
