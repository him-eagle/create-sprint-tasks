# PBI/Task Summary Report Generator — Plan

## Problem
Build a new, standalone Python script that takes an **already-generated**
`PBI_Bug_Task_Report_*.xlsx` workbook (produced by the existing
`pbi_bug_task_report.py` at `C:\Utilities\SprintTask\pbi_bug_task_report.py`) and fills
in its empty **Summary** tab with a per-PBI rollup: story points, total completed-work
hours, a category/hours/percentage breakdown, and an inline pie chart per PBI. The script
must be read-only toward Azure DevOps (only used for live test verification, never for
mutating the workbook's source data) and must never silently use fabricated data in its
tests — all "live" tests re-query Azure DevOps for ground truth.

Confirmed from inspecting a real sample workbook
(`C:\Utilities\SprintTask\Reports\Pbi_bug_tasks\PBI_Bug_Task_Report_103_20261001_104744.xlsx`):
- Tabs: `Summary` (empty, 1x1), `All`, `FPSO`, `Foundation` (no CSR in this report type).
- `All`/`FPSO`/`Foundation` row 1 = `Source URL:` label + hyperlinked team query URL.
- Row 2 headers: `Team, PBI ID, PBI Title, PBI Story Points, PBI Done Date, Task ID,
  Task Type, Task Title, Assigned To, State, Original Estimate (hrs),
  Completed Work (hrs), Task Categorization, Done Date, Description`.
- `PBI Story Points` already shows the literal string `No_Story_Point` when unset.
- `Original Estimate (hrs)` / `Completed Work (hrs)` already show literal `No_Hours` when
  unset (0 is preserved as a real value, not replaced).

## Decisions confirmed with user
- New standalone script (does not modify `sprint_tasks_from_queries.py` or
  `pbi_bug_task_report.py`); lives in this repo alongside the other report scripts.
- "Total hours of PBI" = sum of **Completed Work (hrs)** across that PBI's Task/Bug rows
  (treating `No_Hours` as 0).
- Story Points are read directly from the existing `PBI Story Points` column in `All`
  (already present in the real workbook) — no live ADO fetch needed to get them. **In the
  Summary tab**, when the value is `No_Story_Point`, that cell is also highlighted
  **yellow** (same treatment as the missing-hours cell) so missing story points are
  visually flagged, not just as plain text.
- Missing-hours note: when any underlying task/bug for a PBI had `No_Hours`, the **same
  cell** as the total-hours number gets an appended note (e.g.
  `"32 (some tasks/bugs missing hours)"`) and the cell fill is **yellow**.
- Category breakdown layout: one **PBI header row** (ID/Title/Story Points/Total Hours)
  followed by **one sub-row for every Task Categorization value in the fixed category
  list** (`Bug Fixing`, `QA / Testing`, `Requirement Study`, `Review`,
  `Design / Documentation`, `Development`) — **not just the categories that PBI
  happens to have**. If a PBI has no tasks in a given category, that category's sub-row
  still appears, with **Category Hours = 0** and **Category % = 0%**. The PBI's
  ID/Title/Story Points/Total Hours cells are **vertically merged** across all of its
  category sub-rows (always the same fixed count of sub-rows per PBI, one per category).
  Each sub-row has: Task Categorization name, then **two separate columns** —
  **"Category Hours"** and **"Category %"** (of that PBI's total) — as distinct cells,
  not combined into one.
- A PBI with **no Task/Bug children** gets a single row (ID/Title/Story Points/Total
  Hours = 0), no category sub-rows, and no pie chart.
- Pie chart: **one per PBI**, placed **inline to the right of that PBI's own row block**
  (not a separate gallery section), with each slice labeled with both hours and
  percentage. The chart plots only the **nonzero** categories for that PBI (a 0-hour
  category contributes no visible slice); the fixed-category table rows above still list
  every category including the 0-hour ones. No chart for PBIs with no children at all.
- Validation: workbook must contain tabs `Summary`, `All`, `FPSO`, `Foundation` (error
  otherwise). `Summary` must be empty — if it already has content, **abort with a clear
  error** telling the user to supply a fresh/unprocessed report file.
- The script itself (not just the test suite) scans `Assigned To` across
  `All`/`FPSO`/`Foundation` before building the Summary. If any row is assigned to
  **"Himanshu Pathak"**, the script **aborts** and prints the offending rows (sheet,
  PBI ID, Task/Bug ID, title) with a message suggesting the user **remove those entries
  from the workbook first, then re-run the script**. It does not silently skip/filter
  them or proceed.
- Input file name must match `PBI_Bug_Task_Report_*` (validated, not just conventionally
  named) — reject other filenames with a clear error.
- Output: a **new file**, original untouched. Filename = input name with `_Summary`
  inserted **before the timestamp** component, e.g.
  `PBI_Bug_Task_Report_103_Summary_<timestamp>.xlsx`, where `<timestamp>` is a **fresh
  timestamp** taken when the summary script runs. Saved into
  `C:\Utilities\SprintTask\Reports\Pbi_bug_tasks` (same folder convention as
  `pbi_bug_task_report.py`'s output, regardless of where the input file itself was
  read from; created automatically if missing).

## Live tests requested (must hit real Azure DevOps, no mocked/fake data)
1. Re-query each team's saved query (URL read from row 1 of `FPSO`/`Foundation` tabs) and
   verify every PBI ID appearing in the `All` tab is actually returned by one of those
   queries — on mismatch, report verbose reasoning (which query was checked, what it
   returned, which PBI was expected and missing).
2. Recompute each PBI's total Completed Work hours directly from a fresh query/fetch of
   its Task/Bug children (treating missing hours as 0) and assert it matches the
   Summary tab's total-hours value for that PBI.
3. For every PBI that the Summary shows with **zero children** (no Task/Bug rows),
   independently re-query Azure DevOps for that PBI's child work items and assert the
   live result is indeed empty — confirms the "0 hours / no breakdown" case is real, not
   a scope/fetch bug.
4. Scan `Assigned To` across `All`/`FPSO`/`Foundation` tabs and assert **no** row is
   assigned to **"Himanshu Pathak"** — report any violating rows (sheet, PBI ID, task ID,
   title) if found. This mirrors the script's own pre-flight check: the test
   independently verifies the same rule the script enforces at runtime.
5. For every PBI/category sub-row in the Summary where **Category Hours = 0**,
   independently cross-reference that PBI's rows in `All` (and the matching team tab —
   `FPSO` or `Foundation`) and assert there is genuinely **no** Task/Bug row for that PBI
   whose `Task Categorization` equals that category — i.e. confirm the 0 is real, not a
   grouping/lookup bug, with verbose reasoning (PBI ID, category, rows inspected) on
   failure.
6. Plus standard unit tests (no ADO access) for: categorization passthrough, hours
   summation/placeholder logic, story point passthrough, filename validation, tab
   validation (including the "Summary must be empty" abort path), PBI-only row handling,
   and output filename generation.
7. Additional sensible test cases: category percentage math sums to ~100% per PBI,
   merged-cell ranges are correct for multi-category PBIs, every PBI always has exactly
   six category sub-rows (the fixed list), and each pie chart's slice count matches the
   number of nonzero categories for that PBI.

## Implementation outline
1. **New script** `pbi_bug_task_summary.py`:
   - Prompt for input file path; validate filename prefix `PBI_Bug_Task_Report_`.
   - Load workbook with `openpyxl` (need formulas/formatting preserved for output, so
     load normally — not `read_only` — since we must write a new workbook with charts).
   - Validate required tabs exist and `Summary` is empty; abort with clear message
     otherwise.
   - Scan `Assigned To` across `All`/`FPSO`/`Foundation`; abort with offending-row detail
     if "Himanshu Pathak" is found anywhere.
   - Read `All` tab rows by header name (not fixed column index) for resilience.
   - Group rows by PBI ID; compute total Completed Work hours (treating `No_Hours` as 0)
     and a missing-hours flag; compute hours/percentage for each of the six fixed
     categories (0 when absent).
   - Write the Summary sheet: header row, then per-PBI merged blocks with six category
     sub-rows each, yellow highlight on flagged Story Points/total-hours cells, and an
     inline pie chart (openpyxl `PieChart`, nonzero categories only) anchored beside each
     PBI's block.
   - Save to a new output file in `C:\Utilities\SprintTask\Reports\Pbi_bug_tasks` (input
     filename with `_Summary_<fresh timestamp>` inserted).
2. **New test file** `test_pbi_bug_task_summary.py`:
   - Pure unit tests using in-memory/fixture workbooks (built with openpyxl in a temp
     file) — no ADO access.
   - Live tests that read the real `PAT` via `AZURE_DEVOPS_PAT`/`localconfig` and the
     query URLs embedded in a real generated report; skipped with a clear message when
     no PAT or no report is available (same pattern as the existing test suites).
3. Update `README.md` only if it already documents the sibling scripts (currently just a
   title) — otherwise leave as-is to avoid unrelated doc churn.

## Safety / read-only guarantee
- The script only reads the input workbook and (in live tests only) issues read-only
  Azure DevOps `GET`/WIQL calls. No Azure DevOps data is created, updated, or deleted,
  and the original input workbook is never modified — all output goes to a new file.

## Implementation notes (post-build)
- `pbi_bug_task_summary.py` and its tests (`test_pbi_bug_task_summary.py`) were added to
  `C:\Utilities\SprintTask` alongside the existing scripts (confirmed location).
- Ran the script against the real sample workbook
  (`PBI_Bug_Task_Report_103_20261001_104744.xlsx`); verified the Summary tab's structure,
  merged cells, yellow highlights, and pie charts render correctly.
- Fixed a real openpyxl gotcha: `PatternFill(start_color="FFFF00", ...)` (6-digit hex)
  stores an alpha of `00` (fully transparent) instead of `FF` (opaque) — the fill is
  invisible in Excel unless the 8-digit ARGB form (`"FFFFFF00"`) is used.
- `pbi_bug_task_report.py`'s own test suite (`test_pbi_bug_task_report.py`) globbed
  `PBI_Bug_Task_Report_*.xlsx` in the shared `Reports\Pbi_bug_tasks` folder to find "the
  latest report" — this incidentally matched the new script's
  `PBI_Bug_Task_Report_<sprint>_Summary_<timestamp>.xlsx` output (whose Summary tab is
  intentionally non-empty), causing a false failure in the sibling suite. Fixed by
  excluding `_Summary_` from that glob in `test_pbi_bug_task_report.py`.
- One live test (`test_pbis_shown_with_zero_children_are_genuinely_empty_live`) is
  legitimately skipped against the current sample data, since that sample's PBIs all
  have at least one Task/Bug child — there's no zero-children PBI to verify yet. It will
  run automatically once a workbook with such a PBI is summarized.

## Follow-up refinements (post-build)
- **Consistent category colors:** each of the six fixed categories now has a dedicated
  hex color (`CATEGORY_COLORS`), applied via an openpyxl `DataPoint`/`GraphicalProperties`
  solid fill per pie-chart slice. Because every PBI's chart always lists the same six
  categories in the same order, the same category renders in the exact same color on
  every PBI's chart (e.g. "Development" is always green). An unexpected/unknown category
  falls back to a small cycled palette (`FALLBACK_CATEGORY_COLORS`).
- **Removed the chart legend:** the default legend box only added a generic "Series1"
  entry and wasted space without identifying categories usefully. The legend is now
  disabled (`chart.legend = None`); each slice's own data label instead shows the
  category name, value, and percentage directly (`showCatName`/`showVal`/`showPercent`
  all `True`), so the chart conveys the same information in less space.
- **Category % cell now displays a "%" suffix:** the cell's underlying numeric value is
  unchanged (e.g. `16.7`, not `0.167`), but a custom number format (`0.0"%"`) appends a
  literal `%` character so it reads `16.7%` in Excel, instead of a bare, ambiguous
  number. (A plain Excel percentage format was avoided since that would multiply the
  stored value by 100 for display, which would have required dividing the value by 100
  first — the custom-suffix format keeps the stored number intuitive for anything else
  reading the sheet while still looking like a percentage.)
- **Removed a stray "Series1" label from the pie charts:** even after removing the
  legend, Excel was still rendering a "Series1" text label on the chart because the
  data series was never explicitly told to suppress its series name
  (`showSerName`) — its unset/`None` default does not reliably render as off in Excel.
  Fixed by explicitly setting `chart.dataLabels.showSerName = False` (plus
  `showLegendKey = False` and `showBubbleSize = False` for the same reason), verified
  directly in the saved chart XML (`<showSerName val="0"/>`) and covered by a new test.
- **Eliminated overlapping 0-value slice labels:** the pie chart's data/category
  references previously spanned the full fixed six-category block, so a zero-hour
  category (e.g. "Design / Documentation: 0, 0%") still produced a data label that
  visually overlapped neighboring labels even though its slice had no visible wedge.
  Fixed by writing each PBI's **nonzero-only** categories into a pair of hidden helper
  columns (`CHART_HELPER_CATEGORY_COL` / `CHART_HELPER_HOURS_COL`, columns AX/AY,
  `column_dimensions[...].hidden = True`) and building the chart's data/category
  `Reference`s from that nonzero-only, contiguous range instead of the visible
  six-row table. The visible category breakdown table (columns E/F/G) is unaffected
  and still lists all six fixed categories, including zero-hour ones — only the chart
  itself now omits them, so it only ever shows slices/labels for categories that
  actually have hours.
- **Regression: hiding the helper columns made every chart render completely empty.**
  openpyxl/Excel charts default to "plot visible cells only"
  (`chart.visible_cells_only` / OOXML `<plotVisOnly val="1"/>`). Since the chart's data
  now lived entirely in the hidden helper columns introduced by the fix above, Excel
  excluded all of it from the plot by default — every pie chart rendered with no
  slices at all. **This was not caught by the existing unit tests**, because those
  tests only inspect the openpyxl chart object model (data references, series,
  colors, counts) — they confirm the XML/object structure is well-formed, but
  openpyxl never actually renders a chart, so a "structurally valid but visually
  empty in Excel" bug like this is invisible to that kind of test. Fixed by explicitly
  setting `chart.visible_cells_only = False` on every chart created, and added a
  regression test asserting this attribute directly so the specific setting is
  pinned going forward (still can't substitute for actually opening the file in
  Excel, which remains the authoritative check for visual rendering).
- **Simplified slice labels to percentage-only:** each slice previously showed category
  name + hours + percentage, which overlapped/looked cluttered for longer category
  names (e.g. "Requirement Study", "Design / Documentation"). Since the adjacent,
  always-visible category breakdown table (columns E/F/G) already lists the category
  name and hours for every row, the chart's data labels now show only the percentage
  (`showCatName`/`showVal` set to `False`, `showPercent` stays `True`) — consistent
  category colors (see above) plus the adjacent table are enough to identify each
  slice without repeating text on the chart itself.
- **Re-added a compact bottom legend so a bare "75%" slice is still identifiable.**
  Percentage-only labels made each slice's category ambiguous on its own, so a
  `Legend(legendPos="b")` was added back (Excel resolves pie-chart legend entries from
  the category reference at render time, so this correctly shows real category names,
  not a repeat of the earlier "Series1" issue, which was specifically caused by
  `showSerName` on data labels, not the legend). Chart height was bumped from 7cm to
  8.5cm and `CHART_SPACER_ROWS` from 12 to 14 to make room for the legend without
  overlapping the next PBI's block.
- **Added a live-data regression test for the legend/category mapping.** Per explicit
  request, `test_chart_legend_categories_match_live_recomputed_categorization`
  (in `LatestSummaryLiveApiTests`) does not use dummy/fixture data: for each PBI with
  children in the latest generated summary workbook, it independently re-fetches that
  PBI's current child work items from Azure DevOps, recomputes each child's category
  with the real `categorize_task` logic and its Completed Work hours from scratch, and
  asserts the resulting nonzero-category set exactly matches the hidden helper-column
  data that feeds that PBI's chart slices/legend. Requires a configured PAT and an
  existing summary workbook; skipped with an explanation otherwise, like the other
  live tests.
- **Added a live-data test for the Category % math itself.**
  `test_category_percentages_match_live_recomputation` verifies, for every PBI with
  children, that the Summary's "Category %" value matches
  `percent(category) = hours(category) / total_hours * 100` (0 when `total_hours == 0`,
  rounded to 1 decimal) — the exact formula used in `summarize_pbis()`. Both
  `hours(category)` and `total_hours` are recomputed from scratch from a fresh live
  fetch of that PBI's current Task/Bug children (categorized with `categorize_task`,
  hours via `hours_to_number`), not read from the "All" tab or anywhere else in the
  workbook, so the check is independent of whatever the generator script itself wrote.

## Execution steps
1. **Generate the summary workbook:**
   ```
   cd C:\Utilities\SprintTask
   python pbi_bug_task_summary.py
   ```
   Enter the path to an existing `PBI_Bug_Task_Report_*.xlsx` when prompted. Output is
   written to `Reports\Pbi_bug_tasks\PBI_Bug_Task_Report_<sprint>_Summary_<timestamp>.xlsx`;
   the input file is left untouched.
2. **Run the tests** (after step 1, so file- and live-API-dependent tests execute
   instead of skipping):
   ```
   cd C:\Utilities\SprintTask
   pytest test_pbi_bug_task_summary.py -v
   ```
   or
   ```
   python -m unittest -v test_pbi_bug_task_summary
   ```
   Set `AZURE_DEVOPS_PAT` (or populate `localconfig`) beforehand so the live Azure
   DevOps API checks run instead of being skipped.
