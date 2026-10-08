# PBI/Task Summary Report Generator — Plan

## Goal
Update the standalone summary utility so it reads a user-selected
`PBI_Bug_Task_Report_*.xlsx` workbook and writes a **new output copy** with a flat,
one-row-per-PBI table on a worksheet named exactly `Summary`. The layout, calculations,
category order, and cell-color requirements are specified in this plan; they must not
depend on a particular example workbook continuing to exist. PBI IDs and rows are always
derived dynamically from the selected input report and verified against live Azure DevOps
data.

The original input workbook must never be modified. The script must not make Azure
DevOps write requests. Existing report validation and live read-only test behavior remain
unless explicitly changed below.

## Commands

```powershell
Set-Location C:\Utilities\SprintTask
python .\pbi_bug_task_summary.py
python -m unittest -v test_pbi_bug_task_summary
python -m unittest discover -s . -p "test_*.py" -v
```

## Input workbook

- Prompt for the input `.xlsx` path; do not hardcode the sample filename or PBI IDs.
- Input is produced by `pbi_bug_task_report.py` and has `Summary`, `All`, `FPSO`, and
  `Foundation` tabs.
- Read report data from `All`, using column headers rather than fixed column positions.
- Group source rows dynamically by PBI ID. Include PBIs present in the input, including a
  PBI-only row when it has no Task/Bug rows.
- Continue to validate the input filename prefix, required sheets, and that its existing
  `Summary` sheet is empty before generating output.
- Continue to reject the workbook if any relevant row is assigned to Himanshu Pathak,
  using case/whitespace-insensitive matching and identifying the offending rows.

## Summary worksheet layout

Replace the existing multi-row-per-PBI category breakdown and inline pie charts with a
flat table: exactly one output row per PBI, no charts, no category subrows, no merged
PBI-cell blocks, no hidden chart-helper columns, and no chart spacer rows.

Use two header rows:

- Row 1: merge and label `Effort` above the six category-effort columns; merge and label
  `Percentage` above the six category-percentage columns.
- Row 2: column labels in this order:

1. PBI ID
2. PBI Title
3. PBI Story Points
4. Total Hours
5. Bug Fixing (under Effort)
6. QA / Testing (under Effort)
7. Requirement Study (under Effort)
8. Review (under Effort)
9. Design / Documentation (under Effort)
10. Development (under Effort)
11. Bug Fixing (under Percentage)
12. QA / Testing (under Percentage)
13. Requirement Study (under Percentage)
14. Review (under Percentage)
15. Design / Documentation (under Percentage)
16. Development (under Percentage)

The Summary worksheet name must be exactly `Summary`. Do not create a separate tab per
person or use a person's name as the Summary tab name. Ignore such names in the sample
workbook; it is only a reference for the requested format.

## Calculations and display

- PBI ID, PBI Title, and PBI Story Points come from the PBI fields already repeated in
  the input `All` tab. Preserve the PBI hyperlink when present.
- `Total Hours` is the sum of `Completed Work (hrs)` for all Task/Bug rows belonging to
  that PBI. Treat `No_Hours`, blank, and missing values as zero for arithmetic, while
  retaining a flag indicating that one or more task values were missing.
- Each Effort category value is the sum of Completed Work for that PBI's Task/Bug rows
  whose `Task Categorization` exactly matches that category.
- The six fixed categories and their display order are:
  `Bug Fixing`, `QA / Testing`, `Requirement Study`, `Review`,
  `Design / Documentation`, `Development`.
- Each Percentage cell is `category effort / Total Hours * 100`. If Total Hours is zero,
  every category percentage is zero. Display percentages consistently to one decimal
  place (e.g. `5.9%`, `0.0%`). For positive totals, category percentages should add to
  approximately 100%, allowing only for rounding.
- If any task hours are missing, retain the numeric total and add a visible note to that
  PBI's Total Hours cell, e.g. `35.5 (some tasks/bugs missing hours)`, and highlight it
  yellow. Do not treat an explicit numeric zero as missing.
- Preserve `No_Story_Point` when story points are absent and highlight that cell yellow.
- Fill the complete `A1:D2` header block dark blue (`FF002060`, RGB `#002060`) with bold white text
  (`FFFFFFFF`); retain visible borders around every cell in the block.
- Fill the complete `E1:J2` Effort block dark red (`FF9C0006`, RGB `#9C0006`) with bold
  white text (`FFFFFFFF`); retain visible borders around every cell in the block.
- Fill the complete `K1:P2` Percentage block dark blue (`FF002060`, RGB `#002060`) with
  bold white text (`FFFFFFFF`); retain visible borders around every cell in the block.
- Draw visible thin gray borders around every grouped header, column header, and data cell
  in the Summary table. Use the fixed color `FF808080` for the border lines.
- Match the Summary header/section/cell color scheme used by the approved reference sheets.
  Before implementation, inspect those sheets and record the exact ARGB values and their
  uses in a color specification (group-header fill, column-header fill, missing-data
  fill, font color, and borders). Implement the colors as named constants and test those
  exact values. Tests must not need to reopen the reference workbook to know expectations.
- A PBI with no child Task/Bug rows gets one row, Total Hours 0, and zero effort and
  percentage in every category.
- If a Task/Bug has an unexpected category not in the six fixed categories, do not
  silently drop its hours. Fail with a clear error identifying the PBI, Task/Bug ID,
  title, and unexpected category so the category list or source report can be reviewed.
- If category effort does not reconcile with Total Hours, treat it as a validation error
  and identify the PBI and contributing rows. Do not silently alter the source or
  manufacture balancing effort.

## Output and safety

- Save a new workbook in `Reports/Pbi_bug_tasks/`, creating the directory if necessary.
- Keep the source workbook untouched.
- Use the existing output naming convention: insert `_Summary_<fresh timestamp>` before
  the `.xlsx` extension, replacing an existing trailing timestamp when present.
- Keep local PAT handling and the Himanshu Pathak safety check as currently implemented.
- No pie charts, graph data, or workbook formulas are required for this flat summary.

## Unit-test changes

Update `test_pbi_bug_task_summary.py` to match the flat layout. Tests should use generated
in-memory/temporary workbooks and dynamic IDs; do not depend on the sample's fixed PBI
IDs except in a clearly named optional sample-regression test.

Required pure unit tests:

1. Summary worksheet is named `Summary` and has the two-row grouped header structure.
2. The exact 16 column labels and category ordering are correct.
3. Exactly one Summary data row is written for each unique PBI in `All`.
4. PBI ID, title, story points, and hyperlink are preserved.
5. Total Hours sums Completed Work, treating `No_Hours`/blank as zero and explicit zero
   as present data.
6. Missing-hours note and yellow fill appear only when one or more task values are
   missing; `No_Story_Point` is preserved and highlighted.
7. Category effort values aggregate to the correct category columns, with zeros for
   absent categories.
8. Category percentages use Total Hours as the denominator, use zero when total is zero,
   and sum to approximately 100% for positive totals.
9. A PBI without children still gets one row with zero totals/category values.
10. Unknown categories and any total/category reconciliation mismatch fail clearly and
    identify the relevant PBI/task rather than silently dropping data.
11. No charts, merged data blocks, spacer rows, or hidden helper columns are added.
12. The output is a new file and the input workbook remains unchanged.
13. Existing validation tests remain: filename prefix, required tabs, empty Summary,
    blocked-assignee rejection, and output filename generation.

14. Reference-style colors are represented by named, fixed ARGB constants. Test exact
   header fills, font colors, borders, and missing-data highlights against those
   constants; do not depend on opening a separate reference workbook at test time.

## Live Azure DevOps validation (required when credentials and report are available)

Live checks use the generated Summary workbook as the subject and the input report's
`All`, `FPSO`, and `Foundation` tabs to identify PBI IDs and embedded team query URLs.
Use the PAT from the ignored local `localconfig` file (or the existing environment
variable override), never print or ask the user to paste the secret, and perform only
read-only Azure DevOps API requests. Discover IDs dynamically; do not hardcode sample
PBI or Task/Bug IDs.

For every PBI in Summary:

1. **PBI identity and Story Points:** Fetch that exact PBI from Azure DevOps. Verify its
  type is Product Backlog Item and compare Summary's PBI ID/title and PBI Story Points
  with the live PBI fields. Check `Microsoft.VSTS.Scheduling.StoryPoints`. If Summary
  shows `No_Story_Point`, verify that the live field is actually unset; otherwise check
  the numeric/value match. Report PBI ID, title, Summary value, live value, and field
  name on mismatch.
2. **Task/Bug inventory:** Enumerate the PBI's live hierarchy child links, fetch the
  linked work items, and identify all Task and Bug children. Compare the live set with
  that PBI's Task/Bug rows in the report's `All` tab. Detect missing and extra items;
  report IDs/titles and the PBI they belong to.
3. **Total Hours:** Compute expected Total Hours from the live children themselves by
  summing their numeric Completed Work values. Treat explicit zero as present. Compare
  the result with Summary's Total Hours.
4. **Missing task hours:** For each live child with absent/blank Completed Work, verify
  it is identified as missing by the report-derived Summary. Verify the displayed total
  is the sum of only children with recorded hours, while the missing-hours annotation
  and highlight are present. List each missing Task/Bug ID/title. A `No_Hours` placeholder
  is missing data; numeric 0 is not.
5. **Category Effort:** Independently total live children by the six Summary categories
  using each child's corresponding Task/Bug categorization from the report. Compare each
  category effort cell with the sum of that category's live Completed Work. Missing hours
  contribute zero but remain listed as missing. Report unrecognized/blank categories.
6. **Effort reconciliation:** Verify the six category effort values sum to Total Hours.
  A mismatch is a failure, with the PBI and category values listed. Do not silently
  adjust totals or invent balancing effort.
7. **Percentages:** Recalculate every percentage as `category effort / Total Hours * 100`;
  when total hours is zero, all percentages must be zero. Compare using a tolerance of
  0.1 percentage point for one-decimal display. For positive totals, verify the six
  percentages sum to 100%, allowing only for display rounding.
8. **Missing Story Points:** For every `No_Story_Point` value, confirm directly from that
  PBI's live record that Story Points is unset. If present live but missing in Summary,
  or vice versa, fail with details.
9. **PBI with no children:** For every PBI represented without Task/Bug children, check
  live child relations and confirm there are no Task/Bug children. Verify Total Hours,
  all Effort cells, and all Percentage cells are zero.
10. **Query/team scope:** Re-run each exact query URL embedded in the corresponding team
   tab. Verify each reported PBI is in that team's query results and each child is
   actually a hierarchy child of an included PBI. Do not invent or broaden a query.

If the PAT is unavailable, a query fails, or the API cannot return required live fields,
mark the relevant check NOT EXECUTED or INCONCLUSIVE; never report it as passing. Live
checks must compare the generated Summary to Azure DevOps live PBI/task/bug data, not
merely compare Summary to another sheet in the same workbook.

## Generated-workbook tests

- Validate the newest generated summary workbook's Summary tab structure and row count.
- Independently recompute each PBI's Total Hours, category effort, and percentages from
  that workbook's `All` tab and compare them with Summary.
- Verify all used fills/fonts/borders against the hardcoded approved color constants and
  check the expected yellow missing-data cells.
- Keep live API validations above separate and explicitly marked, with a clear skip reason
  when no PAT or suitable report is available. Never fabricate live query results.

Remove/replace obsolete tests that assert pie-chart counts, chart helper data, per-PBI
category subrows, merged cells, or chart rendering options. Those behaviors are no longer
part of the requested Summary design.

## Implementation sequence

1. Change `CATEGORY_ORDER`/summary headers to the flat 16-column design.
2. Keep/reuse `read_all_tab()` and `summarize_pbis()` where appropriate; make unknown
   categories and reconciliation errors explicit.
3. Rewrite `write_summary_sheet()` to create grouped headers and one row per PBI; remove
   chart creation and chart helper logic/imports/constants.
4. Preserve workbook validation, blocked-assignee validation, output naming, and
   source-file immutability.
5. Replace chart/layout tests with calculation, format, validation, and no-chart tests.
6. Implement the live Azure DevOps tests above using the shared local PAT loader and
  workbook-embedded query URLs. These checks are required when credentials and live
  Azure DevOps are available; they must skip explicitly rather than use fabricated data
  when unavailable.
7. Run the full summary test suite and generate a new output workbook from a copy of a
  current report; inspect the `Summary` tab in Excel for final layout and color validation.

## Known reference-data inconsistency to verify

The shared example shows PBI `34791` with Total Hours 9, while displayed category effort
values are 4, 4, and 5 (sum 13), and its percentages therefore sum to 144.4%. The new
implementation must not reproduce this inconsistency silently. Its primary calculation
source is the Task/Bug rows in the selected input report's `All` tab. It should calculate
both total and category effort from those same rows, then fail clearly if they do not
reconcile. The reference workbook is a formatting example, not a source of hardcoded
values or expected task IDs.
