# Sprint PBI/Task Report Generator — Plan

## Problem
Need a Python script that pulls PBI + child Task/Bug data from Azure DevOps using the
provided saved query for each of three teams (FPSO, Foundation, and CSR) and produces a
single Excel workbook with per-team tabs plus a combined "All" tab. The script is
**read-only** — it must never modify anything in Azure DevOps. Task categorization uses
rule-based keyword matching for now (AI-assisted refinement is a possible future step,
which is why a Description column is included).

## Approach

### 1. Script location & language
- Single Python script at `C:\Utilities\SprintTask\sprint_report.py`.
- Uses `requests` (Azure DevOps REST API calls) and `openpyxl` (Excel output).
- PAT is read from `AZURE_DEVOPS_PAT` when set, then from the local ignored `localconfig`
  file. If neither is available, it is entered interactively as plain visible text.
  `localconfig` must never be committed; `localconfig.example` documents the expected key.

### 2. Run flow (interactive prompts)
On each run, the script prompts in this order:
1. Read `AZURE_DEVOPS_PAT`, then the local ignored `localconfig` file; if neither is available,
  prompt for the PAT as plain visible text (quotes optional and auto-stripped).
2. "Enter the Azure DevOps Query URL for FPSO:" → paste the saved query URL provided for FPSO
3. "Enter the Azure DevOps Query URL for Foundation:" → paste the saved query URL provided for Foundation
4. "Enter the Azure DevOps Query URL for CSR:" → paste WIQL query URL
   (e.g. `.../_queries/query/?tempQueryId=...`)
5. Since CSR is a Kanban board (no sprint start/end dates), the script prints a
   confirmation prompt reminding the user that CSR doesn't follow Agile sprint
   cadence, and asks the user to manually verify/enter the **start date** and
   **end date** they want to use to define the reporting period for CSR — no
   automatic date-field filtering logic is applied by the script; the user
   confirms the range is correct before proceeding.
6. Script fetches data for all three, builds the workbook, and saves it.

Each URL is parsed to extract org/project/query ID. No dictionary of default/prior URLs
is stored — user pastes fresh URLs each run as requested.

### 3. Data retrieval (Azure DevOps REST API, read-only calls only)
**All teams:**
- Execute the saved query identified by the supplied URL using
  `GET .../_apis/wit/wiql/{query-id}` to get exactly the work item IDs returned by
  that query. The script does not construct a replacement WIQL query.
- `GET .../_apis/wit/workitems?ids=...&$expand=relations` in batches to pull full
  field data + parent/child relations.
- Identify PBIs (`Product Backlog Item`) and their child `Task`/`Bug` work items via
  the `System.LinkTypes.Hierarchy-Forward` / `-Reverse` relations.

**Query scope guarantee:**
- Fetch full details only for IDs returned by the supplied query (`$expand=relations`).
- Group children under parent PBIs using each item's `System.Parent` field, but do not
  fetch additional children with a separate query.
- Validate before writing that every PBI/Task/Bug ID in the report is present in the
  corresponding supplied query result.
- The CSR report records the user-confirmed reporting period as a note; it does not
  silently filter query results by dates.

The supplied query is authoritative. If a child Task/Bug is not returned by that query,
it is intentionally excluded from the report.

### 4. Fields pulled per work item
- PBI: ID, Title, URL (constructed from org/project/id)
- Task/Bug: ID, Title, URL, Work Item Type, Assigned To, State
  Description, Original Estimate (`Microsoft.VSTS.Scheduling.OriginalEstimate`),
  Completed Work (`Microsoft.VSTS.Scheduling.CompletedWork`), Closed Date
  (`Microsoft.VSTS.Common.ClosedDate`)
- All states are included (not just Done) — actual state value is shown as-is.
- **Scope filter**: only `Product Backlog Item` work items and their child `Task`/`Bug`
  work items are included. Any other linked work item types (e.g. Feature, Epic, User
  Story used as something else) are ignored. A PBI with no Task/Bug children is still
  included as a row (with blank Task columns) so it's visible in the report.
- **CSR-specific rule**: if the CSR query's results contain no `Product Backlog Item`
  work items at all, the script skips generating the CSR tab entirely (no empty tab is
  created) and prints a message noting this to the user.

### 5. Task Categorization (rule-based, for now)
Applied to each Task/Bug row, in priority order:
1. Work item type = **Bug** → `Bug Fixing`
2. Title/Description contains keywords like `test`, `qa`, `validation`, `verify`,
   `regression` → `QA / Testing`
3. Title/Description contains `requirement`, `analysis`, `study`, `research`, `spec`,
   `investigat` → `Requirement Study`
4. Title/Description contains `review` or `demo` → `Review`
5. Title/Description contains `design` or `document` → `Design / Documentation`
6. Otherwise (Task type, no keyword match) → `Development`
- Keyword lists are defined as simple editable lists/constants near the top of the
  script so they can be tuned later.
- Bugs always remain `Bug Fixing`, regardless of title or description keywords.
- Demo tasks are treated as `Review`, not `Development`.
- A **Description** column is included in the output specifically so a future AI pass
  can re-review/refine categorization using both title and description without needing
  to re-run the ADO fetch.

### 6. Output — Excel workbook
Saved to `C:\Utilities\SprintTask\Reports\` (subfolder is created automatically if it
doesn't exist) with an auto-generated filename
(e.g. `Sprint_Report_<timestamp>.xlsx`).

Tabs:
- **All** — combined rows from FPSO + Foundation + CSR, with the `Team` column (first
  column) clearly identifying which team each PBI/row belongs to
- **FPSO** — FPSO-only rows
- **Foundation** — Foundation-only rows
- **CSR** — CSR-only rows

Each tab has the source Query URL written at the top (row 1), with the data
table starting a couple rows below it.

Columns (per row = one Task/Bug under a PBI):
1. Team (which team this row belongs to — FPSO, Foundation, or CSR; most useful on the
   combined "All" tab, but shown on every tab for consistency)
2. PBI ID (hyperlinked to the PBI in Azure DevOps)
3. PBI Title
4. Task/Bug ID (hyperlinked)
5. Task Type (Task / Bug)
6. Task Title
7. Assigned To
8. State
9. Original Estimate (hrs)
10. Completed Work (hrs)
11. Task Categorization (rule-based, see §5)
12. Done Date (Closed Date field)
13. Description (extra column, for future AI-assisted categorization review)

### 7. Tests and validation
- `test_sprint_report.py` tests categorization rules, including review, demo, and bug
  handling.
- Latest-report tests read the newest non-temporary workbook in `Reports/` and verify:
  - each PBI is assigned to only one team tab;
  - every bug is categorized as `Bug Fixing`, with report file, sheet, PBI, and bug
    details on failure;
  - every demo task is categorized as `Review`, with report file, sheet, PBI, and task
    details on failure;
  - report IDs are limited to the IDs returned by the supplied query.

### 8. Safety / read-only guarantee
- Script only issues read-only Azure DevOps API calls: `GET` for saved query results
  and work item details. No work items are created, updated, or deleted.

## Notes / Assumptions
- PAT is not hardcoded or logged. A local `localconfig` file may hold it for convenience,
  but `localconfig` is ignored by git and must remain local to the user's machine.
- FPSO/Foundation sprint or reporting scope is not hardcoded — it comes from the supplied
  query URL, so the same script works for any future query by pasting a new URL.
- CSR date range confirmation is a manual visual check by the user (prompt only) — the
  script does not silently filter data based on assumed correct dates.
- openpyxl and requests will be installed via pip if not already present.
