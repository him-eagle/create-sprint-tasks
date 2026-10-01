# PBI Collection & Bug/Task Association Report — Plan

## Problem
Need a new, standalone Python script (inspired by the existing
`sprint_tasks_from_queries.py` / `SPRINT_TASKS_FROM_QUERIES_PLAN.md`) that, for two teams
(**FPSO** and **Foundation**):
1. First collects **all PBIs** (Product Backlog Items) returned by that team's supplied
   Azure DevOps query.
2. Then collects **all Bug/Task work items associated with those PBIs** (via parent/child
   relations), regardless of whether those child Bug/Task items were themselves returned
   by the query.
3. Produces an Excel report table with one row per Task/Bug (or one PBI-only row when a
   PBI has no children), containing the following columns:
   - Team
   - PBI ID
   - PBI Title
   - **PBI Story Points**
   - **PBI Done Date**
   - Task/Bug ID
   - Task Type (Task / Bug)
   - Task Title
   - Assigned To
   - State
   - Original Estimate (hrs)
   - Completed Work (hrs)
   - Task Categorization
   - Done Date (Task/Bug's own Closed Date)
   - Description

The query supplied for each team (FPSO, Foundation) is used only to determine **which
PBIs belong in the report**. Once that PBI set is fixed, the script walks each PBI's
parent/child hierarchy links to pull in **every** linked Task/Bug, so the report reflects
the complete, current set of work under each PBI rather than being limited to whatever a
single query happened to return. The script remains **read-only** — no Azure DevOps data
is created, updated, or deleted.

## Approach

### 1. Script location & language
- New file: `C:\Utilities\SprintTask\pbi_bug_task_report.py` (does not modify the existing
  `sprint_tasks_from_queries.py`).
- New test file: `C:\Utilities\SprintTask\test_pbi_bug_task_report.py`.
- Reuses the same style/libraries as the existing script: `requests` (Azure DevOps REST
  API) and `openpyxl` (Excel output).
- PAT handling identical to the existing script: read `AZURE_DEVOPS_PAT` env var, then the
  local ignored `localconfig` file, then prompt interactively (plain visible text) if
  neither is available. Reuses the same `localconfig` file/convention (no new secrets
  file).

### 2. Run flow (interactive prompts)
1. Resolve PAT (env var → `localconfig` → interactive prompt).
2. "Enter the Sprint name or number (used in the report file name):" → free-text entry,
   asked **first**, before either query URL. The script does not reject/parse this value
   with any specific pattern — whatever is entered is sanitized (any character that is
   invalid in a Windows filename, such as `\ / : * ? " < > |`, is replaced with `_`) and
   used as-is.
3. "Enter the Azure DevOps Query URL for FPSO:" → paste the saved query URL for FPSO.
4. "Enter the Azure DevOps Query URL for Foundation:" → paste the saved query URL for
   Foundation.
5. Script fetches/builds data for both teams and writes the workbook.

Only two teams this time — **no CSR step** and no manual date-range confirmation prompt
(that was CSR-specific in the old script and doesn't apply here).

**Output file name format:**
```
PBI_Bug_Task_Report_<Sprint>_<timestamp>.xlsx
```
- `<Sprint>` is the sanitized sprint name/number entered at the prompt above (e.g.
  `Sprint_23`, `2024-S12` → `2024-S12`).
- `<timestamp>` follows the same timestamp format already used by the existing
  `sprint_tasks_from_queries.py` script (`YYYYMMDD_HHMMSS`), so the file is still unique
  per run even if the same sprint name is reused.
- Example: entering `Sprint 23` produces `PBI_Bug_Task_Report_Sprint_23_20260101_093000.xlsx`.
- **Please confirm this format is acceptable, or tell me how you'd like it changed**
  (e.g. different ordering of sprint/timestamp, different separators, sprint name only
  with no timestamp, etc.).

### 3. Data retrieval (read-only Azure DevOps REST API calls only)

**Step A — Collect PBIs from the query:**
- Execute the supplied saved query via `GET .../_apis/wit/wiql/{query-id}` to get the
  work item IDs returned by that query (same helper logic as the existing script:
  `run_wiql_query_by_id` / `parse_wiql_response`).
- Fetch full field data for those IDs via `GET .../_apis/wit/workitems?ids=...&$expand=relations`
  in batches.
- Filter to only `Product Backlog Item` work items from this result set. Any other work
  item types returned by the query (Task, Bug, Feature, Epic, etc.) are **not** included
  as PBIs; they are ignored at this step (a Task/Bug incidentally returned by the query is
  not treated as a "root" item — it will only appear in the report if it is also a child
  of one of the included PBIs).

**Step B — Collect Bug/Task children associated with those PBIs:**
- From each included PBI's `relations` (expanded in Step A), gather child work item IDs
  via `System.LinkTypes.Hierarchy-Forward` links.
- Fetch full field data for all of these child IDs via a second batched
  `GET .../_apis/wit/workitems?ids=...` call (children not already fetched in Step A).
- Filter to only `Task` and `Bug` work item types among these children; ignore any other
  linked type (e.g. a linked Feature or Test Case).
- These children are included **regardless of whether their own ID appears in the
  original query result** — the query is only used to select the PBI set, not the child
  set.
- A PBI with no Task/Bug children is still included as a row (blank Task columns).

**Query scope guarantee:**
- Every **PBI** ID in the report must be present in that team's supplied query result
  (this is validated/asserted).
- Every **Task/Bug** ID in the report must be a child (via hierarchy relation) of a PBI
  that is in the report — it does not need to be in the query result itself.
- The script never performs a broader/replacement WIQL query; it only issues the one
  supplied query plus the two batched `workitems` detail fetches described above.

### 4. Fields pulled per work item
- **PBI**: ID, Title, URL, **Done Date** (`Microsoft.VSTS.Common.ClosedDate`),
  **Story Points** (`Microsoft.VSTS.Scheduling.StoryPoints`).
- **Task/Bug**: ID, Title, URL, Work Item Type, Assigned To, State, Description,
  Original Estimate (`Microsoft.VSTS.Scheduling.OriginalEstimate`), Completed Work
  (`Microsoft.VSTS.Scheduling.CompletedWork`), Closed Date
  (`Microsoft.VSTS.Common.ClosedDate`).
- All states are included (not just Done) — actual state value shown as-is, for both
  PBIs and their child Task/Bug items.
- **Missing-value placeholders:**
  - If a PBI has no Story Points value set, the **PBI Story Points** column shows
    `No_Story_Point` instead of being left blank.
  - If a Task/Bug has no value set for **Original Estimate** and/or **Completed Work**,
    that column shows `No_Hours` instead of being left blank (checked independently per
    column — e.g. a task could show a real Original Estimate but `No_Hours` for
    Completed Work if work hasn't been logged yet).
  - These placeholders apply only when the source ADO field is genuinely empty/null. An
    explicit value of `0` is treated as real data and displayed as `0`, not replaced with
    a placeholder.

### 5. Task Categorization (rule-based)
Same priority-ordered rules as `sprint_tasks_from_queries.py`:
1. Work item type = **Bug** → `Bug Fixing` (always, regardless of keywords).
2. Title/Description contains `test`, `qa`, `validation`, `verify`, `regression` →
   `QA / Testing`.
3. Title/Description contains `requirement`, `analysis`, `study`, `research`, `spec`,
   `investigat` → `Requirement Study`.
4. Title/Description contains `review` or `demo` → `Review`.
5. Title/Description contains `design` or `document` → `Design / Documentation`.
6. Otherwise (Task type, no keyword match) → `Development`.
- Reuse the existing `categorize_task` logic as-is (ported into the new script).
- A **Description** column is kept for future AI-assisted categorization review.

### 6. Output — Excel workbook
Saved to `C:\Utilities\SprintTask\Reports\Pbi_bug_tasks\` (subfolder created
automatically if missing) with the filename format described in §2
(`PBI_Bug_Task_Report_<Sprint>_<timestamp>.xlsx`), using the sprint name/number entered
at the start of the run.

Tabs (same structure as old script, minus CSR, plus a new empty Summary tab):
- **Summary** — left intentionally empty (no data, headers, or formulas) for now; a
  placeholder tab for future manual/automated summary content. Created first so it's the
  leftmost/first tab in the workbook.
- **All** — combined rows from FPSO + Foundation, with `Team` as the first column.
- **FPSO** — FPSO-only rows.
- **Foundation** — Foundation-only rows.

Each tab has the source Query URL written at the top (row 1), with the data table
starting a couple rows below it. (The Summary tab has no query URL or data — it stays
completely empty.)

Note: the `No_Story_Point` / `No_Hours` placeholders only apply when a PBI/Task/Bug row
exists but the field itself is empty. A PBI-only row (no Task/Bug children at all) still
leaves the Task/Bug columns fully blank, as there is no task to apply `No_Hours` to.

Columns (per row = one Task/Bug under a PBI, or one PBI-only row):
1. Team
2. PBI ID (hyperlinked)
3. PBI Title
4. **PBI Story Points** *(new — shows `No_Story_Point` when unset)*
5. **PBI Done Date** *(new — from PBI's Closed Date field)*
6. Task/Bug ID (hyperlinked)
7. Task Type (Task / Bug)
8. Task Title
9. Assigned To
10. State
11. Original Estimate (hrs) *(shows `No_Hours` when unset)*
12. Completed Work (hrs) *(shows `No_Hours` when unset)*
13. Task Categorization (rule-based, see §5)
14. Done Date (Task/Bug's Closed Date — kept distinct from PBI Done Date)
15. Description

### 7. Tests and validation (inspired by `test_sprint_tasks_from_queries.py`)
`test_pbi_bug_task_report.py` has no hardcoded/static Azure DevOps query IDs anywhere.
Tests fall into two groups:

**A. Pure unit tests** (no Azure DevOps access at all) — exercise `categorize_task`,
`format_story_points` / `format_hours` placeholder logic, `sanitize_filename_component`,
`validate_report_scope`, `build_rows`, `filter_types`, `collect_child_ids`, and
`write_workbook`'s sheet structure, using hand-built in-memory work item dictionaries
passed directly into these functions. Covering:
- **Categorization rules**: review/demo tasks categorized as `Review` and not
  `Design/Documentation`; Bugs always categorized as `Bug Fixing` regardless of keywords.
- **Report scope validation**:
  - A PBI ID that is not part of the supplied query result is rejected.
  - A Task/Bug ID that is a child of an included PBI is **accepted even when it is not**
    part of the supplied query result (the core new behavior of this script).
  - A Task/Bug ID that is **not** a child of any included PBI is rejected/excluded.
  - A PBI linked to a non-Task/Bug child (e.g. Feature, Test Case) does not produce a row
    for that child.
- **Missing-value placeholders**: missing Story Points / Original Estimate / Completed
  Work each produce their placeholder; explicit `0` in any of these fields is preserved
  as `0`, not replaced with a placeholder.
- **Workbook structure**: Summary tab is empty and first; a team with zero rows still
  gets its own tab; the output directory is auto-created if missing.

**B. Tests against the latest actually-generated report** — these require a real report
to exist in `Reports\Pbi_bug_tasks\` (produced by actually running
`pbi_bug_task_report.py`, not by mocking). They locate the most recently modified
`PBI_Bug_Task_Report_*.xlsx` file, and:
- `LatestReportStructureTests` — read-only checks against the file itself: tab order
  (Summary, All, then team tabs), Summary stays empty, each PBI appears on only one team
  tab, bug/demo rows are categorized correctly, Story Points/Hours columns never sit
  blank where a placeholder belongs, the "All" tab row count equals the sum of the team
  tabs, and PBI/Task hyperlinks resolve to the correct work item URLs. These tests only
  need the report file — no PAT required — and are **skipped with a message** if no
  report exists yet.
- `LatestReportLiveApiTests` — reads each team's **actual Source URL out of row 1 of its
  tab** (exactly as the script wrote it) and re-runs that exact query against the live
  Azure DevOps API, rebuilding the expected PBI→Task/Bug mapping and expected Story
  Points/Hours values the same way `pbi_bug_task_report.py` itself does (Step A + Step
  B). It then asserts the report's actual contents match that freshly recomputed expected
  state, including a **round-trip PBI coverage/count check**: every PBI returned by
  re-running the live query is present somewhere in that team's report tab, and the
  number of distinct PBIs returned by the live query exactly equals the number of
  distinct PBIs found in the report tab (catching both a PBI the query returns but the
  report is missing, and a stale/mismatched PBI count). These tests require a configured
  PAT (`AZURE_DEVOPS_PAT` or `localconfig`) in addition to an existing report, and are
  **skipped with a message** if either is missing — they never fall back to fake/mocked
  query data.

### Execution order (required — these tests validate a report that must already exist)
1. Generate a report first by actually running the script interactively:
   ```
   cd C:\Utilities\SprintTask
   python pbi_bug_task_report.py
   ```
   Enter the PAT (if not already set via `AZURE_DEVOPS_PAT`/`localconfig`), the sprint
   name/number, and the FPSO and Foundation query URLs when prompted. This writes a new
   `PBI_Bug_Task_Report_<Sprint>_<timestamp>.xlsx` file into
   `Reports\Pbi_bug_tasks\`.
2. Only after that file exists, run the test suite so the latest-report tests have a real
   file (and a real, freshly re-queryable Azure DevOps API) to validate against:
   ```
   cd C:\Utilities\SprintTask
   pytest test_pbi_bug_task_report.py -v

   or

   python -m unittest -v test_pbi_bug_task_report
   ```

If step 2 is run without first completing step 1, the pure unit tests still run and pass,
but `LatestReportStructureTests` and `LatestReportLiveApiTests` are skipped (not failed)
with a message explaining that a report must be generated first.

### 8. Safety / read-only guarantee
- Script only issues read-only Azure DevOps API calls: `GET` for saved query results and
  work item details (including the follow-up children fetch). No work items are created,
  updated, or deleted.

## Notes / Assumptions
- PAT is not hardcoded or logged; same `localconfig` convention as the existing script.
- FPSO/Foundation query scope is not hardcoded — it comes from the supplied query URL
  each run (no stored URL list).
- No CSR team and no manual date-range confirmation step in this script.
- `Microsoft.VSTS.Common.ClosedDate` is reused as the Done Date field for both PBIs and
  Task/Bug items (confirmed with user); `Microsoft.VSTS.Scheduling.StoryPoints` is used
  for PBI Story Points (confirmed with user).
- openpyxl and requests will be installed via pip if not already present (same as before).
