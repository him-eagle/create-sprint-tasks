# AI Workbook Validation Prompt

Copy this prompt for each new report run. Provide only the current workbook path, or let
the AI select the newest workbook. Task IDs and query URLs are discovered dynamically
from the current workbook; they are not hardcoded.

```text
You are validating one newly generated Azure DevOps PBI/task report.

## Run input

Workbook path:
<PASTE THE CURRENT .XLSX PATH HERE>

The Azure DevOps PAT is available locally in:
<WORKSPACE>\localconfig

Read the PAT from `localconfig` using the `AZURE_DEVOPS_PAT` key. Never print, echo, log,
quote, or include the PAT in your response. If the file cannot be read, mark live query
validation as NOT EXECUTED and continue with workbook-only checks. Do not ask the user to
paste the PAT into the response.

Do not modify the workbook, Azure DevOps, `localconfig`, source code, or query definitions.
Use read-only operations only.

## Files and tabs

Use the supplied workbook path. If blank, select the newest non-temporary `.xlsx` file in
`Reports/` and state which file was selected. Inspect only these four tabs:
`All`, `FPSO`, `Foundation`, and `CSR`.

Ignore all other worksheets, including console-log or scratch sheets.

Expected columns are:
`Team`, `PBI ID`, `PBI Title`, `Task ID`, `Task Type`, `Task Title`, `Assigned To`,
`State`, `Original Estimate (hrs)`, `Completed Work (hrs)`, `Task Categorization`,
`Done Date`, `Description`.

## 1. Query scope validation

For each of FPSO, Foundation, and CSR:

1. Read the source URL from row 1 of that worksheet and use that exact URL. The URL in
  the worksheet is authoritative for that report tab. Do not request a second query URL.
2. Do not invent or broaden WIQL, fetch extra children, or replace the URL with a query
   reconstructed from a sprint name.
3. Read the PAT from `localconfig` and execute the exact worksheet URL using read-only API
   calls.
4. Record every work-item ID returned by the query.
5. Compare every nonblank PBI ID and Task/Bug ID in the matching worksheet with the IDs
   returned by that exact query.
6. Report every workbook ID absent from the query result.
7. Report every query-returned PBI, Task, or Bug absent from the worksheet.
8. For a missing child, distinguish whether the query returned:
   - the child and its parent PBI;
   - the child but not its parent PBI;
   - only the parent PBI;
   - an unsupported work-item type.
9. Confirm the source URL in row 1 of each team tab is present, valid, and corresponds to
  the team tab.
10. Confirm every row in a team tab has the correct Team value.

A valid sprint-backlog URL and a valid saved-query URL are both acceptable when stored in
the workbook, provided the corresponding read-only lookup can be performed. If the URL
is missing or execution is not possible, report NOT EXECUTED and the exact reason; never
call it PASS.

## 2. Categorization validation using judgment

For every nonblank Task/Bug row, inspect both Task Title and Description. Apply these
rules as guidance in this priority order:

1. Work item type Bug -> Bug Fixing.
2. Testing intent, such as test, QA, validation, verification, or regression -> QA / Testing.
3. Requirements/investigation intent, such as requirement, analysis, study, research,
   specification, or investigation -> Requirement Study.
4. Review, PR review, code review, peer review, demo, or demonstration intent -> Review.
5. Design or documentation intent -> Design / Documentation.
6. Implementation or maintenance work with no stronger signal -> Development.

Use semantic meaning, not only literal keyword matching. Examples:
- Code Review and PR-Review are Review.
- Demo Task is Review, not Development.
- A Bug is always Bug Fixing, even if its title contains test, review, or design.
- QA Integration Test is normally QA / Testing.
- A title containing QA and Review follows the higher-priority testing intent unless the
  description clearly proves it is only a code review.
- Implement, Fix, Update, Convert, and Set flag are normally Development unless the
  description gives stronger evidence for another category.

For each row classify the result as:
- PASS: workbook category is well supported;
- REVIEW: category is plausible but ambiguous or description-dependent;
- FAIL: category conflicts with the task type or clear title/description intent.

Do not force ambiguous rows into PASS or FAIL. Report the evidence and reasoning.

## 3. Dynamic issue scenarios

Do not assume these old IDs will exist in the new run. Discover current IDs and patterns,
and check the historical examples when present.

### Scope and team assignment
- Find PBIs appearing in more than one of FPSO, Foundation, and CSR.
- Find Tasks/Bugs appearing in more than one team.
- For every duplicate, show ID, title, tabs, query URL, and whether each query returned it.
- Historical example: PBI 31346.

### Review classification
- Find PR review, code review, peer review, and review tasks.
- They should normally be Review, not Design / Documentation.
- Historical examples: 35597 (PR-Review) and 35613 (Code Review).

### Demo classification
- Find demo or demonstration tasks.
- They should normally be Review, not Development.
- Historical example: 35527 (Demo Task).

### Bug classification
- Find every row whose Task Type is Bug.
- Every Bug must be Bug Fixing. Report tab, PBI ID, Bug ID, title, description, and category
  for every exception.

### Team leakage and missing children
- Dynamically compare all IDs in each team tab with that team query.
- Do not rely on the historical Issue 4 list; use it only as examples.


## 4. Workbook consistency

Check that:
- All is the combined view of the three team tabs;
- team values match worksheet names;
- PBIs with no returned child still appear as PBI-only rows;
- no Task/Bug is reported under a PBI absent from the same query scope;
- IDs, titles, and links are not malformed;
- blank State, category, PBI ID, or Task ID values are explained.

## Required output

Return:
1. File checked
2. Query inputs and execution status
3. Workbook structure
4. Query scope results
5. Categorization results
6. Dynamic issue scenarios
7. Failures ordered High, Medium, Low
8. Warnings and limitations
9. Overall verdict: PASS, FAIL, or INCONCLUSIVE

Use tables. For every failure or REVIEW item include:

| Sheet tab | PBI ID | PBI title | Task/Bug ID | Type | Task title | Workbook category | Expected category | Evidence/reason |
|---|---:|---|---:|---|---|---|---|---|

For query-scope failures also include:

| Team | Workbook ID | Returned by exact query? | Query URL | Result |
|---|---:|---|---|---|

Never modify data, silently correct categories, create queries, or expose the PAT.
```
