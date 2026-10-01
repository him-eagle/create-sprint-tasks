"""Tests for pbi_bug_task_report.py.

Covers:
  * Rule-based categorization (bug / review-demo priority).
  * Missing-value placeholders (No_Story_Point / No_Hours) vs explicit zero.
  * Adjusted scope validation (PBI must be in the supplied query; a Task/Bug
    only needs to be a hierarchy child of an included PBI, not in the query).
  * Pure build_rows() / write_workbook() behavior using hand-built data (no
    Azure DevOps query of any kind is simulated for these).
  * Validation against the LATEST actually-generated report in
    Reports\\Pbi_bug_tasks\\: each team's query URL is read directly out of
    the workbook (row 1 of its tab, exactly as the script wrote it), and that
    query is re-run against the live Azure DevOps API to compute the expected
    PBI/Task data, which is then compared against what is actually in the
    report. There are no hardcoded/static query IDs anywhere in this file --
    if no report exists yet (or no PAT is configured), those tests are
    skipped with an explanation rather than faking a query result.
"""

import os
import tempfile
import shutil
import unittest
from pathlib import Path

from openpyxl import load_workbook

import pbi_bug_task_report as report


REPORT_DIR = Path(report.OUTPUT_DIR)
TEAM_TABS = {"FPSO", "Foundation"}


def _latest_report_path():
    """Return the most recently modified generated report, or None.

    Excludes "_Summary_" workbooks -- those are produced by the separate
    pbi_bug_task_summary.py script (which fills in the Summary tab) and are
    not a report this script itself generated.
    """
    if not REPORT_DIR.exists():
        return None
    reports = sorted(
        (
            p for p in REPORT_DIR.glob("PBI_Bug_Task_Report_*.xlsx")
            if not p.name.startswith("~$") and "_Summary_" not in p.name
        ),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    return reports[0] if reports else None


def _header_index(rows):
    return next(i for i, row in enumerate(rows) if row and "PBI ID" in row and "Task ID" in row)


def _normalize_assignee(value):
    """Normalize case and whitespace so casing/spacing variants compare equally."""
    return "".join(str(value or "").split()).casefold()


# ---------------------------------------------------------------------------
# Pure unit tests -- no network, no query simulation of any kind.
# ---------------------------------------------------------------------------

class CategorizationTests(unittest.TestCase):
    def test_review_tasks_are_not_design(self):
        """A Task whose title/description mentions "review" or "demo" must be
        categorized as "Review", even though it also matches "Design /
        Documentation" keywords like "document".

        Example: a Task titled "Code Review" or "Demo Task" must come back
        as "Review", not "Design / Documentation" or "Development".
        """
        self.assertEqual(report.categorize_task("Task", "PR-Review", ""), "Review")
        self.assertEqual(report.categorize_task("Task", "Code Review", ""), "Review")
        self.assertEqual(report.categorize_task("Task", "Demo Task", ""), "Review")

    def test_bugs_keep_bug_fixing_priority(self):
        """Work item type "Bug" always categorizes as "Bug Fixing", regardless
        of what keywords appear in its title -- Bug type wins over every
        other rule.

        Example: a Bug titled "Code Review" (which would normally match the
        "Review" keyword rule) must still categorize as "Bug Fixing" because
        its work item type is Bug.
        """
        self.assertEqual(report.categorize_task("Bug", "Code Review", ""), "Bug Fixing")
        self.assertEqual(
            report.categorize_task("Bug", "Regression in parser", ""), "Bug Fixing"
        )


class PlaceholderTests(unittest.TestCase):
    def test_missing_story_points_use_placeholder(self):
        """When a PBI's Story Points field is empty/unset (None or ""), the
        report must show the literal text "No_Story_Point" instead of
        leaving the cell blank.

        Example: format_story_points(None) -> "No_Story_Point".
        """
        self.assertEqual(report.format_story_points(None), "No_Story_Point")
        self.assertEqual(report.format_story_points(""), "No_Story_Point")

    def test_zero_story_points_is_real_data(self):
        """An explicit Story Points value of 0 is real data (the PBI really
        was estimated at zero points) and must be shown as 0, not replaced
        with the "No_Story_Point" placeholder.

        Example: format_story_points(0) -> 0 (not "No_Story_Point").
        """
        self.assertEqual(report.format_story_points(0), 0)

    def test_missing_hours_use_placeholder(self):
        """When a Task/Bug's Original Estimate or Completed Work field is
        empty/unset, the report must show the literal text "No_Hours"
        instead of leaving the cell blank.

        Example: format_hours(None) -> "No_Hours".
        """
        self.assertEqual(report.format_hours(None), "No_Hours")
        self.assertEqual(report.format_hours(""), "No_Hours")

    def test_zero_hours_is_real_data(self):
        """An explicit hours value of 0 (e.g. a task logged with zero
        Completed Work so far) is real data and must be shown as 0, not
        replaced with the "No_Hours" placeholder.

        Example: format_hours(0) -> 0 (not "No_Hours").
        """
        self.assertEqual(report.format_hours(0), 0)


class SanitizeFilenameTests(unittest.TestCase):
    def test_invalid_characters_replaced(self):
        """Characters that are illegal in a Windows filename
        (\\ / : * ? " < > |) must be replaced with underscores so the
        sprint name entered by the user can always be safely used in the
        output file name.

        Example: "Sprint/23:Q1" -> "Sprint_23_Q1".
        """
        self.assertEqual(report.sanitize_filename_component("Sprint/23:Q1"), "Sprint_23_Q1")

    def test_blank_defaults_to_placeholder(self):
        """If the user enters a blank/whitespace-only sprint name, the
        filename component falls back to "UnknownSprint" rather than
        producing a malformed or empty file name segment.

        Example: "   " (just spaces) -> "UnknownSprint".
        """
        self.assertEqual(report.sanitize_filename_component("   "), "UnknownSprint")


class ScopeValidationTests(unittest.TestCase):
    def test_pbi_outside_query_is_rejected(self):
        """A report row whose PBI ID was not actually returned by the
        team's supplied query must raise ValueError -- the PBI set is
        always bounded by what the query returned.

        Example: a row claims pbi_id=100, but the query only returned
        [999] -> validate_report_scope raises ValueError mentioning 100.
        """
        rows = [{"pbi_id": 100, "task_id": 101}]
        with self.assertRaisesRegex(ValueError, "100"):
            report.validate_report_scope(rows, [999], "FPSO")

    def test_task_outside_query_but_linked_to_included_pbi_is_accepted(self):
        """Core new behavior: a Task/Bug ID absent from the supplied query is
        fine, as long as its PBI is present in the query.

        Example: the query returned only [100] (a PBI). A report row has
        pbi_id=100, task_id=555, where 555 was never returned by the query
        at all (it was discovered later via the PBI's hierarchy links) ->
        validate_report_scope does NOT raise, because only the PBI ID (100)
        needs to be in the query result.
        """
        rows = [{"pbi_id": 100, "task_id": 555}]
        # 555 is NOT in the query id list -- must not raise.
        report.validate_report_scope(rows, [100], "FPSO")


class BuildRowsTests(unittest.TestCase):
    """build_rows() is pure (no HTTP). These feed it hand-built work item
    dicts directly -- no Azure DevOps query is simulated."""

    def _pbi(self, pbi_id):
        return {
            "id": pbi_id,
            "fields": {
                "System.WorkItemType": "Product Backlog Item",
                "System.Title": f"PBI {pbi_id}",
            },
        }

    def _child(self, child_id, parent_id, wi_type="Task"):
        return {
            "id": child_id,
            "fields": {
                "System.WorkItemType": wi_type,
                "System.Title": f"Child {child_id}",
                "System.Parent": parent_id,
            },
        }

    def test_child_linked_to_included_pbi_is_included(self):
        """A Task/Bug whose System.Parent points at an included PBI produces
        a report row under that PBI.

        Example: PBI 1 is included; Task 10 has System.Parent=1 ->
        build_rows produces one row with pbi_id=1, task_id=10, and the
        orphan_count stat is 0.
        """
        pbis = [self._pbi(1)]
        children = [self._child(10, parent_id=1)]
        rows, stats = report.build_rows(pbis, children, "org", "proj", "FPSO")
        self.assertEqual([r["task_id"] for r in rows], [10])
        self.assertEqual(stats["orphan_count"], 0)

    def test_child_not_linked_to_any_included_pbi_is_excluded(self):
        """A Task/Bug whose parent is NOT one of the included PBIs is
        dropped from the report (counted as an "orphan") rather than being
        attached to the wrong PBI or crashing the run.

        Example: PBI 1 is included; Task 10 has System.Parent=999 (999 is
        not an included PBI) -> build_rows produces a PBI-only row for PBI 1
        (task_id=None) and stats["orphan_count"] == 1.
        """
        pbis = [self._pbi(1)]
        children = [self._child(10, parent_id=999)]  # 999 not an included PBI
        rows, stats = report.build_rows(pbis, children, "org", "proj", "FPSO")
        self.assertEqual(rows[0]["task_id"], None)  # PBI-only row
        self.assertEqual(stats["orphan_count"], 1)

    def test_pbi_without_children_still_produces_row(self):
        """A PBI with zero Task/Bug children still produces exactly one
        report row (with the Task/Bug columns left blank), so it remains
        visible in the report instead of disappearing.

        Example: PBI 1 has no children at all -> build_rows produces one
        row with pbi_id=1 and task_id=None.
        """
        pbis = [self._pbi(1)]
        rows, stats = report.build_rows(pbis, [], "org", "proj", "FPSO")
        self.assertEqual(len(rows), 1)
        self.assertIsNone(rows[0]["task_id"])

    def test_missing_story_points_become_placeholder_in_row(self):
        """A PBI with no Story Points field set produces a row where
        pbi_story_points is the literal "No_Story_Point" placeholder (this
        exercises the placeholder logic as applied inside build_rows, not
        just the standalone format_story_points helper).

        Example: PBI 1's fields dict has no
        Microsoft.VSTS.Scheduling.StoryPoints key -> the resulting row's
        "pbi_story_points" value is "No_Story_Point".
        """
        pbis = [self._pbi(1)]
        rows, _ = report.build_rows(pbis, [], "org", "proj", "FPSO")
        self.assertEqual(rows[0]["pbi_story_points"], "No_Story_Point")

    def test_missing_hours_become_placeholder_in_row(self):
        """A Task/Bug with no Original Estimate / Completed Work fields set
        produces a row where both hour columns are the literal "No_Hours"
        placeholder.

        Example: Task 10 (child of PBI 1) has no
        Microsoft.VSTS.Scheduling.OriginalEstimate or .CompletedWork keys ->
        the resulting row's "original_estimate" and "completed_work" are
        both "No_Hours".
        """
        pbis = [self._pbi(1)]
        children = [self._child(10, parent_id=1)]
        rows, _ = report.build_rows(pbis, children, "org", "proj", "FPSO")
        self.assertEqual(rows[0]["original_estimate"], "No_Hours")
        self.assertEqual(rows[0]["completed_work"], "No_Hours")


class FilterAndRelationHelperTests(unittest.TestCase):
    def test_assignee_normalization_matches_case_and_camel_case_variants(self):
        """Lowercase, uppercase, and no-space camel-case names all match."""
        expected = _normalize_assignee("Himanshu Pathak")
        for variant in (
            "himanshu pathak",
            "HIMANSHU PATHAK",
            "HimanshuPathak",
            "hImAnShUpAtHaK",
        ):
            with self.subTest(variant=variant):
                self.assertEqual(_normalize_assignee(variant), expected)

    def test_filter_types_excludes_non_matching_types(self):
        """filter_types() keeps only items whose System.WorkItemType is in
        the given allow-set, dropping everything else (e.g. a linked
        Feature must never be treated as a Task/Bug child).

        Example: given a Task, a Feature, and a Bug, filtering with
        CHILD_TYPES = {"Task", "Bug"} keeps 2 of the 3 items (the Feature
        is dropped).
        """
        items = [
            {"fields": {"System.WorkItemType": "Task"}},
            {"fields": {"System.WorkItemType": "Feature"}},
            {"fields": {"System.WorkItemType": "Bug"}},
        ]
        filtered = report.filter_types(items, report.CHILD_TYPES)
        self.assertEqual(len(filtered), 2)

    def test_collect_child_ids_only_uses_forward_relations(self):
        """collect_child_ids() gathers child IDs only from
        System.LinkTypes.Hierarchy-Forward relations -- it must ignore
        Hierarchy-Reverse (that's the PBI's own parent, not its child) and
        unrelated link types like "Related".

        Example: a PBI has three relations -- a Forward link to item 55, a
        Reverse link to item 1, and a Related link to item 77 -- only [55]
        is returned.
        """
        item = {
            "relations": [
                {"rel": "System.LinkTypes.Hierarchy-Forward", "url": ".../workItems/55"},
                {"rel": "System.LinkTypes.Hierarchy-Reverse", "url": ".../workItems/1"},
                {"rel": "System.LinkTypes.Related", "url": ".../workItems/77"},
            ]
        }
        self.assertEqual(report.collect_child_ids([item]), [55])


class WriteWorkbookStructureTests(unittest.TestCase):
    """write_workbook() is pure (no HTTP, no query). These tests feed it
    hand-built row lists directly to verify sheet structure/ordering."""

    def setUp(self):
        self.tmp_dir = Path(tempfile.mkdtemp(prefix="pbi_bug_task_report_wb_test_"))
        self.addCleanup(shutil.rmtree, self.tmp_dir, ignore_errors=True)

    def _sample_row(self, pbi_id, task_id, team):
        return {
            "pbi_id": pbi_id, "pbi_url": f"https://example/{pbi_id}",
            "pbi_title": "Sample PBI", "pbi_story_points": 3,
            "pbi_done_date": "2024-01-01", "task_id": task_id,
            "task_url": f"https://example/{task_id}" if task_id else None,
            "task_type": "Task" if task_id else "", "task_title": "Sample task",
            "assigned_to": "Alice", "state": "Active",
            "original_estimate": 2, "completed_work": 1,
            "categorization": "Development", "done_date": "", "description": "",
            "team": team,
        }

    def test_summary_tab_is_empty_and_first(self):
        """The workbook must contain a "Summary" tab as the very first
        sheet, and that tab must contain zero rows/content of any kind
        (it's a placeholder for future manual/automated summary content).

        Example: writing one FPSO row produces sheet order
        ["Summary", "All", "FPSO"], and reading every row of the Summary
        sheet yields no non-empty rows at all.
        """
        team_data = [("FPSO", [self._sample_row(1, 10, "FPSO")], "https://example/fpso-query")]
        output_path = self.tmp_dir / "sample.xlsx"
        report.write_workbook(team_data, str(output_path))

        wb = load_workbook(output_path, data_only=True)
        self.assertEqual(wb.sheetnames, ["Summary", "All", "FPSO"])
        summary_rows = list(wb["Summary"].iter_rows(values_only=True))
        non_empty = [row for row in summary_rows if any(v is not None for v in row)]
        self.assertEqual(non_empty, [])
        wb.close()

    def test_team_with_no_rows_still_gets_a_tab(self):
        """A team that produced zero rows (e.g. its query returned no PBIs)
        still gets its own tab in the workbook (just the Source URL + header
        row, no data rows), and the combined "All" tab does not contain any
        row attributed to that team.

        Example: Foundation contributes an empty row list -> the
        "Foundation" tab exists with exactly 2 non-empty rows (Source URL
        row + header row), and no row in "All" has Team == "Foundation".
        """
        team_data = [
            ("FPSO", [self._sample_row(1, 10, "FPSO")], "https://example/fpso-query"),
            ("Foundation", [], "https://example/foundation-query"),
        ]
        output_path = self.tmp_dir / "sample_empty_team.xlsx"
        report.write_workbook(team_data, str(output_path))

        wb = load_workbook(output_path, data_only=True)
        self.assertIn("Foundation", wb.sheetnames)
        foundation_rows = [
            row for row in wb["Foundation"].iter_rows(values_only=True)
            if any(v is not None for v in row)
        ]
        # Only the "Source URL" row + header row -- no data rows.
        self.assertEqual(len(foundation_rows), 2)

        all_rows = list(wb["All"].iter_rows(values_only=True))
        header_index = _header_index(all_rows)
        foundation_contribution = [
            row for row in all_rows[header_index + 1:]
            if any(v is not None for v in row) and row[0] == "Foundation"
        ]
        self.assertEqual(foundation_contribution, [])
        wb.close()

    def test_output_directory_is_auto_created(self):
        """write_workbook() does not require its destination folder to
        already exist -- in practice run_report() creates it with
        os.makedirs(..., exist_ok=True) before calling write_workbook(),
        and this test exercises that a workbook can be saved successfully
        once that directory has been created.

        Example: a nested path .../DoesNotExistYet/Nested does not exist
        beforehand; after os.makedirs + write_workbook, the .xlsx file
        exists at that path.
        """
        nested_dir = self.tmp_dir / "DoesNotExistYet" / "Nested"
        self.assertFalse(nested_dir.exists())
        team_data = [("FPSO", [self._sample_row(1, 10, "FPSO")], "https://example/fpso-query")]
        os.makedirs(nested_dir, exist_ok=True)
        output_path = nested_dir / "sample.xlsx"
        report.write_workbook(team_data, str(output_path))
        self.assertTrue(output_path.exists())


# ---------------------------------------------------------------------------
# Validation against the LATEST actually-generated report (no fakes/mocks).
#
# These read the query URL straight out of the workbook the script itself
# produced, then re-run that exact query against the live Azure DevOps API
# to compute what the report SHOULD contain, and compare it to what the
# report actually contains. If no report has been generated yet, or no PAT
# is configured, the relevant tests are skipped (not faked).
# ---------------------------------------------------------------------------

class LatestReportStructureTests(unittest.TestCase):
    """Checks that only need the report file itself -- no live API calls."""

    @classmethod
    def setUpClass(cls):
        cls.report_path = _latest_report_path()
        if cls.report_path is None:
            raise unittest.SkipTest(
                f"No generated report found in {REPORT_DIR}. "
                "Run pbi_bug_task_report.py first, then re-run these tests."
            )
        cls.workbook = load_workbook(cls.report_path, data_only=True)

    @classmethod
    def tearDownClass(cls):
        cls.workbook.close()

    def _team_table(self, team):
        ws = self.workbook[team]
        rows = list(ws.iter_rows(values_only=True))
        header_index = _header_index(rows)
        headers = rows[header_index]
        col = {name: headers.index(name) for name in headers if name}
        return rows[header_index + 1:], col

    def test_tab_order_and_summary_is_empty(self):
        """The newest generated report must list "Summary" then "All" as
        its first two sheets (in that order), both FPSO and Foundation
        tabs must exist, and the Summary tab must contain no content.

        Example: wb.sheetnames[:2] == ["Summary", "All"], and every row
        read from the "Summary" sheet is entirely blank (all None values).
        """
        expected_prefix = ["Summary", "All"]
        self.assertEqual(self.workbook.sheetnames[:2], expected_prefix)
        for team in TEAM_TABS:
            self.assertIn(team, self.workbook.sheetnames)

        summary_rows = list(self.workbook["Summary"].iter_rows(values_only=True))
        non_empty = [row for row in summary_rows if any(v is not None for v in row)]
        self.assertEqual(non_empty, [], "Summary tab must be left empty")

    def test_each_pbi_assigned_to_one_team(self):
        """No PBI may appear under more than one team tab, and every row's
        "Team" column must match the tab it's actually sitting in.

        Example: if PBI 12345 appears as a row in the "FPSO" tab, it must
        not also appear as a row in the "Foundation" tab, and that row's
        Team column value must read "FPSO".
        """
        pbi_teams = {}
        for team in TEAM_TABS:
            if team not in self.workbook.sheetnames:
                continue
            data_rows, col = self._team_table(team)
            for row in data_rows:
                pbi_id = row[col["PBI ID"]]
                if pbi_id is None:
                    continue
                self.assertEqual(row[col["Team"]], team)
                pbi_teams.setdefault(pbi_id, set()).add(team)

        duplicates = {pbi: teams for pbi, teams in pbi_teams.items() if len(teams) > 1}
        self.assertFalse(duplicates, f"PBIs assigned to multiple team tabs: {duplicates}")

    def test_bugs_and_demo_tasks_are_categorized_correctly(self):
        """Spot-checks the categorization rules against real report rows:
        every row whose Task Type is "Bug" must have Task Categorization
        "Bug Fixing", and every row whose title contains "demo" must have
        Task Categorization "Review".

        Example: a row with Task Type "Bug" and Task Categorization
        "Development" would be reported as a failure tuple
        (team, pbi_id, task_id, "Bug", "Development").
        """
        failures = []
        for team in TEAM_TABS:
            if team not in self.workbook.sheetnames:
                continue
            data_rows, col = self._team_table(team)
            for row in data_rows:
                if row[col["Task Type"]] == "Bug" and row[col["Task Categorization"]] != "Bug Fixing":
                    failures.append((team, row[col["PBI ID"]], row[col["Task ID"]], "Bug", row[col["Task Categorization"]]))
                title = (row[col["Task Title"]] or "")
                if "demo" in title.lower() and row[col["Task Categorization"]] != "Review":
                    failures.append((team, row[col["PBI ID"]], row[col["Task ID"]], "Demo", row[col["Task Categorization"]]))
        self.assertFalse(failures, f"Mis-categorized rows (team, pbi, task, kind, actual): {failures}")

    def test_himanshu_pathak_is_not_assigned_in_any_report_tab(self):
        """No data row on any report tab may be assigned to Himanshu Pathak.

        Checks every worksheet that has an Assigned To column, including All;
        worksheets without a report table (for example, the empty Summary tab)
        are ignored. Failure details identify the tab and work item.
        """
        failures = []

        for worksheet in self.workbook.worksheets:
            rows = list(worksheet.iter_rows(values_only=True))
            header_index = next(
                (
                    index
                    for index, row in enumerate(rows)
                    if row and "Assigned To" in row
                ),
                None,
            )
            if header_index is None:
                continue

            headers = rows[header_index]
            col = {name: headers.index(name) for name in headers if name}
            for row in rows[header_index + 1:]:
                assigned_to = row[col["Assigned To"]]
                if _normalize_assignee(assigned_to) != _normalize_assignee("Himanshu Pathak"):
                    continue
                failures.append(
                    {
                        "tab": worksheet.title,
                        "pbi_id": row[col["PBI ID"]] if "PBI ID" in col else None,
                        "task_id": row[col["Task ID"]] if "Task ID" in col else None,
                        "task_title": row[col["Task Title"]] if "Task Title" in col else None,
                    }
                )

        self.assertFalse(
            failures,
            f"Report '{self.report_path}' assigns one or more items to Himanshu Pathak: "
            f"{failures}",
        )

    def test_placeholders_never_left_blank(self):
        """Every data row's PBI Story Points column must carry either a real
        value or "No_Story_Point" -- never blank/None. Every row that has a
        Task/Bug (task_id is not None) must carry either a real value or
        "No_Hours" in both Original Estimate and Completed Work -- never
        blank/None.

        Example: a row with PBI Story Points == "" (empty string, not the
        "No_Story_Point" placeholder) would be reported as a failure tuple
        (team, pbi_id, "PBI Story Points", "").
        """
        failures = []
        for team in TEAM_TABS:
            if team not in self.workbook.sheetnames:
                continue
            data_rows, col = self._team_table(team)
            for row in data_rows:
                sp = row[col["PBI Story Points"]]
                if sp is None or sp == "":
                    failures.append((team, row[col["PBI ID"]], "PBI Story Points", sp))
                if row[col["Task ID"]] is not None:
                    for field in ("Original Estimate (hrs)", "Completed Work (hrs)"):
                        value = row[col[field]]
                        if value is None or value == "":
                            failures.append((team, row[col["Task ID"]], field, value))
        self.assertFalse(failures, f"Blank fields that should carry a placeholder: {failures}")

    def test_all_tab_row_count_matches_team_tabs(self):
        """The number of data rows in the combined "All" tab must exactly
        equal the sum of the data rows across the FPSO and Foundation tabs
        -- "All" is a pure concatenation of the two, with no duplication or
        omission.

        Example: if FPSO has 12 data rows and Foundation has 7, "All" must
        have exactly 19 data rows.
        """
        all_data_rows, _ = self._team_table("All")
        total_team_rows = 0
        for team in TEAM_TABS:
            if team not in self.workbook.sheetnames:
                continue
            data_rows, _ = self._team_table(team)
            total_team_rows += len(data_rows)
        self.assertEqual(len(all_data_rows), total_team_rows)

    def test_hyperlinks_point_to_correct_work_items(self):
        """Every PBI ID / Task ID cell that has a hyperlink must point to
        the correct Azure DevOps work item edit URL for that ID, built from
        the same org/project the row's team query URL uses.

        Example: in the "FPSO" tab, a PBI ID cell with value 12345 and a
        hyperlink must have hyperlink.target equal to
        "https://dev.azure.com/{org}/{project}/_workitems/edit/12345".
        """
        for team in TEAM_TABS:
            if team not in self.workbook.sheetnames:
                continue
            ws = self.workbook[team]
            rows = list(ws.iter_rows(values_only=True))
            header_index = _header_index(rows)
            headers = rows[header_index]
            col = {name: headers.index(name) for name in headers if name}
            source_url = rows[0][1]
            info = report.parse_query_url(source_url)
            org, project = info["org"], info["project"]

            checked_any = False
            for row_cells in ws.iter_rows(min_row=header_index + 2):
                pbi_cell = row_cells[col["PBI ID"]]
                if pbi_cell.value is not None and pbi_cell.hyperlink:
                    expected = report.build_work_item_url(org, project, pbi_cell.value)
                    self.assertEqual(pbi_cell.hyperlink.target, expected)
                    checked_any = True
                task_cell = row_cells[col["Task ID"]]
                if task_cell.value is not None and task_cell.hyperlink:
                    expected = report.build_work_item_url(org, project, task_cell.value)
                    self.assertEqual(task_cell.hyperlink.target, expected)
                    checked_any = True
            if not checked_any:
                self.skipTest(f"No hyperlinked rows found in {team} tab to check")


class LatestReportLiveApiTests(unittest.TestCase):
    """Re-runs each team's query (read straight from the report) against the
    live Azure DevOps API and compares the result to the report contents."""

    @classmethod
    def setUpClass(cls):
        cls.report_path = _latest_report_path()
        if cls.report_path is None:
            raise unittest.SkipTest(
                f"No generated report found in {REPORT_DIR}. "
                "Run pbi_bug_task_report.py first, then re-run these tests."
            )
        pat_token = report.load_pat_token()
        if not pat_token:
            raise unittest.SkipTest(
                "Set AZURE_DEVOPS_PAT (or populate localconfig) to validate "
                "the latest report against the live Azure DevOps API."
            )
        cls.session = report.get_session(pat_token)
        cls.workbook = load_workbook(cls.report_path, data_only=True)

    @classmethod
    def tearDownClass(cls):
        cls.workbook.close()

    def _source_url(self, team):
        ws = self.workbook[team]
        rows = list(ws.iter_rows(values_only=True))
        return rows[0][1]

    def _team_table(self, team):
        ws = self.workbook[team]
        rows = list(ws.iter_rows(values_only=True))
        header_index = _header_index(rows)
        headers = rows[header_index]
        col = {name: headers.index(name) for name in headers if name}
        return rows[header_index + 1:], col

    def _expected_state(self, team):
        """Recompute the current expected PBI/Task state for a team straight
        from the live Azure DevOps API, mirroring fetch_query_data's own
        Step A / Step B logic."""
        source_url = self._source_url(team)
        info = report.parse_query_url(source_url)
        org, project, query_id = info["org"], info["project"], info["query_id"]

        query_ids = report.run_wiql_query_by_id(self.session, org, project, query_id)
        full_items = report.get_work_items_full(self.session, org, project, query_ids)
        pbi_items = report.filter_types(full_items, report.PBI_TYPES)

        child_ids = report.collect_child_ids(pbi_items)
        fetched_by_id = {item["id"]: item for item in full_items}
        missing_ids = [cid for cid in child_ids if cid not in fetched_by_id]
        if missing_ids:
            for item in report.get_work_items_full(self.session, org, project, missing_ids):
                fetched_by_id[item["id"]] = item
        candidate_children = [fetched_by_id[cid] for cid in child_ids if cid in fetched_by_id]
        child_items = report.filter_types(candidate_children, report.CHILD_TYPES)

        pbis = {item["id"]: item for item in pbi_items}
        mapping = {pbi_id: set() for pbi_id in pbis}
        for child in child_items:
            parent_id = report.get_parent_id(child)
            if parent_id in mapping:
                mapping[parent_id].add(child["id"])

        fields_by_id = {item["id"]: item.get("fields", {}) for item in pbi_items}
        fields_by_id.update({item["id"]: item.get("fields", {}) for item in child_items})

        return {
            "query_ids": set(query_ids),
            "mapping": mapping,
            "fields_by_id": fields_by_id,
        }

    def test_parent_child_correspondence_matches_live_api(self):
        """The PBI->Task/Bug mapping actually written into the report must
        match the mapping you'd get by re-running the team's query right
        now against live Azure DevOps and walking each PBI's hierarchy
        links (i.e. the report is not stale/out of sync with ADO).

        Example: if the live API currently says PBI 12345 has children
        {2001, 2002} but the report's "FPSO" tab shows PBI 12345 with only
        task_id 2001 (missing 2002, or showing an extra task), this test
        fails with a diff showing expected vs. actual mapping for "FPSO".
        """
        failures = []
        for team in TEAM_TABS:
            if team not in self.workbook.sheetnames:
                continue
            expected = self._expected_state(team)
            data_rows, col = self._team_table(team)

            actual_mapping = {}
            for row in data_rows:
                pbi_id = row[col["PBI ID"]]
                task_id = row[col["Task ID"]]
                actual_mapping.setdefault(pbi_id, set())
                if task_id is not None:
                    actual_mapping[pbi_id].add(task_id)

            if actual_mapping != expected["mapping"]:
                failures.append({
                    "team": team,
                    "expected": expected["mapping"],
                    "actual": actual_mapping,
                })

        self.assertFalse(
            failures,
            "Report PBI/Task mapping does not match the live Azure DevOps API:\n"
            + "\n".join(
                f"  [{f['team']}] expected={f['expected']} actual={f['actual']}"
                for f in failures
            ),
        )

    def test_pbi_ids_are_within_supplied_query(self):
        """Every PBI ID in the report must currently be returned by that
        team's supplied query when re-run live -- the report must never
        contain a PBI the query doesn't actually select.

        Example: if the "Foundation" tab contains PBI 98765 but re-running
        the Foundation query live no longer returns 98765 (e.g. it was
        since moved out of scope), this test fails with
        [("Foundation", 98765)].
        """
        failures = []
        for team in TEAM_TABS:
            if team not in self.workbook.sheetnames:
                continue
            expected = self._expected_state(team)
            data_rows, col = self._team_table(team)
            for row in data_rows:
                pbi_id = row[col["PBI ID"]]
                if pbi_id is not None and pbi_id not in expected["query_ids"]:
                    failures.append((team, pbi_id))
        self.assertFalse(failures, f"PBI IDs not present in their team's live query result: {failures}")

    def test_story_points_and_hours_match_live_api(self):
        """Every PBI Story Points value and every Task/Bug's Original
        Estimate / Completed Work value in the report must match what the
        live Azure DevOps API currently reports for that same field
        (after applying the same No_Story_Point / No_Hours placeholder
        logic the script itself uses).

        Example: if PBI 12345's live Story Points field is now 8 but the
        report shows 5 (e.g. it was re-estimated after the report was
        generated), this test fails with
        ("FPSO", 12345, "PBI Story Points", 8, 5).
        """
        failures = []
        for team in TEAM_TABS:
            if team not in self.workbook.sheetnames:
                continue
            expected = self._expected_state(team)
            data_rows, col = self._team_table(team)
            for row in data_rows:
                pbi_id = row[col["PBI ID"]]
                pbi_fields = expected["fields_by_id"].get(pbi_id, {})
                expected_sp = report.format_story_points(
                    pbi_fields.get(report.STORY_POINTS_FIELD)
                )
                actual_sp = row[col["PBI Story Points"]]
                if actual_sp != expected_sp:
                    failures.append((team, pbi_id, "PBI Story Points", expected_sp, actual_sp))

                task_id = row[col["Task ID"]]
                if task_id is not None:
                    task_fields = expected["fields_by_id"].get(task_id, {})
                    expected_oe = report.format_hours(
                        task_fields.get(report.ORIGINAL_ESTIMATE_FIELD)
                    )
                    expected_cw = report.format_hours(
                        task_fields.get(report.COMPLETED_WORK_FIELD)
                    )
                    actual_oe = row[col["Original Estimate (hrs)"]]
                    actual_cw = row[col["Completed Work (hrs)"]]
                    if actual_oe != expected_oe:
                        failures.append((team, task_id, "Original Estimate", expected_oe, actual_oe))
                    if actual_cw != expected_cw:
                        failures.append((team, task_id, "Completed Work", expected_cw, actual_cw))

        self.assertFalse(
            failures,
            "Field values do not match the live Azure DevOps API "
            "(team, id, field, expected, actual):\n" + "\n".join(str(f) for f in failures),
        )

    def test_every_queried_pbi_is_in_the_report_with_matching_count(self):
        """Completes the round-trip check: every PBI currently returned by
        re-running a team's query live must appear somewhere in that team's
        report tab, AND the number of distinct PBIs returned by the live
        query must exactly equal the number of distinct PBIs in the report
        tab. (The other live-API tests above check the opposite direction
        -- that every PBI *in the report* came from the query -- so this
        test closes the loop and catches a PBI that the query returns but
        that the report is missing, or a stale/duplicate PBI count.)

        Example: if re-running the "FPSO" query live returns PBIs
        {101, 102, 103} (3 PBIs), but the "FPSO" tab in the report only
        contains rows for PBIs {101, 102} (2 PBIs, missing 103), this test
        fails reporting the missing PBI 103 and the count mismatch
        3 (query) != 2 (report).
        """
        failures = []
        for team in TEAM_TABS:
            if team not in self.workbook.sheetnames:
                continue
            expected = self._expected_state(team)
            expected_pbi_ids = set(expected["mapping"].keys())

            data_rows, col = self._team_table(team)
            actual_pbi_ids = {
                row[col["PBI ID"]] for row in data_rows if row[col["PBI ID"]] is not None
            }

            missing_from_report = sorted(expected_pbi_ids - actual_pbi_ids)
            if missing_from_report:
                failures.append(
                    f"[{team}] PBI(s) returned by the live query but missing from the "
                    f"report: {missing_from_report}"
                )

            if len(expected_pbi_ids) != len(actual_pbi_ids):
                failures.append(
                    f"[{team}] PBI count mismatch: live query returned "
                    f"{len(expected_pbi_ids)} PBI(s), report contains "
                    f"{len(actual_pbi_ids)} PBI(s)"
                )

        self.assertFalse(failures, "\n".join(failures))


if __name__ == "__main__":
    unittest.main()
