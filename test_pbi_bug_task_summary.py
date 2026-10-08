"""Unit and live Azure DevOps tests for the flat PBI Summary workbook."""

import os
import re
import shutil
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

import pbi_bug_task_report as report
import pbi_bug_task_summary as summary


REPORT_DIR = Path(summary.OUTPUT_DIR)
TEAM_TABS = ("FPSO", "Foundation")


def _latest_summary_report_path():
    if not REPORT_DIR.exists():
        return None
    reports = sorted(
        (path for path in REPORT_DIR.glob("PBI_Bug_Task_Report_*_Summary_*.xlsx")
         if not path.name.startswith("~$")),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )
    return reports[0] if reports else None


def _source_workbook(path, rows, *, summary_content=False):
    """Create a minimal source workbook for tests, with dynamic PBI IDs."""
    wb = Workbook()
    ws_summary = wb.active
    ws_summary.title = "Summary"
    if summary_content:
        ws_summary["A1"] = "already processed"

    columns = [
        "Team", "PBI ID", "PBI Title", "PBI Story Points", "PBI Done Date",
        "Task ID", "Task Type", "Task Title", "Assigned To", "State",
        "Original Estimate (hrs)", "Completed Work (hrs)",
        "Task Categorization", "Done Date", "Description",
    ]
    for name in ("All", "FPSO", "Foundation"):
        ws = wb.create_sheet(name)
        ws.cell(1, 1, "Source URL:")
        ws.cell(1, 2, f"https://dev.azure.com/org/proj/_queries/query/{name.lower()}/")
        for col, value in enumerate(columns, 1):
            ws.cell(2, col, value)
        relevant = rows if name == "All" else [row for row in rows if row["Team"] == name]
        for row_num, row in enumerate(relevant, 3):
            for col, heading in enumerate(columns, 1):
                ws.cell(row_num, col, row.get(heading))
            if row.get("PBI URL"):
                ws.cell(row_num, 2).hyperlink = row["PBI URL"]
    wb.save(path)
    wb.close()
    return path


def _sample_rows():
    shared = {
        "PBI Title": "Dynamic sample PBI",
        "PBI Story Points": 3,
        "PBI Done Date": "2026-10-01",
        "Assigned To": "Engineer",
        "State": "Done",
        "PBI URL": "https://dev.azure.com/org/proj/_workitems/edit/88001",
    }
    return [
        {**shared, "Team": "FPSO", "PBI ID": 88001, "Task ID": None,
         "Task Type": None, "Task Title": None, "Completed Work (hrs)": None,
         "Task Categorization": None},
        {**shared, "Team": "FPSO", "PBI ID": 88001, "Task ID": 88011,
         "Task Type": "Task", "Task Title": "implement feature",
         "Completed Work (hrs)": 6, "Task Categorization": "Development"},
        {**shared, "Team": "FPSO", "PBI ID": 88001, "Task ID": 88012,
         "Task Type": "Bug", "Task Title": "issue fix",
         "Completed Work (hrs)": 2, "Task Categorization": "Bug Fixing"},
        {**shared, "Team": "FPSO", "PBI ID": 88001, "Task ID": 88013,
         "Task Type": "Task", "Task Title": "verification",
         "Completed Work (hrs)": "No_Hours", "Task Categorization": "QA / Testing"},
        {**shared, "Team": "Foundation", "PBI ID": 88002, "PBI Title": "No child PBI",
         "PBI Story Points": "No_Story_Point", "Task ID": None,
         "Task Type": None, "Task Title": None, "Completed Work (hrs)": None,
         "Task Categorization": None, "PBI URL": "https://dev.azure.com/org/proj/_workitems/edit/88002"},
    ]


def _header_map(ws, header_row=2):
    return {ws.cell(header_row, col).value: col for col in range(1, ws.max_column + 1)
            if ws.cell(header_row, col).value is not None}


def _summary_map(ws):
    return {header: index for index, header in enumerate(summary.BASE_HEADERS, 1)}


def _effort_col(category):
    return 5 + summary.CATEGORY_ORDER.index(category)


def _percentage_col(category):
    return 11 + summary.CATEGORY_ORDER.index(category)


def _parse_work_item_url(url):
    match = re.search(r"dev\.azure\.com/([^/]+)/([^/]+)/_workitems/edit/(\d+)", url or "")
    return match.groups() if match else None


class FilenameAndWorkbookValidationTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.temp_dir, ignore_errors=True)

    def test_filename_prefix_required(self):
        summary.validate_input_filename(self.temp_dir / "PBI_Bug_Task_Report_1.xlsx")
        with self.assertRaisesRegex(ValueError, "PBI_Bug_Task_Report_"):
            summary.validate_input_filename(self.temp_dir / "other.xlsx")

    def test_required_tabs_required(self):
        wb = Workbook()
        wb.active.title = "Summary"
        wb.create_sheet("All")
        with self.assertRaisesRegex(ValueError, "FPSO.*Foundation|Foundation.*FPSO"):
            summary.validate_workbook_structure(wb)
        wb.close()

    def test_summary_must_be_empty(self):
        path = _source_workbook(self.temp_dir / "PBI_Bug_Task_Report_1.xlsx", [], summary_content=True)
        wb = load_workbook(path)
        with self.assertRaisesRegex(ValueError, "already contains content"):
            summary.validate_workbook_structure(wb)
        wb.close()

    def test_valid_workbook_structure(self):
        path = _source_workbook(self.temp_dir / "PBI_Bug_Task_Report_1.xlsx", _sample_rows())
        wb = load_workbook(path)
        summary.validate_workbook_structure(wb)
        wb.close()

    def test_output_name_replaces_existing_timestamp(self):
        result = summary.build_output_filename(
            "PBI_Bug_Task_Report_103_20261001_104744.xlsx", "20261008_120000"
        )
        self.assertEqual(result, "PBI_Bug_Task_Report_103_Summary_20261008_120000.xlsx")


class SummaryCalculationTests(unittest.TestCase):
    def test_rollup_aggregates_hours_and_percentages(self):
        pbi_rows = summary.summarize_pbis(self._load_sample())
        pbi = next(item for item in pbi_rows if item["pbi_id"] == 88001)
        self.assertEqual(pbi["category_hours"]["Development"], 6)
        self.assertEqual(pbi["category_hours"]["Bug Fixing"], 2)
        self.assertEqual(pbi["category_hours"]["QA / Testing"], 0)
        self.assertAlmostEqual(pbi["category_percentages"]["Development"], 75.0)
        self.assertAlmostEqual(sum(pbi["category_percentages"].values()), 100.0)
        self.assertTrue(pbi["missing_hours"])

    def _load_sample(self):
        if not hasattr(self, "temp_dir"):
            self.temp_dir = Path(tempfile.mkdtemp())
            self.addCleanup(shutil.rmtree, self.temp_dir, ignore_errors=True)
        path = _source_workbook(self.temp_dir / "PBI_Bug_Task_Report_1.xlsx", _sample_rows())
        wb = load_workbook(path)
        try:
            return summary.read_all_tab(wb)
        finally:
            wb.close()

    def test_placeholder_missing_zero_and_zero_total_semantics(self):
        self.assertTrue(summary.is_missing_hours(None))
        self.assertTrue(summary.is_missing_hours("No_Hours"))
        self.assertFalse(summary.is_missing_hours(0))
        self.assertEqual(summary.hours_to_number("No_Hours"), 0)
        self.assertEqual(summary.hours_to_number(0), 0)
        rows = [{"pbi_id": 9, "pbi_title": "Empty", "pbi_story_points": 0,
                 "pbi_url": None, "task_id": None, "task_title": None,
                 "completed_work": None, "categorization": None}]
        [empty] = summary.summarize_pbis(rows)
        self.assertEqual(empty["total_hours"], 0)
        self.assertEqual(set(empty["category_hours"].values()), {0.0})
        self.assertEqual(set(empty["category_percentages"].values()), {0.0})

    def test_unexpected_or_blank_category_fails_with_pbi_and_task(self):
        rows = [{"pbi_id": 900, "pbi_title": "P", "pbi_story_points": 1,
                 "pbi_url": None, "task_id": 901, "task_title": "Unclassified",
                 "completed_work": 2, "categorization": ""}]
        with self.assertRaisesRegex(ValueError, "PBI 900.*901.*unexpected"):
            summary.summarize_pbis(rows)

    def test_categories_and_total_are_from_same_task_hours(self):
        rows = [
            {"pbi_id": 5, "pbi_title": "P", "pbi_story_points": 2, "pbi_url": None,
             "task_id": 50, "task_title": "A", "completed_work": 2, "categorization": "Review"},
            {"pbi_id": 5, "pbi_title": "P", "pbi_story_points": 2, "pbi_url": None,
             "task_id": 51, "task_title": "B", "completed_work": 3, "categorization": "Development"},
        ]
        [pbi] = summary.summarize_pbis(rows)
        self.assertEqual(sum(pbi["category_hours"].values()), pbi["total_hours"])


class WriteSummarySheetTests(unittest.TestCase):
    def _create(self, rows):
        wb = Workbook()
        wb.active.title = "Summary"
        result = summary.summarize_pbis(rows)
        summary.write_summary_sheet(wb, result)
        return wb, result

    def test_grouped_headers_columns_and_flat_one_row_per_pbi(self):
        wb, pbi = self._create(self._load_source())
        ws = wb["Summary"]
        self.assertEqual((ws["E1"].value, ws["K1"].value), ("Effort", "Percentage"))
        self.assertEqual([ws.cell(2, col).value for col in range(1, 17)], summary.SUMMARY_HEADERS)
        self.assertEqual(ws.max_row, 2 + len(pbi))
        self.assertEqual([str(r) for r in ws.merged_cells.ranges], ["E1:J1", "K1:P1"])
        self.assertEqual(len(ws._charts), 0)
        self.assertFalse(any(dim.hidden for dim in ws.column_dimensions.values()))
        wb.close()

    def _load_source(self):
        if not hasattr(self, "temp_dir"):
            self.temp_dir = Path(tempfile.mkdtemp())
            self.addCleanup(shutil.rmtree, self.temp_dir, ignore_errors=True)
        path = _source_workbook(self.temp_dir / "PBI_Bug_Task_Report_2.xlsx", _sample_rows())
        wb = load_workbook(path)
        try:
            return summary.read_all_tab(wb)
        finally:
            wb.close()

    def test_summary_row_values_hyperlink_and_missing_hours(self):
        wb, _ = self._create(self._load_source())
        ws = wb["Summary"]
        cols = _summary_map(ws)
        row = 3
        self.assertEqual(ws.cell(row, cols["PBI ID"]).value, 88001)
        self.assertEqual(ws.cell(row, cols["PBI Title"]).value, "Dynamic sample PBI")
        self.assertEqual(ws.cell(row, cols["PBI Story Points"]).value, 3)
        self.assertEqual(ws.cell(row, cols["Total Hours"]).value, "8 (some tasks/bugs missing hours)")
        self.assertEqual(ws.cell(row, _effort_col("Development")).value, 6)
        self.assertEqual(ws.cell(row, _effort_col("Bug Fixing")).value, 2)
        self.assertEqual(ws.cell(row, _effort_col("QA / Testing")).value, 0)
        self.assertAlmostEqual(ws.cell(row, _percentage_col("Development")).value, 75)
        self.assertEqual(ws.cell(row, 1).hyperlink.target,
                         "https://dev.azure.com/org/proj/_workitems/edit/88001")
        self.assertEqual(ws.cell(row, 4).fill.fgColor.rgb, summary.YELLOW_FILL.fgColor.rgb)
        wb.close()

    def test_percentage_values_and_number_formats(self):
        wb, _ = self._create(self._load_source())
        ws = wb["Summary"]
        row = 3
        self.assertAlmostEqual(ws.cell(row, _percentage_col("Bug Fixing")).value, 25.0)
        self.assertAlmostEqual(ws.cell(row, _percentage_col("Development")).value, 75.0)
        self.assertEqual(ws.cell(row, _percentage_col("Bug Fixing")).number_format, '0.0"%"')
        wb.close()

    def test_reference_colors_and_visible_borders(self):
        wb, _ = self._create(self._load_source())
        ws = wb["Summary"]
        self.assertTrue(ws["A2"].font.bold)
        for row in range(1, 3):
            for column in range(1, 5):
                cell = ws.cell(row, column)
                self.assertEqual(cell.fill.fgColor.rgb, summary.DARK_BLUE_HEADER_ARGB)
                self.assertEqual(cell.font.color.rgb, summary.HEADER_FONT_COLOR)
                self.assertTrue(cell.font.bold)
            for column in range(5, 11):
                cell = ws.cell(row, column)
                self.assertEqual(cell.fill.fgColor.rgb, summary.EFFORT_HEADER_ARGB)
                self.assertEqual(cell.font.color.rgb, summary.HEADER_FONT_COLOR)
                self.assertTrue(cell.font.bold)
            for column in range(11, 17):
                cell = ws.cell(row, column)
                self.assertEqual(cell.fill.fgColor.rgb, summary.DARK_BLUE_HEADER_ARGB)
                self.assertEqual(cell.font.color.rgb, summary.HEADER_FONT_COLOR)
                self.assertTrue(cell.font.bold)
        self.assertIsNone(ws["A3"].fill.fill_type)
        for address in ("A2", "P2", "A3", "P3"):
            cell = ws[address]
            for edge in (cell.border.left, cell.border.right, cell.border.top, cell.border.bottom):
                self.assertEqual(edge.style, "thin", f"{address} is missing a visible border")
            self.assertEqual(cell.border.top.color.rgb, summary.VISIBLE_BORDER_COLOR)
        for address, expected_left, expected_right in (
            ("E1", "thin", None), ("J1", None, "thin"),
            ("K1", "thin", None), ("P1", None, "thin"),
        ):
            cell = ws[address]
            self.assertEqual(cell.border.left.style, expected_left)
            self.assertEqual(cell.border.right.style, expected_right)
            self.assertEqual(cell.border.top.style, "thin")
            self.assertEqual(cell.border.bottom.style, "thin")
        self.assertEqual(ws["A3"].font.color.rgb, summary.HYPERLINK_FONT_COLOR)
        self.assertEqual(ws["C4"].fill.fgColor.rgb, "FFFFFF00")
        wb.close()

    def test_unknown_category_reports_pbi_and_task(self):
        rows = [{"pbi_id": 42, "pbi_title": "P", "pbi_story_points": 1,
                 "pbi_url": None, "task_id": 420, "task_title": "bad cat",
                 "completed_work": 3, "categorization": "Unmapped"}]
        with self.assertRaisesRegex(ValueError, "PBI 42.*420.*Unmapped"):
            summary.summarize_pbis(rows)


class BuildOutputTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.temp_dir, ignore_errors=True)

    def test_build_writes_new_copy_and_preserves_source(self):
        source = _source_workbook(self.temp_dir / "PBI_Bug_Task_Report_77_20261001_000000.xlsx",
                                  _sample_rows())
        before = source.read_bytes()
        output = summary.build_summary_workbook(str(source), output_dir=str(self.temp_dir / "out"))
        self.assertNotEqual(Path(output), source)
        self.assertTrue(Path(output).exists())
        self.assertEqual(source.read_bytes(), before)
        wb = load_workbook(output, data_only=False)
        self.assertEqual(wb.sheetnames, ["Summary", "All", "FPSO", "Foundation"])
        self.assertEqual(wb["Summary"].max_row, 4)
        wb.close()


class LatestSummaryStructureTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.report_path = _latest_summary_report_path()
        if cls.report_path is None:
            raise unittest.SkipTest(f"No generated flat Summary workbook found in {REPORT_DIR}.")
        cls.workbook = load_workbook(cls.report_path, data_only=True)

    @classmethod
    def tearDownClass(cls):
        cls.workbook.close()

    def test_latest_summary_is_flat_and_reconciles_to_all_tab(self):
        ws = self.workbook["Summary"]
        self.assertEqual(ws["E1"].value, "Effort")
        self.assertEqual(ws["K1"].value, "Percentage")
        self.assertEqual([ws.cell(2, c).value for c in range(1, 17)], summary.SUMMARY_HEADERS)
        all_rows = summary.read_all_tab(self.workbook)
        expected = summary.summarize_pbis(all_rows)
        self.assertEqual(ws.max_row, 2 + len(expected))
        cols = _summary_map(ws)
        for row_num, pbi in enumerate(expected, start=3):
            self.assertEqual(ws.cell(row_num, cols["PBI ID"]).value, pbi["pbi_id"])
            self.assertAlmostEqual(float(re.match(r"^[\d.]+", str(ws.cell(row_num, 4).value)).group()),
                                   pbi["total_hours"])
            for category in summary.CATEGORY_ORDER:
                self.assertAlmostEqual(ws.cell(row_num, _effort_col(category)).value,
                                       pbi["category_hours"][category])
                self.assertAlmostEqual(ws.cell(row_num, _percentage_col(category)).value,
                                       pbi["category_percentages"][category])

    def test_latest_summary_has_no_charts_or_data_merges(self):
        ws = self.workbook["Summary"]
        self.assertEqual(len(ws._charts), 0)
        self.assertEqual({str(r) for r in ws.merged_cells.ranges}, {"E1:J1", "K1:P1"})

    def test_latest_summary_colors_and_missing_values(self):
        ws = self.workbook["Summary"]
        self.assertTrue(ws["A2"].font.bold)
        self.assertIsNone(ws["A2"].fill.fill_type)
        self.assertIsNone(ws["A3"].fill.fill_type)
        self.assertIsNone(ws["A2"].border.bottom.style)
        for row in range(3, ws.max_row + 1):
            if ws.cell(row, 3).value == summary.NO_STORY_POINT_PLACEHOLDER:
                self.assertEqual(ws.cell(row, 3).fill.fgColor.rgb, "FFFFFF00")
            if isinstance(ws.cell(row, 4).value, str) and summary.MISSING_HOURS_NOTE in ws.cell(row, 4).value:
                self.assertEqual(ws.cell(row, 4).fill.fgColor.rgb, "FFFFFF00")


class LatestSummaryLiveApiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.report_path = _latest_summary_report_path()
        if cls.report_path is None:
            raise unittest.SkipTest("No generated summary workbook is available for live validation.")
        token = report.load_pat_token()
        if not token:
            raise unittest.SkipTest("PAT unavailable in AZURE_DEVOPS_PAT/localconfig; live validation skipped.")
        cls.session = report.get_session(token)
        cls.workbook = load_workbook(cls.report_path, data_only=True)
        cls.all_rows = summary.read_all_tab(cls.workbook)
        cls.pbi_rows = summary.summarize_pbis(cls.all_rows)
        cls.team_by_pbi = {}
        for row in cls.all_rows:
            cls.team_by_pbi[row["pbi_id"]] = row.get("team")
        cls.query_ids_by_team = {}
        for team in TEAM_TABS:
            ws = cls.workbook[team]
            source_url = ws.cell(1, 2).value
            info = report.parse_query_url(source_url)
            cls.query_ids_by_team[team] = set(report.run_wiql_query_by_id(
                cls.session, info["org"], info["project"], info["query_id"]))

        cls.live_by_pbi = {}
        for pbi in cls.pbi_rows:
            pbi_id = pbi["pbi_id"]
            source_row = next(row for row in cls.all_rows if row["pbi_id"] == pbi_id)
            parsed = _parse_work_item_url(source_row["pbi_url"])
            if parsed is None:
                raise AssertionError(f"PBI {pbi_id} has no parseable Azure DevOps hyperlink")
            org, project, _ = parsed
            pbi_items = report.get_work_items_full(cls.session, org, project, [pbi_id])
            if not pbi_items:
                raise AssertionError(f"Live Azure DevOps returned no details for PBI {pbi_id}")
            pbi_item = pbi_items[0]
            child_ids = report.collect_child_ids([pbi_item])
            child_items = report.get_work_items_full(cls.session, org, project, child_ids) if child_ids else []
            children = report.filter_types(child_items, report.CHILD_TYPES)
            cls.live_by_pbi[pbi_id] = {
                "org": org, "project": project, "item": pbi_item, "children": children,
            }

    @classmethod
    def tearDownClass(cls):
        if hasattr(cls, "workbook"):
            cls.workbook.close()

    def _summary_row_map(self):
        ws = self.workbook["Summary"]
        return {ws.cell(row, 1).value: row for row in range(3, ws.max_row + 1)
                if ws.cell(row, 1).value is not None}

    def test_each_summary_pbi_identity_and_story_points_match_live_pbi(self):
        ws = self.workbook["Summary"]
        rows = self._summary_row_map()
        failures = []
        for pbi_id, row_num in rows.items():
            live_item = self.live_by_pbi[pbi_id]["item"]
            live = live_item.get("fields", {})
            if live.get("System.WorkItemType") != "Product Backlog Item":
                failures.append((pbi_id, "Work Item Type", live.get("System.WorkItemType"),
                                 "Product Backlog Item"))
            summary_title = ws.cell(row_num, 2).value
            live_title = live.get("System.Title")
            if summary_title != live_title:
                failures.append((pbi_id, "PBI Title", summary_title, live_title))
            expected = report.format_story_points(live.get(report.STORY_POINTS_FIELD))
            actual = ws.cell(row_num, 3).value
            if actual != expected:
                failures.append((pbi_id, "Story Points", actual, expected))
        self.assertFalse(failures, f"Summary PBI fields differ from live PBIs: {failures}")

    def test_summary_pbis_are_returned_by_their_team_queries(self):
        failures = []
        for pbi_id, team in self.team_by_pbi.items():
            if team not in self.query_ids_by_team or pbi_id not in self.query_ids_by_team[team]:
                failures.append((pbi_id, team, "PBI absent from embedded team query result"))
        self.assertFalse(failures, f"PBI query-scope mismatches: {failures}")

    def test_live_child_inventory_hours_categories_and_missing_hours_match(self):
        ws = self.workbook["Summary"]
        row_map = self._summary_row_map()
        all_task_ids_by_pbi = {}
        source_rows_by_pbi_task = {}
        for row in self.all_rows:
            if row["task_id"] is not None:
                all_task_ids_by_pbi.setdefault(row["pbi_id"], set()).add(row["task_id"])
                source_rows_by_pbi_task[(row["pbi_id"], row["task_id"])] = row

        failures = []
        for pbi in self.pbi_rows:
            pbi_id = pbi["pbi_id"]
            live_children = self.live_by_pbi[pbi_id]["children"]
            live_ids = {child["id"] for child in live_children}
            if live_ids != all_task_ids_by_pbi.get(pbi_id, set()):
                failures.append((pbi_id, "child IDs", sorted(live_ids), sorted(all_task_ids_by_pbi.get(pbi_id, set()))))

            total = 0.0
            missing_ids = []
            by_category = {category: 0.0 for category in summary.CATEGORY_ORDER}
            missing_task_value_mismatches = []
            task_category_mismatches = []
            for child in live_children:
                fields = child.get("fields", {})
                raw_hours = fields.get(report.COMPLETED_WORK_FIELD)
                if raw_hours is None or raw_hours == "":
                    missing_ids.append(child["id"])
                hours = summary.hours_to_number(raw_hours)
                total += hours
                category = report.categorize_task(
                    fields.get("System.WorkItemType", ""),
                    fields.get("System.Title", ""),
                    report.strip_html(fields.get("System.Description", "")),
                )
                source_row = source_rows_by_pbi_task.get((pbi_id, child["id"]))
                if source_row is None:
                    continue
                if source_row["categorization"] != category:
                    task_category_mismatches.append(
                        (child["id"], fields.get("System.Title"),
                         source_row["categorization"], category)
                    )
                source_missing = summary.is_missing_hours(source_row["completed_work"])
                live_missing = raw_hours is None or raw_hours == ""
                if source_missing != live_missing:
                    missing_task_value_mismatches.append(
                        (child["id"], fields.get("System.Title"),
                         source_row["completed_work"], raw_hours)
                    )
                by_category[category] += hours

            row_num = row_map[pbi_id]
            summary_total_raw = ws.cell(row_num, 4).value
            total_text = str(summary_total_raw)
            match = re.match(r"^([\d.]+)", total_text)
            summary_total = float(match.group(1)) if match else None
            if summary_total is None or abs(summary_total - total) > 0.01:
                failures.append((pbi_id, "Total Hours", summary_total_raw, total))
            has_missing_note = summary.MISSING_HOURS_NOTE in total_text
            if bool(missing_ids) != has_missing_note:
                failures.append((pbi_id, "missing-hours annotation", missing_ids, summary_total_raw))
            if missing_task_value_mismatches:
                failures.append((pbi_id, "task-level missing-hours mismatch",
                                 missing_task_value_mismatches))
            if missing_ids:
                if ws.cell(row_num, 4).fill.fgColor.rgb != "FFFFFF00":
                    failures.append((pbi_id, "missing-hours total is not highlighted yellow", missing_ids))
            if task_category_mismatches:
                failures.append((pbi_id, "task-level categorization mismatch",
                                 task_category_mismatches))
            for offset, category in enumerate(summary.CATEGORY_ORDER):
                effort_col = 5 + offset
                percent_col = 11 + offset
                actual_effort = ws.cell(row_num, effort_col).value
                actual_percent = ws.cell(row_num, percent_col).value
                expected_percent = by_category[category] * 100.0 / total if total else 0.0
                if actual_effort is None or abs(float(actual_effort) - by_category[category]) > 0.01:
                    failures.append((pbi_id, category, "effort", actual_effort, by_category[category]))
                if actual_percent is None or abs(float(actual_percent) - expected_percent) > 0.11:
                    failures.append((pbi_id, category, "percentage", actual_percent, expected_percent))
            if abs(sum(by_category.values()) - total) > 0.01:
                failures.append((pbi_id, "live category/total reconciliation", sum(by_category.values()), total))
            if total and abs(sum(by_category.values()) * 100 / total - 100) > 0.01:
                failures.append((pbi_id, "live percentages do not total 100%"))

        self.assertFalse(failures, f"Live PBI/child summary validation failures:\n{failures}")


if __name__ == "__main__":
    unittest.main()
