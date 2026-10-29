"""Unit and live Azure DevOps tests for the flat PBI Summary workbook."""

import os
import re
import shutil
import tempfile
import unittest
from unittest.mock import patch
from pathlib import Path

from openpyxl import Workbook, load_workbook

import pbi_bug_task_report as report
import pbi_bug_task_summary as summary


REPORT_DIR = Path(summary.OUTPUT_DIR)
TEAM_TABS = ("FPSO", "Foundation")

QA_TESTING_CATEGORY = "QA / Testing"

_ANSI_RED = "\033[91m"
_ANSI_RESET = "\033[0m"


def _red(text):
    """Wrap text in ANSI red so QA-owner violations stand out in terminal output."""
    return f"{_ANSI_RED}{text}{_ANSI_RESET}"


_SELECTED_SUMMARY_PATH = "__not_chosen__"


def _prompt_for_summary_report(candidates):
    """Ask which Summary workbook these tests should validate against.

    Always prompts, even when no file was auto-discovered in REPORT_DIR --
    the user can pick one of the listed candidates (newest-first by
    modification time) or type a full path to the actual live file, rather
    than the tests silently trusting mtime or giving up because the file
    lives somewhere other than REPORT_DIR.
    """
    if candidates:
        print(f"\nSummary workbooks found in {REPORT_DIR}:")
        for index, path in enumerate(candidates, 1):
            print(f"  [{index}] {path.name}")
        hint = f"1-{len(candidates)}, Enter for newest, or a full path"
    else:
        print(f"\nNo Summary workbooks auto-discovered in {REPORT_DIR}.")
        hint = "a full path to the Summary workbook .xlsx to validate"

    while True:
        try:
            choice = input(f"Which Summary workbook should these tests validate? [{hint}]: ").strip()
        except EOFError:
            choice = ""
        if not choice:
            return candidates[0] if candidates else None
        if choice.isdigit():
            index = int(choice)
            if candidates and 1 <= index <= len(candidates):
                return candidates[index - 1]
            print(f"Invalid selection '{choice}'.")
            continue
        if len(choice) >= 2 and choice[0] == choice[-1] and choice[0] in "\"'":
            choice = choice[1:-1]
        path = Path(choice)
        if path.is_file():
            return path
        print(f"File not found: {path}")


def _latest_summary_report_path():
    """Return the Summary workbook the user selects to validate against, or
    None. The user is always prompted once per test run; the choice is
    cached so both "Latest..." test classes below use the same file.
    """
    global _SELECTED_SUMMARY_PATH
    if _SELECTED_SUMMARY_PATH != "__not_chosen__":
        return _SELECTED_SUMMARY_PATH

    reports = []
    if REPORT_DIR.exists():
        reports = sorted(
            (path for path in REPORT_DIR.glob("PBI_Bug_Task_Report_*_Summary_*.xlsx")
             if not path.name.startswith("~$")),
            key=lambda path: path.stat().st_mtime,
            reverse=True,
        )
    _SELECTED_SUMMARY_PATH = _prompt_for_summary_report(reports)
    return _SELECTED_SUMMARY_PATH


def _has_pbi_state_summary(path):
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        return any(
            cell.value == "PBI State"
            for row in workbook["Summary"].iter_rows()
            for cell in row
        )
    finally:
        workbook.close()


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
    return 6 + summary.CATEGORY_ORDER.index(category)


def _percentage_col(category):
    return 12 + summary.CATEGORY_ORDER.index(category)


def _parse_work_item_url(url):
    match = re.search(r"dev\.azure\.com/([^/]+)/([^/]+)/_workitems/edit/(\d+)", url or "")
    return match.groups() if match else None


def _format_failures(failures):
    """Render failure dicts one per line, showing which tab/row each came from."""
    if not failures:
        return "(no failures)"
    lines = []
    for failure in failures:
        if isinstance(failure, dict):
            location_bits = []
            if failure.get("tab") is not None:
                location_bits.append(f"tab={failure['tab']}")
            if failure.get("row") is not None:
                location_bits.append(f"row={failure['row']}")
            location = f"[{', '.join(location_bits)}] " if location_bits else ""
            details = ", ".join(
                f"{key}={value!r}" for key, value in failure.items()
                if key not in ("tab", "row")
            )
            lines.append(f"{location}{details}")
        else:
            lines.append(repr(failure))
    return "\n" + "\n".join(lines)


class FilenameAndWorkbookValidationTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.temp_dir, ignore_errors=True)

    def test_filename_prefix_required(self):
        summary.validate_input_filename(self.temp_dir / "Effort_Level_Analysis_1.xlsx")
        with self.assertRaisesRegex(ValueError, "Effort_Level_Analysis_"):
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

    def test_himanshu_email_and_display_name_are_recognized(self):
        configured_email = summary.himanshu_email()
        self.assertTrue(configured_email, "HIMANSHU_EMAIL must be set in localconfig or the environment")
        self.assertTrue(summary.is_himanshu_assignee(configured_email))
        self.assertTrue(summary.is_himanshu_assignee("HimanshuPathak"))
        self.assertFalse(summary.is_himanshu_assignee("Another Person"))

    def test_himanshu_assigned_pbi_and_its_task_rows_are_excluded(self):
        rows = [
            {"pbi_id": 77, "task_id": 770, "pbi_title": "Assigned PBI"},
            {"pbi_id": 78, "task_id": 780, "pbi_title": "Other PBI"},
        ]
        live_metadata = {
            77: {"state": "Active", "assigned_to": {"uniqueName": summary.himanshu_email()}},
            78: {"state": "New", "assigned_to": {"displayName": "Another Person"}},
        }

        included_rows, excluded_ids = summary.exclude_himanshu_assigned_pbis(
            rows, live_metadata
        )

        self.assertEqual(excluded_ids, {77})
        self.assertEqual({row["pbi_id"] for row in included_rows}, {78})

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
            rows = summary.read_all_tab(wb)
            for row in rows:
                row["pbi_state"] = "Active" if row["pbi_id"] == 88001 else "New"
            return rows
        finally:
            wb.close()

    def test_placeholder_missing_zero_and_zero_total_semantics(self):
        self.assertTrue(summary.is_missing_hours(None))
        self.assertTrue(summary.is_missing_hours("No_Hours"))
        self.assertFalse(summary.is_missing_hours(0))
        self.assertEqual(summary.hours_to_number("No_Hours"), 0)
        self.assertEqual(summary.hours_to_number(0), 0)
        rows = [{"pbi_id": 9, "pbi_title": "Empty", "pbi_state": "New", "pbi_story_points": 0,
                 "pbi_url": None, "task_id": None, "task_title": None,
                 "completed_work": None, "categorization": None}]
        [empty] = summary.summarize_pbis(rows)
        self.assertEqual(empty["total_hours"], 0)
        self.assertEqual(set(empty["category_hours"].values()), {0.0})
        self.assertEqual(set(empty["category_percentages"].values()), {0.0})

    def test_unexpected_or_blank_category_fails_with_pbi_and_task(self):
        rows = [{"pbi_id": 900, "pbi_title": "P", "pbi_state": "New", "pbi_story_points": 1,
                 "pbi_url": None, "task_id": 901, "task_title": "Unclassified",
                 "completed_work": 2, "categorization": ""}]
        with self.assertRaisesRegex(ValueError, "PBI 900.*901.*unexpected"):
            summary.summarize_pbis(rows)

    def test_categories_and_total_are_from_same_task_hours(self):
        rows = [
            {"pbi_id": 5, "pbi_title": "P", "pbi_state": "Active", "pbi_story_points": 2, "pbi_url": None,
             "task_id": 50, "task_title": "A", "completed_work": 2, "categorization": "Review"},
            {"pbi_id": 5, "pbi_title": "P", "pbi_state": "Active", "pbi_story_points": 2, "pbi_url": None,
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
        self.assertEqual((ws["F1"].value, ws["L1"].value), ("Effort", "Percentage"))
        self.assertEqual([ws.cell(2, col).value for col in range(1, 18)], summary.SUMMARY_HEADERS)
        self.assertEqual(ws.max_row, 2 + len(pbi))
        self.assertEqual([str(r) for r in ws.merged_cells.ranges], ["F1:K1", "L1:Q1"])
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
            rows = summary.read_all_tab(wb)
            for row in rows:
                row["pbi_state"] = "Active" if row["pbi_id"] == 88001 else "New"
            return rows
        finally:
            wb.close()

    def test_summary_row_values_hyperlink_and_missing_hours(self):
        wb, _ = self._create(self._load_source())
        ws = wb["Summary"]
        cols = _summary_map(ws)
        row = 3
        self.assertEqual(ws.cell(row, cols["PBI ID"]).value, 88001)
        self.assertEqual(ws.cell(row, cols["PBI Title"]).value, "Dynamic sample PBI")
        self.assertEqual(ws.cell(row, cols["PBI State"]).value, "Active")
        self.assertEqual(ws.cell(row, cols["PBI Story Points"]).value, 3)
        self.assertEqual(ws.cell(row, cols["Total Hours"]).value, "8 (some tasks/bugs missing hours)")
        self.assertEqual(ws.cell(row, _effort_col("Development")).value, 6)
        self.assertEqual(ws.cell(row, _effort_col("Bug Fixing")).value, 2)
        self.assertEqual(ws.cell(row, _effort_col("QA / Testing")).value, 0)
        self.assertAlmostEqual(ws.cell(row, _percentage_col("Development")).value, 75)
        self.assertEqual(ws.cell(row, 1).hyperlink.target,
                         "https://dev.azure.com/org/proj/_workitems/edit/88001")
        self.assertEqual(ws.cell(row, 5).fill.fgColor.rgb, summary.YELLOW_FILL.fgColor.rgb)
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
            for column in range(1, 6):
                cell = ws.cell(row, column)
                self.assertEqual(cell.fill.fgColor.rgb, summary.DARK_BLUE_HEADER_ARGB)
                self.assertEqual(cell.font.color.rgb, summary.HEADER_FONT_COLOR)
                self.assertTrue(cell.font.bold)
            for column in range(6, 12):
                cell = ws.cell(row, column)
                self.assertEqual(cell.fill.fgColor.rgb, summary.EFFORT_HEADER_ARGB)
                self.assertEqual(cell.font.color.rgb, summary.HEADER_FONT_COLOR)
                self.assertTrue(cell.font.bold)
            for column in range(12, 18):
                cell = ws.cell(row, column)
                self.assertEqual(cell.fill.fgColor.rgb, summary.DARK_BLUE_HEADER_ARGB)
                self.assertEqual(cell.font.color.rgb, summary.HEADER_FONT_COLOR)
                self.assertTrue(cell.font.bold)
        self.assertIsNone(ws["A3"].fill.fill_type)
        for address in ("A2", "Q2", "A3", "Q3"):
            cell = ws[address]
            for edge in (cell.border.left, cell.border.right, cell.border.top, cell.border.bottom):
                self.assertEqual(edge.style, "thin", f"{address} is missing a visible border")
            self.assertEqual(cell.border.top.color.rgb, summary.VISIBLE_BORDER_COLOR)
        for address, expected_left, expected_right in (
            ("F1", "thin", None), ("K1", None, "thin"),
            ("L1", "thin", None), ("Q1", None, "thin"),
        ):
            cell = ws[address]
            self.assertEqual(cell.border.left.style, expected_left)
            self.assertEqual(cell.border.right.style, expected_right)
            self.assertEqual(cell.border.top.style, "thin")
            self.assertEqual(cell.border.bottom.style, "thin")
        self.assertEqual(ws["A3"].font.color.rgb, summary.HYPERLINK_FONT_COLOR)
        self.assertEqual(ws["D4"].fill.fgColor.rgb, "FFFFFF00")
        wb.close()

    def test_unknown_category_reports_pbi_and_task(self):
        rows = [{"pbi_id": 42, "pbi_title": "P", "pbi_state": "Active", "pbi_story_points": 1,
                 "pbi_url": None, "task_id": 420, "task_title": "bad cat",
                 "completed_work": 3, "categorization": "Unmapped"}]
        with self.assertRaisesRegex(ValueError, "PBI 42.*420.*Unmapped"):
            summary.summarize_pbis(rows)


class BuildOutputTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.temp_dir, ignore_errors=True)

    def test_build_writes_new_copy_and_preserves_source(self):
        source = _source_workbook(self.temp_dir / "Effort_Level_Analysis_77_20261001_000000.xlsx",
                                  _sample_rows())
        before = source.read_bytes()
        with patch.object(
            summary,
            "fetch_live_pbi_metadata",
            return_value={
                88001: {"state": "Active", "assigned_to": "Engineer"},
                88002: {"state": "New", "assigned_to": "Engineer"},
            },
        ):
            output = summary.build_summary_workbook(
                str(source), output_dir=str(self.temp_dir / "out")
            )
        self.assertNotEqual(Path(output), source)
        self.assertTrue(Path(output).exists())
        self.assertEqual(source.read_bytes(), before)
        wb = load_workbook(output, data_only=False)
        self.assertEqual(wb.sheetnames, ["Summary", "All", "FPSO", "Foundation"])
        self.assertEqual(wb["Summary"].max_row, 4)
        self.assertEqual(wb["Summary"]["C3"].value, "Active")
        wb.close()


class LatestSummaryStructureTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.report_path = _latest_summary_report_path()
        if cls.report_path is None:
            raise unittest.SkipTest(f"No generated flat Summary workbook found in {REPORT_DIR}.")
        if not _has_pbi_state_summary(cls.report_path):
            raise unittest.SkipTest(
            "Latest generated Summary predates the PBI State column; regenerate it "
            "before running current generated-workbook checks."
            )
        cls.workbook = load_workbook(cls.report_path, data_only=True)

    @classmethod
    def tearDownClass(cls):
        cls.workbook.close()

    def test_latest_summary_is_flat_and_reconciles_to_all_tab(self):
        ws = self.workbook["Summary"]
        self.assertEqual(ws["F1"].value, "Effort")
        self.assertEqual(ws["L1"].value, "Percentage")
        self.assertEqual([ws.cell(2, c).value for c in range(1, 18)], summary.SUMMARY_HEADERS)
        all_rows = summary.read_all_tab(self.workbook)
        expected = summary.summarize_pbis(all_rows)

        actual_ids = [ws.cell(r, 1).value for r in range(3, ws.max_row + 1)]
        expected_ids = [pbi["pbi_id"] for pbi in expected]
        diagnostics = (
            f"report_path={self.report_path}; "
            f"Summary data rows={ws.max_row - 2} (ws.max_row={ws.max_row}); "
            f"summarize_pbis(all_rows) PBI count={len(expected)}; "
            f"All-tab rows={len(all_rows)}, unique PBI IDs in All tab="
            f"{len({row['pbi_id'] for row in all_rows})}; "
            f"Summary rows with a blank PBI ID (row numbers)="
            f"{[r for r, pid in zip(range(3, ws.max_row + 1), actual_ids) if pid is None]}; "
            f"duplicate PBI IDs within Summary tab="
            f"{sorted({pid for pid in actual_ids if pid is not None and actual_ids.count(pid) > 1}, key=str)}; "
            f"PBI IDs present in Summary but not expected from All tab="
            f"{sorted(set(filter(None, actual_ids)) - set(expected_ids), key=str)}; "
            f"PBI IDs expected from All tab but missing from Summary="
            f"{sorted(set(expected_ids) - set(actual_ids), key=str)}"
        )
        self.assertEqual(ws.max_row, 2 + len(expected), diagnostics)
        cols = _summary_map(ws)
        for row_num, pbi in enumerate(expected, start=3):
            self.assertEqual(ws.cell(row_num, cols["PBI ID"]).value, pbi["pbi_id"],
                              f"PBI ID mismatch at Summary row {row_num} (report_path={self.report_path})")
            self.assertAlmostEqual(
                float(re.match(r"^[\d.]+", str(ws.cell(row_num, 5).value)).group()),
                pbi["total_hours"],
                msg=f"Total Hours mismatch for PBI {pbi['pbi_id']} at Summary row {row_num}")
            for category in summary.CATEGORY_ORDER:
                self.assertAlmostEqual(
                    ws.cell(row_num, _effort_col(category)).value,
                    pbi["category_hours"][category],
                    msg=f"{category!r} hours mismatch for PBI {pbi['pbi_id']} at Summary row {row_num}")
                self.assertAlmostEqual(
                    ws.cell(row_num, _percentage_col(category)).value,
                    pbi["category_percentages"][category],
                    msg=f"{category!r} percentage mismatch for PBI {pbi['pbi_id']} at Summary row {row_num}")

    def test_latest_summary_has_no_charts_or_data_merges(self):
        ws = self.workbook["Summary"]
        self.assertEqual(len(ws._charts), 0)
        self.assertEqual({str(r) for r in ws.merged_cells.ranges}, {"F1:K1", "L1:Q1"})

    def test_latest_summary_colors_and_missing_values(self):
        ws = self.workbook["Summary"]
        self.assertTrue(ws["A2"].font.bold, "[tab=Summary, row=2] header row must be bold")
        self.assertIsNone(ws["A2"].fill.fill_type, "[tab=Summary, row=2] header row must not be filled")
        self.assertIsNone(ws["A3"].fill.fill_type, "[tab=Summary, row=3] first data row must not be filled")
        self.assertIsNone(ws["A2"].border.bottom.style,
                          "[tab=Summary, row=2] header row must not have a bottom border")
        for row in range(3, ws.max_row + 1):
            if ws.cell(row, 4).value == summary.NO_STORY_POINT_PLACEHOLDER:
                self.assertEqual(ws.cell(row, 4).fill.fgColor.rgb, "FFFFFF00",
                                  f"[tab=Summary, row={row}] No_Story_Point placeholder must be highlighted yellow")
            if isinstance(ws.cell(row, 5).value, str) and summary.MISSING_HOURS_NOTE in ws.cell(row, 5).value:
                self.assertEqual(ws.cell(row, 5).fill.fgColor.rgb, "FFFFFF00",
                                  f"[tab=Summary, row={row}] missing-hours note must be highlighted yellow")

    def test_qa_testing_tasks_are_owned_by_expected_qa_engineers(self):
        """Every Task/Bug categorized "QA / Testing" in the All tab must be
        assigned to one of the dedicated QA engineers (Hanumanth Sunkara or
        Salina Hamal). Anything else is a red flag for mis-categorization --
        e.g. Task 39082 ("Identify Areas of Investigation in Code"), which
        was wrongly bucketed into "QA / Testing" because its description
        happened to mention "test"/"testing", despite being owned by Jon
        Nance and actually being requirement-study work.
        """
        expected_owners = summary.qa_owners()
        self.assertTrue(expected_owners, "QA_OWNERS must be set in localconfig or the environment")

        ws = self.workbook["All"]
        cols = _header_map(ws)
        violations = []
        for row_idx in range(3, ws.max_row + 1):
            if ws.cell(row_idx, cols["Task Categorization"]).value != QA_TESTING_CATEGORY:
                continue
            owner = ws.cell(row_idx, cols["Assigned To"]).value
            if owner in expected_owners:
                continue
            violations.append({
                "pbi_id": ws.cell(row_idx, cols["PBI ID"]).value,
                "task_id": ws.cell(row_idx, cols["Task ID"]).value,
                "task_title": ws.cell(row_idx, cols["Task Title"]).value,
                "owner": owner,
            })

        if violations:
            lines = "\n".join(
                f"  - PBI {v['pbi_id']}, Task/Bug {v['task_id']} ({v['task_title']}): "
                f"owner={v['owner']!r}"
                for v in violations
            )
            print(_red(
                f"\nFound {len(violations)} task(s) categorized {QA_TESTING_CATEGORY!r} "
                f"whose owner is not one of {sorted(expected_owners)}:\n{lines}"
            ))
        self.assertFalse(violations, _red(
            f"{len(violations)} {QA_TESTING_CATEGORY!r} task(s) owned by someone other than "
            f"{sorted(expected_owners)}; see printed list above (report_path={self.report_path})."
        ))


class LatestSummaryLiveApiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.report_path = _latest_summary_report_path()
        if cls.report_path is None:
            raise unittest.SkipTest("No generated summary workbook is available for live validation.")
        if not _has_pbi_state_summary(cls.report_path):
            raise unittest.SkipTest(
            "Latest generated Summary predates the PBI State column; regenerate it "
            "before live validation."
            )
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
        cls.org_project_by_team = {}
        for team in TEAM_TABS:
            ws = cls.workbook[team]
            source_url = ws.cell(1, 2).value
            info = report.parse_query_url(source_url)
            cls.org_project_by_team[team] = (info["org"], info["project"])
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
                failures.append({"tab": "Summary", "row": row_num, "pbi_id": pbi_id,
                                  "field": "Work Item Type", "actual": live.get("System.WorkItemType"),
                                  "expected": "Product Backlog Item"})
            summary_title = ws.cell(row_num, 2).value
            live_title = live.get("System.Title")
            if summary_title != live_title:
                failures.append({"tab": "Summary", "row": row_num, "pbi_id": pbi_id,
                                  "field": "PBI Title", "actual": summary_title, "expected": live_title})
            expected = report.format_story_points(live.get(report.STORY_POINTS_FIELD))
            actual = ws.cell(row_num, 4).value
            if actual != expected:
                failures.append({"tab": "Summary", "row": row_num, "pbi_id": pbi_id,
                                  "field": "Story Points", "actual": actual, "expected": expected})
        self.assertFalse(failures, f"Summary PBI fields differ from live PBIs: {_format_failures(failures)}")

    def test_every_summary_pbi_has_state_matching_live_azure_devops(self):
        ws = self.workbook["Summary"]
        failures = []
        for pbi_id, row_num in self._summary_row_map().items():
            actual_state = ws.cell(row_num, 3).value
            live_state = self.live_by_pbi[pbi_id]["item"].get("fields", {}).get("System.State")
            if not actual_state:
                failures.append({"tab": "Summary", "row": row_num, "pbi_id": pbi_id,
                                  "issue": "PBI State is blank", "actual": actual_state, "live": live_state})
            elif actual_state != live_state:
                failures.append({"tab": "Summary", "row": row_num, "pbi_id": pbi_id,
                                  "issue": "PBI State mismatch", "actual": actual_state, "live": live_state})
        self.assertFalse(failures, f"Summary PBI State differs from live Azure DevOps: {_format_failures(failures)}")

    def test_live_query_pbIs_assigned_to_himanshu_are_excluded_from_summary(self):
        """Any PBI assigned to Himanshu in the live team query results must
        not appear in the generated Summary, matching email or display name.
        """
        summary_row_map = self._summary_row_map()
        summary_ids = set(summary_row_map)
        assigned_pbi_ids = set()
        failures = []
        for team in TEAM_TABS:
            org, project = self.org_project_by_team[team]
            query_items = report.get_work_items_full(
                self.session, org, project, sorted(self.query_ids_by_team[team])
            )
            for pbi in report.filter_types(query_items, report.PBI_TYPES):
                assignee = pbi.get("fields", {}).get("System.AssignedTo")
                if not summary.is_himanshu_assignee(assignee):
                    continue
                pbi_id = pbi["id"]
                assigned_pbi_ids.add(pbi_id)
                if pbi_id in summary_ids:
                    failures.append({"tab": "Summary", "row": summary_row_map[pbi_id], "team": team,
                                      "pbi_id": pbi_id, "issue": "Himanshu-assigned PBI present in Summary"})

        if not assigned_pbi_ids:
            self.skipTest(
                "No Himanshu-assigned PBIs were returned by embedded team queries; "
                "live inclusion could not be exercised."
            )
        self.assertFalse(
            failures,
            f"Live PBIs assigned to Himanshu were not excluded from Summary: "
            f"{_format_failures(failures)}; live matching PBI IDs={sorted(assigned_pbi_ids)}",
        )

    def test_summary_pbis_are_returned_by_their_team_queries(self):
        failures = []
        summary_row_map = self._summary_row_map()
        for pbi_id, team in self.team_by_pbi.items():
            if team not in self.query_ids_by_team or pbi_id not in self.query_ids_by_team[team]:
                failures.append({"tab": "Summary", "row": summary_row_map.get(pbi_id), "team": team,
                                  "pbi_id": pbi_id, "issue": "PBI absent from embedded team query result"})
        self.assertFalse(failures, f"PBI query-scope mismatches: {_format_failures(failures)}")

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
            row_num = row_map[pbi_id]
            live_children = self.live_by_pbi[pbi_id]["children"]
            live_ids = {child["id"] for child in live_children}
            if live_ids != all_task_ids_by_pbi.get(pbi_id, set()):
                failures.append({"tab": "Summary", "row": row_num, "pbi_id": pbi_id,
                                  "field": "child IDs", "live": sorted(live_ids),
                                  "source": sorted(all_task_ids_by_pbi.get(pbi_id, set()))})

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

            summary_total_raw = ws.cell(row_num, 5).value
            total_text = str(summary_total_raw)
            match = re.match(r"^([\d.]+)", total_text)
            summary_total = float(match.group(1)) if match else None
            if summary_total is None or abs(summary_total - total) > 0.01:
                failures.append({"tab": "Summary", "row": row_num, "pbi_id": pbi_id,
                                  "field": "Total Hours", "actual": summary_total_raw, "expected": total})
            has_missing_note = summary.MISSING_HOURS_NOTE in total_text
            if bool(missing_ids) != has_missing_note:
                failures.append({"tab": "Summary", "row": row_num, "pbi_id": pbi_id,
                                  "field": "missing-hours annotation", "missing_task_ids": missing_ids,
                                  "total_cell": summary_total_raw})
            if missing_task_value_mismatches:
                failures.append({"tab": "Summary", "row": row_num, "pbi_id": pbi_id,
                                  "field": "task-level missing-hours mismatch",
                                  "mismatches": missing_task_value_mismatches})
            if missing_ids:
                if ws.cell(row_num, 5).fill.fgColor.rgb != "FFFFFF00":
                    failures.append({"tab": "Summary", "row": row_num, "pbi_id": pbi_id,
                                      "field": "missing-hours total is not highlighted yellow",
                                      "missing_task_ids": missing_ids})
            if task_category_mismatches:
                failures.append({"tab": "Summary", "row": row_num, "pbi_id": pbi_id,
                                  "field": "task-level categorization mismatch",
                                  "mismatches": task_category_mismatches})
            for offset, category in enumerate(summary.CATEGORY_ORDER):
                effort_col = 6 + offset
                percent_col = 12 + offset
                actual_effort = ws.cell(row_num, effort_col).value
                actual_percent = ws.cell(row_num, percent_col).value
                expected_percent = by_category[category] * 100.0 / total if total else 0.0
                if actual_effort is None or abs(float(actual_effort) - by_category[category]) > 0.01:
                    failures.append({"tab": "Summary", "row": row_num, "pbi_id": pbi_id,
                                      "field": f"{category} effort", "actual": actual_effort,
                                      "expected": by_category[category]})
                if actual_percent is None or abs(float(actual_percent) - expected_percent) > 0.11:
                    failures.append({"tab": "Summary", "row": row_num, "pbi_id": pbi_id,
                                      "field": f"{category} percentage", "actual": actual_percent,
                                      "expected": expected_percent})
            if abs(sum(by_category.values()) - total) > 0.01:
                failures.append({"tab": "Summary", "row": row_num, "pbi_id": pbi_id,
                                  "field": "live category/total reconciliation",
                                  "category_sum": sum(by_category.values()), "total": total})
            if total and abs(sum(by_category.values()) * 100 / total - 100) > 0.01:
                failures.append({"tab": "Summary", "row": row_num, "pbi_id": pbi_id,
                                  "field": "live percentages do not total 100%"})

        self.assertFalse(failures, f"Live PBI/child summary validation failures:{_format_failures(failures)}")


if __name__ == "__main__":
    unittest.main()
