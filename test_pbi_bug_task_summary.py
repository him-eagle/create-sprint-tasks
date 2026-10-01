"""Tests for pbi_bug_task_summary.py.

Covers:
  * Filename / workbook-structure validation (prefix, required tabs, Summary
    must be empty) and the "Himanshu Pathak" pre-flight abort check.
  * Hours/placeholder math (`No_Hours` / blank treated as 0, explicit 0
    preserved) and the fixed six-category breakdown (every category always
    present, even at 0 hours/0%).
  * `write_summary_sheet` structure: vertically merged PBI header cells,
    yellow highlighting for `No_Story_Point` / missing-hours cells, and one
    pie chart per PBI with a slice count matching its nonzero categories
    (no chart at all for a PBI with no Task/Bug children).
  * Output filename generation (`_Summary_<timestamp>` inserted before the
    extension, replacing any existing trailing timestamp).
  * Validation against the LATEST actually-generated summary workbook in
    Reports\\Pbi_bug_tasks\\ (`PBI_Bug_Task_Report_*_Summary_*.xlsx`):
      - every zero-hours category sub-row is cross-checked against the
        underlying "All" tab and the matching team tab to confirm the PBI
        genuinely has no task in that category;
      - no row on any tab is Assigned To "Himanshu Pathak".
    These two checks only need the generated workbook itself -- no Azure
    DevOps access -- and are skipped with an explanation if no summary
    workbook exists yet.
  * Live Azure DevOps API checks (skipped with an explanation if no PAT is
    configured or no summary workbook exists):
      - every PBI ID in the Summary tab is returned by its team's query,
        read straight from the embedded FPSO/Foundation tab source URLs;
      - each PBI's Summary total-hours figure matches a fresh recomputation
        from the live API;
      - every PBI shown in the Summary with zero children is independently
        re-queried and confirmed to genuinely have no child work items;
      - each PBI's pie chart legend/category data (the hidden helper
        columns feeding its slices and bottom legend) matches the set of
        nonzero categories from a fresh, independent recomputation against
        live Azure DevOps data (not the already-generated workbook);
      - each PBI's "Category %" value matches
        hours(category) / total_hours * 100, both computed from a fresh,
        independent recomputation against live Azure DevOps data.

There are no hardcoded/static Azure DevOps IDs or query URLs anywhere in
this file.
"""

import os
import re
import shutil
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

import pbi_bug_task_report as report
import pbi_bug_task_summary as summary


REPORT_DIR = Path(summary.OUTPUT_DIR)
TEAM_TABS = {"FPSO", "Foundation"}


def _latest_summary_report_path():
    """Return the most recently modified generated summary workbook, or
    None if none exists yet."""
    if not REPORT_DIR.exists():
        return None
    reports = sorted(
        (
            p for p in REPORT_DIR.glob("PBI_Bug_Task_Report_*_Summary_*.xlsx")
            if not p.name.startswith("~$")
        ),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    return reports[0] if reports else None


def _normalize_assignee(value):
    return "".join(str(value or "").split()).casefold()


def _parse_work_item_url(url):
    """Parse org/project/id out of a .../_workitems/edit/{id} URL."""
    match = re.search(r"dev\.azure\.com/([^/]+)/([^/]+)/_workitems/edit/(\d+)", url or "")
    if not match:
        return None
    org, project, item_id = match.groups()
    return org, project, int(item_id)


def _parse_total_hours_display(value):
    """Extract the leading numeric total out of a Summary "Total Hours"
    cell, which may be a plain number or a string like
    "35.5 (some tasks/bugs missing hours)"."""
    if isinstance(value, (int, float)):
        return float(value)
    match = re.match(r"^([\d.]+)", str(value or ""))
    return float(match.group(1)) if match else None


def _make_source_workbook(tmp_path, *, summary_rows=None, all_rows=None, team_rows=None, extra_tabs=True):
    """Build a minimal workbook shaped like pbi_bug_task_report.py's output,
    for pure unit tests that don't need a real report file.

    `all_rows` / `team_rows` are lists of dicts with the same keys as the
    real "All" tab columns; `summary_rows` pre-populates the Summary tab
    (used to test the "already processed" abort path).
    """
    wb = Workbook()
    wb.remove(wb.active)

    summary_ws = wb.create_sheet("Summary")
    if summary_rows:
        for r, row in enumerate(summary_rows, start=1):
            for c, value in enumerate(row, start=1):
                summary_ws.cell(row=r, column=c, value=value)

    headers = [
        "Team", "PBI ID", "PBI Title", "PBI Story Points", "PBI Done Date",
        "Task ID", "Task Type", "Task Title", "Assigned To", "State",
        "Original Estimate (hrs)", "Completed Work (hrs)",
        "Task Categorization", "Done Date", "Description",
    ]

    def _write_table(ws, rows):
        ws.cell(row=1, column=1, value="Source URL:")
        ws.cell(row=1, column=2, value="https://dev.azure.com/org/proj/_queries/query/abc/")
        for c, header in enumerate(headers, start=1):
            ws.cell(row=2, column=c, value=header)
        for r, row in enumerate(rows or [], start=3):
            for c, header in enumerate(headers, start=1):
                value = row.get(header)
                cell = ws.cell(row=r, column=c, value=value)
                if header == "PBI ID" and row.get("pbi_url"):
                    cell.hyperlink = row["pbi_url"]

    _write_table(wb.create_sheet("All"), all_rows or [])
    if extra_tabs:
        fpso_rows = [r for r in (all_rows or []) if r.get("Team") == "FPSO"]
        foundation_rows = [r for r in (all_rows or []) if r.get("Team") == "Foundation"]
        _write_table(wb.create_sheet("FPSO"), fpso_rows)
        _write_table(wb.create_sheet("Foundation"), foundation_rows)

    path = tmp_path / "PBI_Bug_Task_Report_99_20260101_000000.xlsx"
    wb.save(path)
    wb.close()
    return path


# ---------------------------------------------------------------------------
# Pure unit tests -- no network, no real report file required.
# ---------------------------------------------------------------------------

class FilenameValidationTests(unittest.TestCase):
    def test_valid_prefix_accepted(self):
        """A file name starting with PBI_Bug_Task_Report_ passes silently."""
        summary.validate_input_filename(r"C:\x\PBI_Bug_Task_Report_103_20260101_000000.xlsx")

    def test_invalid_prefix_rejected(self):
        """Any other file name is rejected with a clear error."""
        with self.assertRaisesRegex(ValueError, "PBI_Bug_Task_Report_"):
            summary.validate_input_filename(r"C:\x\Sprint_Report_20260101_000000.xlsx")


class HoursPlaceholderTests(unittest.TestCase):
    def test_missing_hours_and_placeholder_treated_as_zero(self):
        """None, blank, and the literal No_Hours placeholder all sum as 0."""
        self.assertEqual(summary.hours_to_number(None), 0.0)
        self.assertEqual(summary.hours_to_number(""), 0.0)
        self.assertEqual(summary.hours_to_number("No_Hours"), 0.0)

    def test_explicit_zero_hours_is_real_data_not_missing(self):
        """An explicit 0 counts as 0 hours but is NOT flagged as missing."""
        self.assertEqual(summary.hours_to_number(0), 0.0)
        self.assertFalse(summary.is_missing_hours(0))

    def test_is_missing_hours_detects_blank_and_placeholder(self):
        self.assertTrue(summary.is_missing_hours(None))
        self.assertTrue(summary.is_missing_hours(""))
        self.assertTrue(summary.is_missing_hours("No_Hours"))
        self.assertFalse(summary.is_missing_hours(4))


class SummarizePbisTests(unittest.TestCase):
    def _rows(self, pbi_id, tasks, title="Sample PBI", story_points=5, url="https://x/_workitems/edit/1"):
        rows = [{
            "pbi_id": pbi_id, "pbi_title": title, "pbi_story_points": story_points,
            "pbi_url": url, "task_id": None, "completed_work": None, "categorization": None,
        }]
        for i, (categorization, hours) in enumerate(tasks, start=1):
            rows.append({
                "pbi_id": pbi_id, "pbi_title": title, "pbi_story_points": story_points,
                "pbi_url": url, "task_id": 1000 + i, "completed_work": hours,
                "categorization": categorization,
            })
        return rows

    def test_total_hours_sums_completed_work_treating_missing_as_zero(self):
        """Total hours for a PBI = sum of Completed Work across its tasks,
        with No_Hours/blank treated as 0.

        Example: tasks with Completed Work 4, "No_Hours", and 6 sum to 10,
        not an error and not skipped.
        """
        rows = self._rows(1, [("Development", 4), ("Development", "No_Hours"), ("Development", 6)])
        [pbi] = summary.summarize_pbis(rows)
        self.assertEqual(pbi["total_hours"], 10.0)
        self.assertTrue(pbi["missing_hours"])

    def test_missing_hours_flag_false_when_all_tasks_have_hours(self):
        rows = self._rows(1, [("Development", 4), ("Review", 2)])
        [pbi] = summary.summarize_pbis(rows)
        self.assertFalse(pbi["missing_hours"])

    def test_all_six_fixed_categories_always_present(self):
        """Even a PBI whose only task is "Development" still lists all six
        fixed categories, with 0 hours/0% for the rest.
        """
        rows = self._rows(1, [("Development", 8)])
        [pbi] = summary.summarize_pbis(rows)
        cats = {c["category"]: c["hours"] for c in pbi["categories"]}
        self.assertEqual(set(summary.CATEGORY_ORDER), set(cats) & set(summary.CATEGORY_ORDER))
        for cat in summary.CATEGORY_ORDER:
            if cat != "Development":
                self.assertEqual(cats[cat], 0.0)
        self.assertEqual(cats["Development"], 8.0)

    def test_category_percentages_sum_to_100(self):
        rows = self._rows(1, [("Development", 6), ("Review", 3), ("Bug Fixing", 1)])
        [pbi] = summary.summarize_pbis(rows)
        total_pct = sum(c["percent"] for c in pbi["categories"])
        self.assertAlmostEqual(total_pct, 100.0, places=6)

    def test_pbi_with_no_children_has_zero_total_and_has_children_false(self):
        rows = self._rows(1, [])
        [pbi] = summary.summarize_pbis(rows)
        self.assertEqual(pbi["total_hours"], 0.0)
        self.assertFalse(pbi["has_children"])
        self.assertFalse(pbi["missing_hours"])

    def test_unexpected_category_is_not_dropped(self):
        """A categorization value outside the fixed list (e.g. stale data)
        is still counted, appended after the six fixed categories, instead
        of silently losing those hours."""
        rows = self._rows(1, [("Some Future Category", 5)])
        [pbi] = summary.summarize_pbis(rows)
        cats = {c["category"]: c["hours"] for c in pbi["categories"]}
        self.assertEqual(cats["Some Future Category"], 5.0)
        self.assertEqual(pbi["total_hours"], 5.0)


class OutputFilenameTests(unittest.TestCase):
    def test_strips_existing_timestamp_and_inserts_summary(self):
        result = summary.build_output_filename(
            "PBI_Bug_Task_Report_103_20261001_104744.xlsx", "20261001_135000"
        )
        self.assertEqual(result, "PBI_Bug_Task_Report_103_Summary_20261001_135000.xlsx")

    def test_handles_name_without_trailing_timestamp(self):
        result = summary.build_output_filename("PBI_Bug_Task_Report_103.xlsx", "20261001_135000")
        self.assertEqual(result, "PBI_Bug_Task_Report_103_Summary_20261001_135000.xlsx")


class WorkbookValidationTests(unittest.TestCase):
    def setUp(self):
        self.tmp_dir = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.tmp_dir, ignore_errors=True)

    def test_missing_required_tab_is_rejected(self):
        wb = Workbook()
        wb.remove(wb.active)
        wb.create_sheet("Summary")
        wb.create_sheet("All")
        with self.assertRaisesRegex(ValueError, "Foundation"):
            summary.validate_workbook_structure(wb)
        wb.close()

    def test_non_empty_summary_tab_is_rejected(self):
        path = _make_source_workbook(
            self.tmp_dir, summary_rows=[["already", "has", "content"]], all_rows=[]
        )
        wb = load_workbook(path)
        with self.assertRaisesRegex(ValueError, "already contains content"):
            summary.validate_workbook_structure(wb)
        wb.close()

    def test_valid_structure_passes(self):
        path = _make_source_workbook(self.tmp_dir, all_rows=[])
        wb = load_workbook(path)
        summary.validate_workbook_structure(wb)  # should not raise
        wb.close()


class BlockedAssigneeTests(unittest.TestCase):
    def setUp(self):
        self.tmp_dir = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.tmp_dir, ignore_errors=True)

    def test_detects_himanshu_pathak_case_and_spacing_variants(self):
        """The row appears on both the "All" tab and its team tab (FPSO),
        exactly as the real report structure does, so two violations are
        expected -- one per tab -- both pointing at the same PBI/Task."""
        rows = [
            {"Team": "FPSO", "PBI ID": 1, "PBI Title": "P", "Task ID": 10,
             "Task Title": "T", "Assigned To": "himanshupathak"},
        ]
        path = _make_source_workbook(self.tmp_dir, all_rows=rows)
        wb = load_workbook(path)
        violations = summary.find_blocked_assignee_rows(wb)
        wb.close()
        self.assertEqual(len(violations), 2)
        self.assertEqual({v["sheet"] for v in violations}, {"All", "FPSO"})
        for v in violations:
            self.assertEqual(v["pbi_id"], 1)
            self.assertEqual(v["task_id"], 10)

    def test_no_violation_when_assignee_absent(self):
        rows = [
            {"Team": "FPSO", "PBI ID": 1, "PBI Title": "P", "Task ID": 10,
             "Task Title": "T", "Assigned To": "Someone Else"},
        ]
        path = _make_source_workbook(self.tmp_dir, all_rows=rows)
        wb = load_workbook(path)
        violations = summary.find_blocked_assignee_rows(wb)
        wb.close()
        self.assertEqual(violations, [])

    def test_build_summary_workbook_aborts_when_blocked_assignee_present(self):
        rows = [
            {"Team": "FPSO", "PBI ID": 1, "PBI Title": "P", "Task ID": 10,
             "Task Title": "T", "Assigned To": "Himanshu Pathak",
             "Completed Work (hrs)": 4, "Task Categorization": "Development"},
        ]
        path = _make_source_workbook(self.tmp_dir, all_rows=rows)
        out_dir = self.tmp_dir / "out"
        with self.assertRaisesRegex(ValueError, "Himanshu Pathak"):
            summary.build_summary_workbook(str(path), output_dir=str(out_dir))


class WriteSummarySheetTests(unittest.TestCase):
    """Exercises write_summary_sheet directly against hand-built summaries,
    without going through a full source workbook."""

    def _build(self, pbi_summaries):
        wb = Workbook()
        wb.remove(wb.active)
        wb.create_sheet("Summary")
        summary.write_summary_sheet(wb, pbi_summaries)
        return wb

    def test_merged_cell_ranges_cover_all_category_rows(self):
        [pbi] = summary.summarize_pbis([
            {"pbi_id": 1, "pbi_title": "P", "pbi_story_points": 3, "pbi_url": None,
             "task_id": None, "completed_work": None, "categorization": None},
            {"pbi_id": 1, "pbi_title": "P", "pbi_story_points": 3, "pbi_url": None,
             "task_id": 10, "completed_work": 5, "categorization": "Development"},
        ])
        wb = self._build([pbi])
        ws = wb["Summary"]
        merged_ranges = {str(r) for r in ws.merged_cells.ranges}
        block_len = len(summary.CATEGORY_ORDER)
        for col_letter in ("A", "B", "C", "D"):
            expected = f"{col_letter}2:{col_letter}{1 + block_len}"
            self.assertIn(expected, merged_ranges)
        wb.close()

    def test_six_category_rows_written_even_with_one_task(self):
        [pbi] = summary.summarize_pbis([
            {"pbi_id": 1, "pbi_title": "P", "pbi_story_points": 3, "pbi_url": None,
             "task_id": None, "completed_work": None, "categorization": None},
            {"pbi_id": 1, "pbi_title": "P", "pbi_story_points": 3, "pbi_url": None,
             "task_id": 10, "completed_work": 5, "categorization": "Bug Fixing"},
        ])
        wb = self._build([pbi])
        ws = wb["Summary"]
        category_names = [ws.cell(row=r, column=5).value for r in range(2, 2 + len(summary.CATEGORY_ORDER))]
        self.assertEqual(category_names, summary.CATEGORY_ORDER)
        wb.close()

    def test_category_percent_cell_displays_with_percent_sign(self):
        """Category % cells keep their raw numeric value (e.g. 100.0, not
        1.0) but use a custom number format that appends a literal "%" so
        Excel displays "100.0%" instead of a bare number.
        """
        [pbi] = summary.summarize_pbis([
            {"pbi_id": 1, "pbi_title": "P", "pbi_story_points": 3, "pbi_url": None,
             "task_id": None, "completed_work": None, "categorization": None},
            {"pbi_id": 1, "pbi_title": "P", "pbi_story_points": 3, "pbi_url": None,
             "task_id": 10, "completed_work": 5, "categorization": "Bug Fixing"},
        ])
        wb = self._build([pbi])
        ws = wb["Summary"]
        bug_fixing_row = 2 + summary.CATEGORY_ORDER.index("Bug Fixing")
        pct_cell = ws.cell(row=bug_fixing_row, column=7)
        self.assertEqual(pct_cell.value, 100.0)
        self.assertIn("%", pct_cell.number_format)
        wb.close()

    def test_yellow_fill_applied_for_missing_story_points_and_missing_hours(self):
        [pbi] = summary.summarize_pbis([
            {"pbi_id": 1, "pbi_title": "P", "pbi_story_points": "No_Story_Point", "pbi_url": None,
             "task_id": None, "completed_work": None, "categorization": None},
            {"pbi_id": 1, "pbi_title": "P", "pbi_story_points": "No_Story_Point", "pbi_url": None,
             "task_id": 10, "completed_work": "No_Hours", "categorization": "Development"},
        ])
        wb = self._build([pbi])
        ws = wb["Summary"]
        sp_cell = ws.cell(row=2, column=3)
        total_cell = ws.cell(row=2, column=4)
        self.assertEqual(sp_cell.fill.fgColor.rgb, "FFFFFF00")
        self.assertEqual(total_cell.fill.fgColor.rgb, "FFFFFF00")
        self.assertIn("some tasks/bugs missing hours", str(total_cell.value))
        wb.close()

    def test_no_highlight_when_story_points_and_hours_present(self):
        [pbi] = summary.summarize_pbis([
            {"pbi_id": 1, "pbi_title": "P", "pbi_story_points": 5, "pbi_url": None,
             "task_id": None, "completed_work": None, "categorization": None},
            {"pbi_id": 1, "pbi_title": "P", "pbi_story_points": 5, "pbi_url": None,
             "task_id": 10, "completed_work": 4, "categorization": "Development"},
        ])
        wb = self._build([pbi])
        ws = wb["Summary"]
        sp_cell = ws.cell(row=2, column=3)
        total_cell = ws.cell(row=2, column=4)
        self.assertNotEqual(getattr(sp_cell.fill.fgColor, "rgb", None), "FFFFFF00")
        self.assertNotEqual(getattr(total_cell.fill.fgColor, "rgb", None), "FFFFFF00")
        wb.close()

    def test_pie_chart_slice_count_matches_nonzero_categories(self):
        """The chart's data/category references must span only the
        nonzero-category rows (written into hidden helper columns), not
        the full six-row fixed-category block -- so zero-hour categories
        never become slices/labels and can't clutter/overlap the chart.
        """
        [pbi] = summary.summarize_pbis([
            {"pbi_id": 1, "pbi_title": "P", "pbi_story_points": 5, "pbi_url": None,
             "task_id": None, "completed_work": None, "categorization": None},
            {"pbi_id": 1, "pbi_title": "P", "pbi_story_points": 5, "pbi_url": None,
             "task_id": 10, "completed_work": 4, "categorization": "Development"},
            {"pbi_id": 1, "pbi_title": "P", "pbi_story_points": 5, "pbi_url": None,
             "task_id": 11, "completed_work": 2, "categorization": "Review"},
        ])
        wb = self._build([pbi])
        ws = wb["Summary"]
        self.assertEqual(len(ws._charts), 1)
        chart = ws._charts[0]
        data_ref = chart.series[0].val.numRef.f
        # Only 2 nonzero categories (Review, Development) -- the data
        # reference must span exactly 2 rows in the hidden helper column,
        # not all 6 fixed-category rows.
        helper_col_letter = get_column_letter(summary.CHART_HELPER_HOURS_COL)
        self.assertEqual(
            data_ref, f"'Summary'!${helper_col_letter}$2:${helper_col_letter}$3"
        )
        self.assertEqual(len(chart.series[0].data_points), 2)
        wb.close()

    def test_zero_hour_categories_are_not_written_to_chart_helper_columns(self):
        """Zero-hour categories must not appear in the hidden chart-helper
        columns at all (not even as a 0), confirming they truly cannot
        produce a slice/label -- only genuinely nonzero categories are
        written there.
        """
        [pbi] = summary.summarize_pbis([
            {"pbi_id": 1, "pbi_title": "P", "pbi_story_points": 5, "pbi_url": None,
             "task_id": None, "completed_work": None, "categorization": None},
            {"pbi_id": 1, "pbi_title": "P", "pbi_story_points": 5, "pbi_url": None,
             "task_id": 10, "completed_work": 4, "categorization": "Development"},
        ])
        wb = self._build([pbi])
        ws = wb["Summary"]
        helper_categories = [
            ws.cell(row=r, column=summary.CHART_HELPER_CATEGORY_COL).value
            for r in range(2, 2 + len(summary.CATEGORY_ORDER))
        ]
        self.assertEqual([c for c in helper_categories if c is not None], ["Development"])
        wb.close()

    def test_chart_helper_columns_are_hidden(self):
        wb = self._build([])
        ws = wb["Summary"]
        cat_col_letter = get_column_letter(summary.CHART_HELPER_CATEGORY_COL)
        hours_col_letter = get_column_letter(summary.CHART_HELPER_HOURS_COL)
        self.assertTrue(ws.column_dimensions[cat_col_letter].hidden)
        self.assertTrue(ws.column_dimensions[hours_col_letter].hidden)
        wb.close()

    def test_category_colors_are_identical_across_different_pbis_charts(self):
        """The same category must render in the exact same color on every
        PBI's pie chart (e.g. "Development" is always green), not just
        within a single PBI's own chart.
        """
        summaries = summary.summarize_pbis([
            {"pbi_id": 1, "pbi_title": "P1", "pbi_story_points": 5, "pbi_url": None,
             "task_id": None, "completed_work": None, "categorization": None},
            {"pbi_id": 1, "pbi_title": "P1", "pbi_story_points": 5, "pbi_url": None,
             "task_id": 10, "completed_work": 4, "categorization": "Development"},
            {"pbi_id": 1, "pbi_title": "P1", "pbi_story_points": 5, "pbi_url": None,
             "task_id": 11, "completed_work": 2, "categorization": "Bug Fixing"},
            {"pbi_id": 2, "pbi_title": "P2", "pbi_story_points": 3, "pbi_url": None,
             "task_id": None, "completed_work": None, "categorization": None},
            {"pbi_id": 2, "pbi_title": "P2", "pbi_story_points": 3, "pbi_url": None,
             "task_id": 20, "completed_work": 6, "categorization": "Development"},
            {"pbi_id": 2, "pbi_title": "P2", "pbi_story_points": 3, "pbi_url": None,
             "task_id": 21, "completed_work": 1, "categorization": "Review"},
        ])
        wb = self._build(summaries)
        ws = wb["Summary"]
        self.assertEqual(len(ws._charts), 2)

        def _colors_by_category(chart, pbi_summary):
            nonzero = [cat for cat in pbi_summary["categories"] if cat["hours"] > 0]
            return {
                nonzero[dp.idx]["category"]: dp.spPr.solidFill.srgbClr
                for dp in chart.series[0].data_points
            }

        colors_pbi1 = _colors_by_category(ws._charts[0], summaries[0])
        colors_pbi2 = _colors_by_category(ws._charts[1], summaries[1])
        common_categories = set(colors_pbi1) & set(colors_pbi2)
        self.assertEqual(common_categories, {"Development"})
        for category in common_categories:
            self.assertEqual(
                colors_pbi1[category], colors_pbi2[category],
                f"Category '{category}' has different colors across PBI charts",
            )
        for category, color in {**colors_pbi1, **colors_pbi2}.items():
            self.assertEqual(color, summary.category_color(category))
        wb.close()

    def test_chart_plots_hidden_helper_columns_despite_visible_cells_only_default(self):
        """Regression test: the chart's data source lives in hidden helper
        columns (by design), but openpyxl/Excel charts default to "plot
        visible cells only" (`visible_cells_only` / OOXML `plotVisOnly`
        both default True). Left at the default, Excel renders the chart
        completely empty -- no slices at all -- because none of its source
        cells are visible. `visible_cells_only` must be explicitly False on
        every chart this script creates.
        """
        [pbi] = summary.summarize_pbis([
            {"pbi_id": 1, "pbi_title": "P", "pbi_story_points": 5, "pbi_url": None,
             "task_id": None, "completed_work": None, "categorization": None},
            {"pbi_id": 1, "pbi_title": "P", "pbi_story_points": 5, "pbi_url": None,
             "task_id": 10, "completed_work": 4, "categorization": "Development"},
        ])
        wb = self._build([pbi])
        ws = wb["Summary"]
        self.assertEqual(len(ws._charts), 1)
        self.assertFalse(ws._charts[0].visible_cells_only)
        wb.close()

    def test_chart_has_bottom_legend_naming_categories_and_percentage_only_labels(self):
        """Each slice's data label shows ONLY its percentage -- not the
        category name or hours value, since repeating a long category name
        (e.g. "Requirement Study") directly on a slice caused confusing,
        overlapping label text. To avoid a bare, unidentifiable "75%"
        slice, the chart keeps a compact bottom legend so the viewer can
        still see which category each color/slice belongs to.
        """
        [pbi] = summary.summarize_pbis([
            {"pbi_id": 1, "pbi_title": "P", "pbi_story_points": 5, "pbi_url": None,
             "task_id": None, "completed_work": None, "categorization": None},
            {"pbi_id": 1, "pbi_title": "P", "pbi_story_points": 5, "pbi_url": None,
             "task_id": 10, "completed_work": 4, "categorization": "Development"},
        ])
        wb = self._build([pbi])
        ws = wb["Summary"]
        chart = ws._charts[0]
        self.assertIsNotNone(chart.legend)
        self.assertEqual(chart.legend.position, "b")
        self.assertFalse(chart.dataLabels.showCatName)
        self.assertFalse(chart.dataLabels.showVal)
        self.assertTrue(chart.dataLabels.showPercent)
        wb.close()

    def test_slice_labels_do_not_show_series_name(self):
        """Each slice's data label must explicitly suppress the series name
        -- without this, Excel renders a stray "Series1" label on the
        chart even though the legend itself is removed (the series was
        never given its own title, so "Series1" is Excel's placeholder).
        """
        [pbi] = summary.summarize_pbis([
            {"pbi_id": 1, "pbi_title": "P", "pbi_story_points": 5, "pbi_url": None,
             "task_id": None, "completed_work": None, "categorization": None},
            {"pbi_id": 1, "pbi_title": "P", "pbi_story_points": 5, "pbi_url": None,
             "task_id": 10, "completed_work": 4, "categorization": "Development"},
        ])
        wb = self._build([pbi])
        ws = wb["Summary"]
        chart = ws._charts[0]
        self.assertFalse(chart.dataLabels.showSerName)
        self.assertFalse(chart.dataLabels.showLegendKey)
        wb.close()

    def test_no_chart_and_no_category_rows_for_pbi_without_children(self):
        [pbi] = summary.summarize_pbis([
            {"pbi_id": 1, "pbi_title": "P", "pbi_story_points": 5, "pbi_url": None,
             "task_id": None, "completed_work": None, "categorization": None},
        ])
        wb = self._build([pbi])
        ws = wb["Summary"]
        self.assertEqual(len(ws._charts), 0)
        self.assertIsNone(ws.cell(row=2, column=5).value)
        self.assertEqual(ws.cell(row=2, column=4).value, 0)
        wb.close()

    def test_no_chart_when_all_categories_are_zero_hours(self):
        """A PBI with children whose Completed Work all sum to 0 across
        every category produces no visible slices, so no chart is added."""
        [pbi] = summary.summarize_pbis([
            {"pbi_id": 1, "pbi_title": "P", "pbi_story_points": 5, "pbi_url": None,
             "task_id": None, "completed_work": None, "categorization": None},
            {"pbi_id": 1, "pbi_title": "P", "pbi_story_points": 5, "pbi_url": None,
             "task_id": 10, "completed_work": 0, "categorization": "Development"},
        ])
        wb = self._build([pbi])
        ws = wb["Summary"]
        self.assertEqual(len(ws._charts), 0)
        wb.close()


# ---------------------------------------------------------------------------
# Tests against the LATEST actually-generated summary workbook: these only
# need the workbook file (no Azure DevOps access) and are skipped with a
# message if none has been generated yet.
# ---------------------------------------------------------------------------

class LatestSummaryStructureTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.report_path = _latest_summary_report_path()
        if cls.report_path is None:
            raise unittest.SkipTest(
                f"No generated summary workbook found in {REPORT_DIR}. "
                "Run pbi_bug_task_summary.py first, then re-run these tests."
            )
        cls.workbook = load_workbook(cls.report_path, data_only=True)

    @classmethod
    def tearDownClass(cls):
        cls.workbook.close()

    def _category_hours_for_pbi(self, sheet_name, pbi_id, category):
        """Recompute the total Completed Work hours for `pbi_id`'s tasks in
        `category` directly from `sheet_name` (treating No_Hours/blank as
        0), independent of summarize_pbis()."""
        if sheet_name not in self.workbook.sheetnames:
            return None
        ws = self.workbook[sheet_name]
        header_row = summary._find_header_row(ws)
        cols = summary._column_map(header_row)
        total = 0.0
        for r in range(header_row[0].row + 1, ws.max_row + 1):
            if ws.cell(row=r, column=cols["PBI ID"]).value != pbi_id:
                continue
            if ws.cell(row=r, column=cols["Task ID"]).value is None:
                continue
            if ws.cell(row=r, column=cols["Task Categorization"]).value != category:
                continue
            total += summary.hours_to_number(
                ws.cell(row=r, column=cols["Completed Work (hrs)"]).value
            )
        return total

    def test_zero_category_rows_match_recomputed_hours_in_all_and_team_tab(self):
        """For every Summary category sub-row showing Category Hours == 0,
        independently recompute that category's total Completed Work hours
        for the same PBI directly from the "All" tab and from its matching
        team tab (FPSO/Foundation) -- treating No_Hours/blank as 0, exactly
        like summarize_pbis() does -- and assert both recomputations are
        also 0. This confirms the Summary's 0 is real (no task in that
        category was silently dropped or mis-summed), not merely that no
        task of that category exists at all (a task can exist with
        No_Hours logged and still correctly sum to 0).

        Example: if the Summary shows PBI 34236 / "Review" = 0 hours, but
        recomputing directly from the "All" tab finds a Review task with
        Completed Work = 3, this test fails identifying PBI 34236,
        category "Review", tab "All", and the mismatched recomputed total.
        """
        ws_summary = self.workbook["Summary"]
        ws_all = self.workbook["All"]
        header_row = summary._find_header_row(ws_all)
        cols = summary._column_map(header_row)

        team_by_pbi = {}
        for r in range(header_row[0].row + 1, ws_all.max_row + 1):
            pbi_id = ws_all.cell(row=r, column=cols["PBI ID"]).value
            if pbi_id is not None:
                team_by_pbi[pbi_id] = ws_all.cell(row=r, column=cols["Team"]).value

        failures = []
        current_pbi = None
        for r in range(2, ws_summary.max_row + 1):
            pbi_cell = ws_summary.cell(row=r, column=1).value
            if pbi_cell is not None:
                current_pbi = pbi_cell
            category = ws_summary.cell(row=r, column=5).value
            hours = ws_summary.cell(row=r, column=6).value
            if category is None or current_pbi is None or hours != 0:
                continue

            recomputed_all = self._category_hours_for_pbi("All", current_pbi, category)
            if recomputed_all not in (0, 0.0, None):
                failures.append((current_pbi, category, "All", recomputed_all))

            team = team_by_pbi.get(current_pbi)
            if team:
                recomputed_team = self._category_hours_for_pbi(team, current_pbi, category)
                if recomputed_team not in (0, 0.0, None):
                    failures.append((current_pbi, category, team, recomputed_team))

        self.assertFalse(
            failures,
            "Summary shows 0 hours for a category whose recomputed total "
            "from All/team tab is actually nonzero "
            "(pbi_id, category, tab, recomputed_hours):\n"
            + "\n".join(str(f) for f in failures),
        )

    def test_no_row_on_any_tab_is_assigned_to_himanshu_pathak(self):
        """Reads every tab in the generated workbook and asserts no row is
        Assigned To "Himanshu Pathak" -- mirrors the script's own
        pre-flight abort check, re-verified against the final artifact."""
        failures = []
        for worksheet in self.workbook.worksheets:
            header_row = next(
                (row for row in worksheet.iter_rows() if "Assigned To" in [c.value for c in row]),
                None,
            )
            if header_row is None:
                continue
            cols = {cell.value: cell.column for cell in header_row if cell.value}
            for r in range(header_row[0].row + 1, worksheet.max_row + 1):
                assigned_to = worksheet.cell(row=r, column=cols["Assigned To"]).value
                if _normalize_assignee(assigned_to) != _normalize_assignee("Himanshu Pathak"):
                    continue
                failures.append({
                    "tab": worksheet.title,
                    "pbi_id": worksheet.cell(row=r, column=cols.get("PBI ID", 0)).value if "PBI ID" in cols else None,
                    "task_id": worksheet.cell(row=r, column=cols.get("Task ID", 0)).value if "Task ID" in cols else None,
                })
        self.assertFalse(
            failures,
            f"Workbook '{self.report_path}' assigns one or more items to "
            f"Himanshu Pathak: {failures}",
        )

    def test_category_percentages_sum_to_roughly_100_per_pbi(self):
        ws = self.workbook["Summary"]
        pbi_row_indices = [r for r in range(2, ws.max_row + 1) if ws.cell(row=r, column=1).value is not None]
        for i, start in enumerate(pbi_row_indices):
            end = pbi_row_indices[i + 1] if i + 1 < len(pbi_row_indices) else ws.max_row + 1
            if ws.cell(row=start, column=5).value is None:
                continue  # PBI-only row, no categories to check
            total_pct = sum(
                ws.cell(row=r, column=7).value or 0
                for r in range(start, min(start + len(summary.CATEGORY_ORDER), end))
            )
            self.assertAlmostEqual(total_pct, 100.0, delta=0.5, msg=f"PBI row {start}")


# ---------------------------------------------------------------------------
# Live Azure DevOps API checks: require both a generated summary workbook
# and a configured PAT. Skipped with an explanation otherwise.
# ---------------------------------------------------------------------------

class LatestSummaryLiveApiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.report_path = _latest_summary_report_path()
        if cls.report_path is None:
            raise unittest.SkipTest(
                f"No generated summary workbook found in {REPORT_DIR}. "
                "Run pbi_bug_task_summary.py first, then re-run these tests."
            )
        pat_token = report.load_pat_token()
        if not pat_token:
            raise unittest.SkipTest(
                "Set AZURE_DEVOPS_PAT (or populate localconfig) to validate "
                "the latest summary workbook against the live Azure DevOps API."
            )
        cls.session = report.get_session(pat_token)
        cls.workbook = load_workbook(cls.report_path, data_only=True)

    @classmethod
    def tearDownClass(cls):
        cls.workbook.close()

    def _team_source_url(self, team):
        ws = self.workbook[team]
        return ws.cell(row=1, column=2).value

    def _pbi_team_map(self):
        ws = self.workbook["All"]
        header_row = summary._find_header_row(ws)
        cols = summary._column_map(header_row)
        mapping = {}
        for r in range(header_row[0].row + 1, ws.max_row + 1):
            pbi_id = ws.cell(row=r, column=cols["PBI ID"]).value
            if pbi_id is not None:
                mapping[pbi_id] = ws.cell(row=r, column=cols["Team"]).value
        return mapping

    def _summary_pbi_rows(self):
        """Yield (pbi_id, pbi_url, total_hours_display, has_children) for
        every PBI block in the Summary tab."""
        ws = self.workbook["Summary"]
        results = []
        for r in range(2, ws.max_row + 1):
            pbi_id = ws.cell(row=r, column=1).value
            if pbi_id is None:
                continue
            results.append({
                "pbi_id": pbi_id,
                "total_hours_display": ws.cell(row=r, column=4).value,
                "has_children": ws.cell(row=r, column=5).value is not None,
            })
        return results

    def test_every_summary_pbi_is_returned_by_its_team_query(self):
        """Every PBI ID in the Summary tab must be returned right now by
        its team's saved query (URL read from row 1 of that team's tab).

        Example: if PBI 34236 is shown in the Summary but re-running the
        FPSO query live no longer returns 34236, this test fails reporting
        team "FPSO", the query URL used, and the missing PBI.
        """
        pbi_team = self._pbi_team_map()
        query_ids_by_team = {}
        for team in TEAM_TABS:
            if team not in self.workbook.sheetnames:
                continue
            source_url = self._team_source_url(team)
            info = report.parse_query_url(source_url)
            ids = report.run_wiql_query_by_id(
                self.session, info["org"], info["project"], info["query_id"]
            )
            query_ids_by_team[team] = (set(ids), source_url)

        failures = []
        for row in self._summary_pbi_rows():
            pbi_id = row["pbi_id"]
            team = pbi_team.get(pbi_id)
            if team not in query_ids_by_team:
                failures.append({
                    "pbi_id": pbi_id, "team": team, "query_url": None,
                    "reason": "PBI's team could not be resolved from the All tab",
                })
                continue
            query_ids, source_url = query_ids_by_team[team]
            if pbi_id not in query_ids:
                failures.append({
                    "pbi_id": pbi_id, "team": team, "query_url": source_url,
                    "reason": "Query no longer returns this PBI",
                })

        self.assertFalse(
            failures,
            "Summary contains PBI(s) not returned by their team's live query:\n"
            + "\n".join(str(f) for f in failures),
        )

    def test_total_hours_match_live_recomputation(self):
        """Recompute each PBI's total Completed Work hours directly from a
        fresh fetch of its current Task/Bug children (treating missing
        hours as 0) and assert it matches the Summary's total-hours figure.
        """
        failures = []
        for row in self._summary_pbi_rows():
            pbi_id = row["pbi_id"]
            ws = self.workbook["Summary"]
            pbi_cell_row = next(
                r for r in range(2, ws.max_row + 1)
                if ws.cell(row=r, column=1).value == pbi_id
            )
            pbi_url = ws.cell(row=pbi_cell_row, column=1).hyperlink
            pbi_url = pbi_url.target if pbi_url else None
            parsed = _parse_work_item_url(pbi_url) if pbi_url else None
            if parsed is None:
                continue
            org, project, _id = parsed

            item = report.get_work_items_full(self.session, org, project, [pbi_id])
            if not item:
                continue
            child_ids = report.collect_child_ids(item)
            live_total = 0.0
            if child_ids:
                children = report.get_work_items_full(self.session, org, project, child_ids)
                children = report.filter_types(children, report.CHILD_TYPES)
                for child in children:
                    live_total += summary.hours_to_number(
                        child.get("fields", {}).get(report.COMPLETED_WORK_FIELD)
                    )

            expected_display = _parse_total_hours_display(row["total_hours_display"])
            if expected_display is None or abs(expected_display - live_total) > 0.01:
                failures.append({
                    "pbi_id": pbi_id,
                    "summary_total": row["total_hours_display"],
                    "live_total": live_total,
                })

        self.assertFalse(
            failures,
            "Summary total-hours do not match a live recomputation "
            f"(pbi_id, summary_total, live_total):\n"
            + "\n".join(str(f) for f in failures),
        )

    def test_pbis_shown_with_zero_children_are_genuinely_empty_live(self):
        """For every PBI the Summary shows with no Task/Bug children,
        independently re-query Azure DevOps for that PBI's current child
        work items and assert the live result is indeed empty.
        """
        zero_child_rows = [row for row in self._summary_pbi_rows() if not row["has_children"]]
        if not zero_child_rows:
            self.skipTest("No PBI with zero children in the latest Summary tab to verify.")

        ws = self.workbook["Summary"]
        failures = []
        for row in zero_child_rows:
            pbi_id = row["pbi_id"]
            pbi_cell_row = next(
                r for r in range(2, ws.max_row + 1)
                if ws.cell(row=r, column=1).value == pbi_id
            )
            pbi_url = ws.cell(row=pbi_cell_row, column=1).hyperlink
            pbi_url = pbi_url.target if pbi_url else None
            parsed = _parse_work_item_url(pbi_url) if pbi_url else None
            if parsed is None:
                continue
            org, project, _id = parsed

            item = report.get_work_items_full(self.session, org, project, [pbi_id])
            if not item:
                continue
            child_ids = report.collect_child_ids(item)
            if child_ids:
                failures.append({"pbi_id": pbi_id, "live_child_ids": child_ids})

        self.assertFalse(
            failures,
            "Summary shows a PBI with zero children, but the live API "
            f"currently returns child work items for it: {failures}",
        )

    def _chart_legend_categories_for_pbi(self, pbi_cell_row):
        """Read the hidden chart-helper category column for one PBI's
        block, starting at `pbi_cell_row` -- i.e. exactly the nonzero
        categories that feed that PBI's pie chart (slices + legend)."""
        ws = self.workbook["Summary"]
        categories = []
        r = pbi_cell_row
        while True:
            value = ws.cell(row=r, column=summary.CHART_HELPER_CATEGORY_COL).value
            if value is None:
                break
            categories.append(value)
            r += 1
        return categories

    def test_chart_legend_categories_match_live_recomputed_categorization(self):
        """The chart's legend must name exactly the categories that a fresh
        recomputation from live Azure DevOps data says are nonzero for
        that PBI -- i.e. a viewer reading "75%" next to a legend entry is
        actually being told the correct category, verified against live
        data rather than anything already written into the workbook.

        For each PBI shown in the Summary with Task/Bug children, this
        independently re-fetches that PBI's current child work items from
        Azure DevOps, recomputes each child's category with
        `categorize_task` (the same rule-based logic the generator script
        uses) and its Completed Work hours (treating missing hours as 0),
        and derives the live set of categories with nonzero total hours.
        That live-derived set must exactly match the category names found
        in the chart's hidden helper column for that PBI -- the exact data
        that feeds the chart's slices and its bottom legend.

        Example: if a fresh fetch shows PBI 34236 currently has hours
        logged under "Review" and "Development" only, but the chart's
        legend data lists "Review", "Development", and "QA / Testing" (or
        is missing one of the two that should be there), this test fails
        identifying PBI 34236 with both the live-expected and the
        chart's-actual category sets.
        """
        ws = self.workbook["Summary"]
        rows_with_children = [row for row in self._summary_pbi_rows() if row["has_children"]]
        if not rows_with_children:
            self.skipTest("No PBI with Task/Bug children in the latest Summary tab to verify.")

        failures = []
        for row in rows_with_children:
            pbi_id = row["pbi_id"]
            pbi_cell_row = next(
                r for r in range(2, ws.max_row + 1)
                if ws.cell(row=r, column=1).value == pbi_id
            )
            pbi_url = ws.cell(row=pbi_cell_row, column=1).hyperlink
            pbi_url = pbi_url.target if pbi_url else None
            parsed = _parse_work_item_url(pbi_url) if pbi_url else None
            if parsed is None:
                continue
            org, project, _id = parsed

            item = report.get_work_items_full(self.session, org, project, [pbi_id])
            if not item:
                continue
            child_ids = report.collect_child_ids(item)
            if not child_ids:
                continue
            children = report.get_work_items_full(self.session, org, project, child_ids)
            children = report.filter_types(children, report.CHILD_TYPES)

            live_category_hours = {}
            for child in children:
                fields = child.get("fields", {})
                wi_type = fields.get("System.WorkItemType", "")
                title = fields.get("System.Title", "")
                description = report.strip_html(fields.get("System.Description", ""))
                category = report.categorize_task(wi_type, title, description)
                hours = summary.hours_to_number(fields.get(report.COMPLETED_WORK_FIELD))
                live_category_hours[category] = live_category_hours.get(category, 0.0) + hours

            live_nonzero_categories = {
                cat for cat, hours in live_category_hours.items() if hours > 0
            }
            chart_categories = set(self._chart_legend_categories_for_pbi(pbi_cell_row))

            if live_nonzero_categories != chart_categories:
                failures.append({
                    "pbi_id": pbi_id,
                    "live_expected_categories": sorted(live_nonzero_categories),
                    "chart_actual_categories": sorted(chart_categories),
                })

        self.assertFalse(
            failures,
            "Chart legend/category data does not match a live recomputation "
            "of nonzero categories (pbi_id, live_expected_categories, "
            "chart_actual_categories):\n"
            + "\n".join(str(f) for f in failures),
        )

    def _summary_category_percentages_for_pbi(self, pbi_cell_row):
        """Read the visible category breakdown rows for one PBI's block,
        starting at `pbi_cell_row` -- returns {category_name: percent} as
        literally stored in the Summary tab's "Category %" column (column
        G), for every category row written for that PBI (including
        zero-hour ones)."""
        ws = self.workbook["Summary"]
        percentages = {}
        r = pbi_cell_row
        while True:
            category = ws.cell(row=r, column=5).value
            if category is None:
                break
            percentages[category] = ws.cell(row=r, column=7).value
            r += 1
        return percentages

    def test_category_percentages_match_live_recomputation(self):
        """Each PBI's "Category %" values in the Summary tab must match a
        fresh, independent recomputation from live Azure DevOps data,
        using the exact same formula the generator script itself applies
        in `summarize_pbis()`:

            percent(category) = hours(category) / total_hours * 100
                                 (rounded to 1 decimal place)
            percent(category) = 0  when total_hours == 0

        Data used for the live recomputation (independent of anything
        already written into the workbook, to avoid a circular check):
          1. For each PBI shown with children in the Summary tab, its
             current child Task/Bug work items are re-fetched fresh from
             Azure DevOps (via the PBI's own hyperlink -> org/project ->
             `get_work_items_full` -> `collect_child_ids` -> fetch
             children -> filter to Task/Bug types).
          2. Each child's category is recomputed with the real
             `categorize_task(work_item_type, title, description)` logic,
             and its hours with `hours_to_number(Completed Work field)`
             (treating missing/"No_Hours" as 0) -- both computed from
             scratch, not read from the "All" tab or anywhere else in the
             workbook.
          3. `hours(category)` = sum of step-2 hours for children in that
             category; `total_hours` = sum across ALL of that PBI's
             children regardless of category.

        The resulting live percentage (rounded to 1 decimal, matching the
        Summary's own rounding) is compared against the literal numeric
        value already stored in that PBI's "Category %" cell, for every
        category row (zero and nonzero), with a small tolerance (0.1) to
        absorb rounding.

        Example: if a fresh fetch shows PBI 34236 has 6 hours of
        "Development" out of 10 total hours (60.0%), but the Summary's
        "Category %" cell for PBI 34236 / "Development" reads 50.0, this
        test fails identifying PBI 34236, category "Development",
        live_percent 60.0, and summary_percent 50.0.
        """
        ws = self.workbook["Summary"]
        rows_with_children = [row for row in self._summary_pbi_rows() if row["has_children"]]
        if not rows_with_children:
            self.skipTest("No PBI with Task/Bug children in the latest Summary tab to verify.")

        failures = []
        for row in rows_with_children:
            pbi_id = row["pbi_id"]
            pbi_cell_row = next(
                r for r in range(2, ws.max_row + 1)
                if ws.cell(row=r, column=1).value == pbi_id
            )
            pbi_url = ws.cell(row=pbi_cell_row, column=1).hyperlink
            pbi_url = pbi_url.target if pbi_url else None
            parsed = _parse_work_item_url(pbi_url) if pbi_url else None
            if parsed is None:
                continue
            org, project, _id = parsed

            item = report.get_work_items_full(self.session, org, project, [pbi_id])
            if not item:
                continue
            child_ids = report.collect_child_ids(item)
            if not child_ids:
                continue
            children = report.get_work_items_full(self.session, org, project, child_ids)
            children = report.filter_types(children, report.CHILD_TYPES)

            live_category_hours = {}
            live_total_hours = 0.0
            for child in children:
                fields = child.get("fields", {})
                wi_type = fields.get("System.WorkItemType", "")
                title = fields.get("System.Title", "")
                description = report.strip_html(fields.get("System.Description", ""))
                category = report.categorize_task(wi_type, title, description)
                hours = summary.hours_to_number(fields.get(report.COMPLETED_WORK_FIELD))
                live_category_hours[category] = live_category_hours.get(category, 0.0) + hours
                live_total_hours += hours

            summary_percentages = self._summary_category_percentages_for_pbi(pbi_cell_row)

            for category, summary_percent in summary_percentages.items():
                live_hours = live_category_hours.get(category, 0.0)
                live_percent = (
                    round(live_hours / live_total_hours * 100.0, 1)
                    if live_total_hours > 0 else 0.0
                )
                if summary_percent is None or abs(live_percent - summary_percent) > 0.1:
                    failures.append({
                        "pbi_id": pbi_id,
                        "category": category,
                        "live_percent": live_percent,
                        "summary_percent": summary_percent,
                    })

        self.assertFalse(
            failures,
            "Summary Category % does not match a live recomputation using "
            "percent = hours / total_hours * 100 "
            "(pbi_id, category, live_percent, summary_percent):\n"
            + "\n".join(str(f) for f in failures),
        )


if __name__ == "__main__":
    unittest.main()
