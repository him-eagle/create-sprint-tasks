"""
PBI/Task Summary Report Generator
=================================

Reads an already-generated ``PBI_Bug_Task_Report_*.xlsx`` workbook (produced
by ``pbi_bug_task_report.py``) and fills in its empty "Summary" tab with a
per-PBI rollup:

  * PBI ID (hyperlinked), Title, Story Points, and Total Hours (sum of
    Completed Work across that PBI's Task/Bug rows).
  * A fixed six-category breakdown (Bug Fixing, QA / Testing, Requirement
    Study, Review, Design / Documentation, Development) showing hours and
    percentage for every category, even ones the PBI has no tasks in
    (shown as 0 / 0%).
  * An inline pie chart per PBI (percentage-only slice labels plus a
    compact bottom legend naming each nonzero category), placed beside
    that PBI's own row block.

Missing Story Points (``No_Story_Point``) and a PBI whose total-hours
figure includes at least one task/bug with no logged hours are both
highlighted yellow in the Summary tab so they're easy to spot.

IMPORTANT: This script never modifies the input workbook. It loads it,
writes the Summary sheet into the in-memory copy, and saves the result to a
brand-new file in ``Reports\\Pbi_bug_tasks`` -- the original file on disk is
left untouched.

Before building the Summary, the script validates:
  * The input file name starts with ``PBI_Bug_Task_Report_``.
  * The workbook has ``Summary``, ``All``, ``FPSO``, and ``Foundation`` tabs.
  * The ``Summary`` tab is still empty (if not, this is likely a workbook
    that has already been processed -- the script aborts rather than
    overwriting/duplicating content).
  * No row in ``All``/``FPSO``/``Foundation`` is Assigned To "Himanshu
    Pathak" -- if any are found, the script aborts and lists them, asking
    the user to remove those entries from the workbook first and re-run.

Usage:
    python pbi_bug_task_summary.py

You will be prompted for the path to the input .xlsx file.
"""

import os
import re
import sys
import datetime

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.legend import Legend
from openpyxl.chart.marker import DataPoint
from openpyxl.chart.shapes import GraphicalProperties

# ---------------------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------------------
SCRIPT_DIR = r"C:\Utilities\SprintTask"
OUTPUT_DIR = SCRIPT_DIR + r"\Reports\Pbi_bug_tasks"

INPUT_FILENAME_PREFIX = "PBI_Bug_Task_Report_"
REQUIRED_TABS = ["Summary", "All", "FPSO", "Foundation"]
BLOCKED_ASSIGNEE = "Himanshu Pathak"

# Fixed category list, in display order. Matches categorize_task()'s
# possible outputs in pbi_bug_task_report.py / sprint_tasks_from_queries.py.
# Every PBI gets a sub-row for every one of these, even when its hours are 0.
CATEGORY_ORDER = [
    "Bug Fixing",
    "QA / Testing",
    "Requirement Study",
    "Review",
    "Design / Documentation",
    "Development",
]

NO_STORY_POINT_PLACEHOLDER = "No_Story_Point"
NO_HOURS_PLACEHOLDER = "No_Hours"
MISSING_HOURS_NOTE = "some tasks/bugs missing hours"

# Rows left blank below each charted PBI block so its inline pie chart has
# room and doesn't visually overlap the next PBI's rows/chart. Sized with
# some margin for the chart's bottom legend (chart.height below).
CHART_SPACER_ROWS = 14

YELLOW_FILL = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")

# Fixed color per category (hex RGB, no leading '#') so the same category
# always renders in the same color across every PBI's pie chart. Picked for
# clear visual contrast; edit freely to taste.
CATEGORY_COLORS = {
    "Bug Fixing": "C00000",             # dark red
    "QA / Testing": "ED7D31",           # orange
    "Requirement Study": "FFC000",      # gold
    "Review": "5B9BD5",                 # blue
    "Design / Documentation": "7030A0", # purple
    "Development": "70AD47",            # green
}
# Fallback palette for any category outside the fixed list (e.g. stale data),
# cycled through in encounter order so each still gets a consistent color.
FALLBACK_CATEGORY_COLORS = ["808080", "00B0F0", "FF66CC", "A9D18E"]

# Hidden helper columns holding ONLY each PBI's nonzero categories
# (contiguous, in CATEGORY_ORDER order) so pie charts are built from just
# those rows -- zero-hour categories never become slices/labels on the
# chart at all, which avoids overlapping "Category 0 0%" labels. The
# visible category table (columns E/F/G) still lists all six categories,
# including the zero ones, unaffected by this.
CHART_HELPER_CATEGORY_COL = 50
CHART_HELPER_HOURS_COL = 51

SUMMARY_HEADERS = [
    "PBI ID",
    "PBI Title",
    "PBI Story Points",
    "Total Hours",
    "Task Categorization",
    "Category Hours",
    "Category %",
]

TIMESTAMP_SUFFIX_RE = re.compile(r"_(\d{8}_\d{6})\.xlsx$", re.IGNORECASE)


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------

def normalize_assignee(value):
    """Normalize case/whitespace so assignee-name comparisons are robust to
    casing and spacing variants (e.g. "Himanshu Pathak" vs "HimanshuPathak")."""
    return "".join(str(value or "").split()).casefold()


def validate_input_filename(path):
    """Raise ValueError unless the file's base name starts with the expected
    PBI_Bug_Task_Report_ prefix."""
    name = os.path.basename(path)
    if not name.startswith(INPUT_FILENAME_PREFIX):
        raise ValueError(
            f"Input file name '{name}' does not start with "
            f"'{INPUT_FILENAME_PREFIX}'. Please supply a workbook produced by "
            "pbi_bug_task_report.py."
        )


def hours_to_number(value):
    """Convert a Completed Work cell value to a number, treating the
    No_Hours placeholder and blanks as 0. An explicit 0 stays 0."""
    if value is None or value == "" or value == NO_HOURS_PLACEHOLDER:
        return 0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def is_missing_hours(value):
    return value is None or value == "" or value == NO_HOURS_PLACEHOLDER


def format_number(value):
    """Render a number without a trailing .0 for whole numbers."""
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def category_color(category):
    """Return the fixed hex color for a category so the same category
    always renders in the same color across every PBI's pie chart. Falls
    back to a cycled palette for any category outside the fixed
    CATEGORY_COLORS list (e.g. stale/unexpected data)."""
    if category in CATEGORY_COLORS:
        return CATEGORY_COLORS[category]
    index = hash(category) % len(FALLBACK_CATEGORY_COLORS)
    return FALLBACK_CATEGORY_COLORS[index]


def _find_header_row(ws):
    """Return the list of Cell objects for the row containing both
    "PBI ID" and "Task ID" headers, or raise ValueError if not found."""
    for row in ws.iter_rows():
        values = [cell.value for cell in row]
        if "PBI ID" in values and "Task ID" in values:
            return list(row)
    raise ValueError(
        f"Could not find a header row with 'PBI ID'/'Task ID' columns in "
        f"sheet '{ws.title}'."
    )


def _column_map(header_row):
    return {cell.value: cell.column for cell in header_row if cell.value}


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------

def validate_workbook_structure(workbook):
    """Ensure the required tabs exist and the Summary tab is still empty."""
    missing = [tab for tab in REQUIRED_TABS if tab not in workbook.sheetnames]
    if missing:
        raise ValueError(
            f"Workbook is missing required tab(s): {', '.join(missing)}. "
            f"Expected tabs: {', '.join(REQUIRED_TABS)}."
        )

    summary_ws = workbook["Summary"]
    non_empty = any(
        cell.value is not None
        for row in summary_ws.iter_rows()
        for cell in row
    )
    if non_empty:
        raise ValueError(
            "The 'Summary' tab already contains content, which means this "
            "workbook has likely already been processed by this script (or "
            "edited). Please remove the existing Summary tab entries first "
            "(or supply a fresh, unprocessed PBI_Bug_Task_Report_*.xlsx file "
            "straight from pbi_bug_task_report.py), then re-run this script."
        )


def find_blocked_assignee_rows(workbook, blocked_name=BLOCKED_ASSIGNEE):
    """Scan All/FPSO/Foundation for rows Assigned To `blocked_name`.

    Returns a list of dicts describing each violation (sheet, PBI ID,
    Task/Bug ID, title), empty if none are found.
    """
    target = normalize_assignee(blocked_name)
    violations = []
    for sheet_name in ("All", "FPSO", "Foundation"):
        if sheet_name not in workbook.sheetnames:
            continue
        ws = workbook[sheet_name]
        try:
            header_row = _find_header_row(ws)
        except ValueError:
            continue
        cols = _column_map(header_row)
        if "Assigned To" not in cols:
            continue

        data_start = header_row[0].row + 1
        for r in range(data_start, ws.max_row + 1):
            assigned_to = ws.cell(row=r, column=cols["Assigned To"]).value
            if normalize_assignee(assigned_to) != target:
                continue
            violations.append({
                "sheet": sheet_name,
                "pbi_id": ws.cell(row=r, column=cols["PBI ID"]).value if "PBI ID" in cols else None,
                "task_id": ws.cell(row=r, column=cols["Task ID"]).value if "Task ID" in cols else None,
                "task_title": ws.cell(row=r, column=cols["Task Title"]).value if "Task Title" in cols else None,
            })
    return violations


# ---------------------------------------------------------------------------
# Reading the "All" tab
# ---------------------------------------------------------------------------

def read_all_tab(workbook):
    """Read every data row of the "All" tab into plain dicts, preserving
    each PBI ID cell's hyperlink target (if any)."""
    ws = workbook["All"]
    header_row = _find_header_row(ws)
    cols = _column_map(header_row)
    data_start = header_row[0].row + 1

    rows = []
    for r in range(data_start, ws.max_row + 1):
        pbi_id = ws.cell(row=r, column=cols["PBI ID"]).value
        if pbi_id is None:
            continue
        pbi_cell = ws.cell(row=r, column=cols["PBI ID"])
        rows.append({
            "pbi_id": pbi_id,
            "pbi_title": ws.cell(row=r, column=cols["PBI Title"]).value,
            "pbi_story_points": ws.cell(row=r, column=cols["PBI Story Points"]).value,
            "pbi_url": pbi_cell.hyperlink.target if pbi_cell.hyperlink else None,
            "task_id": ws.cell(row=r, column=cols["Task ID"]).value,
            "completed_work": ws.cell(row=r, column=cols["Completed Work (hrs)"]).value,
            "categorization": ws.cell(row=r, column=cols["Task Categorization"]).value,
        })
    return rows


# ---------------------------------------------------------------------------
# Per-PBI rollup
# ---------------------------------------------------------------------------

def summarize_pbis(all_rows):
    """Group the All-tab rows by PBI and compute the rollups needed for the
    Summary tab. Returns a list of dicts sorted by PBI ID.

    Each dict has: pbi_id, pbi_title, pbi_story_points, pbi_url,
    total_hours, missing_hours (bool), has_children (bool), and
    categories (list of {"category", "hours", "percent"} covering every
    fixed category, in CATEGORY_ORDER, plus any unexpected category found
    in the data appended at the end).
    """
    pbis = {}
    order = []
    for row in all_rows:
        pbi_id = row["pbi_id"]
        if pbi_id not in pbis:
            pbis[pbi_id] = {
                "pbi_id": pbi_id,
                "pbi_title": row["pbi_title"],
                "pbi_story_points": row["pbi_story_points"],
                "pbi_url": row["pbi_url"],
                "tasks": [],
            }
            order.append(pbi_id)
        if row["task_id"] is not None:
            pbis[pbi_id]["tasks"].append(row)

    summaries = []
    for pbi_id in sorted(order):
        pbi = pbis[pbi_id]
        tasks = pbi["tasks"]
        total_hours = sum(hours_to_number(t["completed_work"]) for t in tasks)
        missing_hours = any(is_missing_hours(t["completed_work"]) for t in tasks)

        category_hours = {cat: 0.0 for cat in CATEGORY_ORDER}
        for t in tasks:
            cat = t["categorization"] or "Development"
            category_hours.setdefault(cat, 0.0)
            category_hours[cat] += hours_to_number(t["completed_work"])

        category_order = CATEGORY_ORDER + sorted(
            cat for cat in category_hours if cat not in CATEGORY_ORDER
        )

        categories = []
        for cat in category_order:
            hours = category_hours.get(cat, 0.0)
            percent = (hours / total_hours * 100.0) if total_hours > 0 else 0.0
            categories.append({"category": cat, "hours": hours, "percent": percent})

        summaries.append({
            "pbi_id": pbi_id,
            "pbi_title": pbi["pbi_title"],
            "pbi_story_points": pbi["pbi_story_points"],
            "pbi_url": pbi["pbi_url"],
            "total_hours": total_hours,
            "missing_hours": missing_hours,
            "has_children": bool(tasks),
            "categories": categories,
        })
    return summaries


# ---------------------------------------------------------------------------
# Writing the Summary sheet
# ---------------------------------------------------------------------------

def write_summary_sheet(workbook, pbi_summaries):
    ws = workbook["Summary"]

    for col_idx, header in enumerate(SUMMARY_HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(wrap_text=True)

    widths = [10, 40, 16, 16, 22, 16, 12]
    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width

    # Helper columns are written but hidden -- they exist only so the
    # charts below can reference a nonzero-only, contiguous range.
    ws.column_dimensions[get_column_letter(CHART_HELPER_CATEGORY_COL)].hidden = True
    ws.column_dimensions[get_column_letter(CHART_HELPER_HOURS_COL)].hidden = True

    chart_anchor_col = len(SUMMARY_HEADERS) + 2  # leave one blank column gap
    row_cursor = 2

    for pbi in pbi_summaries:
        start_row = row_cursor
        block_len = len(pbi["categories"]) if pbi["has_children"] else 1

        pbi_id_cell = ws.cell(row=start_row, column=1, value=pbi["pbi_id"])
        if pbi["pbi_url"]:
            pbi_id_cell.hyperlink = pbi["pbi_url"]
            pbi_id_cell.font = Font(color="0563C1", underline="single")

        ws.cell(row=start_row, column=2, value=pbi["pbi_title"])

        sp_cell = ws.cell(row=start_row, column=3, value=pbi["pbi_story_points"])
        if pbi["pbi_story_points"] == NO_STORY_POINT_PLACEHOLDER:
            sp_cell.fill = YELLOW_FILL

        total_value = format_number(pbi["total_hours"])
        if pbi["missing_hours"]:
            total_display = f"{total_value} ({MISSING_HOURS_NOTE})"
        else:
            total_display = total_value
        total_cell = ws.cell(row=start_row, column=4, value=total_display)
        if pbi["missing_hours"]:
            total_cell.fill = YELLOW_FILL

        for col in (1, 2, 3, 4):
            ws.cell(row=start_row, column=col).alignment = Alignment(
                vertical="center", wrap_text=True
            )

        if pbi["has_children"]:
            for offset, cat in enumerate(pbi["categories"]):
                r = start_row + offset
                ws.cell(row=r, column=5, value=cat["category"])
                ws.cell(row=r, column=6, value=format_number(cat["hours"]))
                pct_cell = ws.cell(row=r, column=7, value=round(cat["percent"], 1))
                # Custom format appends a literal "%" suffix for readability
                # without multiplying the stored value by 100 (the cell's
                # numeric value stays e.g. 16.7, not 0.167) -- so existing
                # consumers/tests that read the raw percentage number are
                # unaffected, while the displayed text reads "16.7%".
                pct_cell.number_format = '0.0"%"'

            if block_len > 1:
                for col in (1, 2, 3, 4):
                    ws.merge_cells(
                        start_row=start_row, start_column=col,
                        end_row=start_row + block_len - 1, end_column=col,
                    )

            nonzero_categories = [cat for cat in pbi["categories"] if cat["hours"] > 0]
            if nonzero_categories:
                # Write the nonzero-only helper rows (hidden columns) that
                # the chart below will reference, so zero-hour categories
                # never appear as slices/labels at all.
                for offset, cat in enumerate(nonzero_categories):
                    r = start_row + offset
                    ws.cell(row=r, column=CHART_HELPER_CATEGORY_COL, value=cat["category"])
                    ws.cell(row=r, column=CHART_HELPER_HOURS_COL, value=format_number(cat["hours"]))

                chart = PieChart()
                chart.title = f"PBI {pbi['pbi_id']} - {pbi['pbi_title']}"[:60]
                # The chart's data lives in hidden helper columns (by
                # design -- see CHART_HELPER_*_COL above). openpyxl/Excel
                # charts default to "plot visible cells only", which would
                # otherwise make the chart render completely empty since
                # none of its source cells are visible.
                chart.visible_cells_only = False
                data = Reference(
                    ws, min_col=CHART_HELPER_HOURS_COL, min_row=start_row,
                    max_row=start_row + len(nonzero_categories) - 1,
                )
                cats = Reference(
                    ws, min_col=CHART_HELPER_CATEGORY_COL, min_row=start_row,
                    max_row=start_row + len(nonzero_categories) - 1,
                )
                chart.add_data(data, titles_from_data=False)
                chart.set_categories(cats)

                # Fixed color per category, consistent across every PBI's
                # chart (so "Development" is always the same color, etc.).
                chart.series[0].data_points = [
                    DataPoint(
                        idx=offset,
                        spPr=GraphicalProperties(solidFill=category_color(cat["category"])),
                    )
                    for offset, cat in enumerate(nonzero_categories)
                ]

                # Each slice's data label shows only its percentage (see
                # below) -- without any other indicator, a bare "75%" slice
                # doesn't tell the viewer which category it is. A compact
                # legend (bottom, to minimize how much horizontal/space it
                # eats) lists each nonzero category's name next to its
                # color swatch, resolved by Excel from the same category
                # reference used for the slices -- so it reads real
                # category names (e.g. "Development"), not a generic
                # series name.
                chart.legend = Legend(legendPos="b")
                chart.dataLabels = DataLabelList()
                chart.dataLabels.showCatName = False
                chart.dataLabels.showVal = False
                chart.dataLabels.showPercent = True
                # Explicitly suppress the series name ("Series1") on each
                # slice's data label -- without this, Excel falls back to
                # displaying it even though the attribute's unset default
                # is supposed to be off.
                chart.dataLabels.showSerName = False
                chart.dataLabels.showLegendKey = False
                chart.dataLabels.showBubbleSize = False
                # A bit taller than before to leave room for the bottom
                # legend without squeezing the pie itself.
                chart.height = 8.5
                chart.width = 11
                anchor = f"{get_column_letter(chart_anchor_col)}{start_row}"
                ws.add_chart(chart, anchor)

            row_cursor = start_row + block_len + CHART_SPACER_ROWS
        else:
            row_cursor = start_row + block_len

    ws.freeze_panes = "A2"


# ---------------------------------------------------------------------------
# Output filename
# ---------------------------------------------------------------------------

def build_output_filename(input_filename, timestamp):
    """Insert "_Summary_<timestamp>" before the extension, replacing any
    existing trailing "_<timestamp>" segment from the input filename so the
    result reads e.g. "PBI_Bug_Task_Report_103_Summary_<timestamp>.xlsx"."""
    match = TIMESTAMP_SUFFIX_RE.search(input_filename)
    if match:
        stem = input_filename[: match.start()]
    else:
        stem, _ext = os.path.splitext(input_filename)
    return f"{stem}_Summary_{timestamp}.xlsx"


# ---------------------------------------------------------------------------
# Orchestration (separated from main() so it can be exercised by tests
# without interactive input()).
# ---------------------------------------------------------------------------

def build_summary_workbook(input_path, output_dir=OUTPUT_DIR):
    """Validate, summarize, and write the Summary tab for `input_path`,
    saving the result as a new file in `output_dir`. Returns the output
    path. Never modifies `input_path` on disk."""
    validate_input_filename(input_path)

    workbook = load_workbook(input_path)
    try:
        validate_workbook_structure(workbook)

        violations = find_blocked_assignee_rows(workbook)
        if violations:
            details = "\n".join(
                f"  - Sheet '{v['sheet']}': PBI {v['pbi_id']}, "
                f"Task/Bug {v['task_id']} ({v['task_title']})"
                for v in violations
            )
            raise ValueError(
                f"Found {len(violations)} row(s) assigned to "
                f"'{BLOCKED_ASSIGNEE}'. Please remove these entries from the "
                f"workbook first, then re-run this script:\n{details}"
            )

        all_rows = read_all_tab(workbook)
        pbi_summaries = summarize_pbis(all_rows)
        write_summary_sheet(workbook, pbi_summaries)

        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        output_name = build_output_filename(os.path.basename(input_path), timestamp)
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, output_name)
        workbook.save(output_path)
        return output_path
    finally:
        workbook.close()


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    print("=" * 70)
    print("PBI/Task Summary Report Generator")
    print("=" * 70)

    input_path = input(
        "\nEnter the path to the PBI_Bug_Task_Report_*.xlsx file: "
    ).strip().strip('"').strip("'")

    if not input_path:
        print("[ERROR] No file path entered. Exiting.")
        sys.exit(1)
    if not os.path.isfile(input_path):
        print(f"[ERROR] File not found: {input_path}")
        sys.exit(1)

    try:
        output_path = build_summary_workbook(input_path)
    except ValueError as e:
        print(f"\n[ERROR] {e}")
        sys.exit(1)

    print(f"\nSummary workbook written to {output_path}")


if __name__ == "__main__":
    main()
