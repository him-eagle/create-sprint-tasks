"""Create a flat, one-row-per-PBI effort summary workbook.

The input report is never modified. A new workbook copy is saved with a flat
Summary sheet containing PBI metadata, total completed hours, category effort,
and category percentages. Azure DevOps is accessed only by the separate test
suite, through read-only endpoints.
"""

import datetime
import os
import re
import sys
from urllib.parse import unquote, urlparse

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import pbi_bug_task_report as report

SCRIPT_DIR = r"C:\Utilities\SprintTask"
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "Reports", "Pbi_bug_tasks")
INPUT_FILENAME_PREFIX = "Effort_Level_Analysis_"
REQUIRED_TABS = ["Summary", "All", "FPSO", "Foundation"]
BLOCKED_ASSIGNEE = "Himanshu Pathak"

CATEGORY_ORDER = [
    "Bug Fixing",
    "QA / Testing",
    "Requirement Study",
    "Review",
    "Design / Documentation",
    "Development",
]
BASE_HEADERS = ["PBI ID", "PBI Title", "PBI State", "PBI Story Points", "Total Hours"]
SUMMARY_HEADERS = BASE_HEADERS + CATEGORY_ORDER + CATEGORY_ORDER
NO_STORY_POINT_PLACEHOLDER = "No_Story_Point"
NO_HOURS_PLACEHOLDER = "No_Hours"
MISSING_HOURS_NOTE = "some tasks/bugs missing hours"

# Styles inspected from the locally available approved reference workbook:
# plain/default header and body fills; bold headers; blue PBI hyperlinks;
# opaque yellow on missing Story Points and missing task hours.
REFERENCE_HEADER_FILL = PatternFill(fill_type=None)
REFERENCE_BODY_FILL = PatternFill(fill_type=None)
DARK_BLUE_HEADER_ARGB = "FF002060"
DARK_BLUE_HEADER_FILL = PatternFill(
    start_color=DARK_BLUE_HEADER_ARGB,
    end_color=DARK_BLUE_HEADER_ARGB,
    fill_type="solid",
)
EFFORT_HEADER_ARGB = "FF9C0006"
EFFORT_HEADER_FILL = PatternFill(
    start_color=EFFORT_HEADER_ARGB,
    end_color=EFFORT_HEADER_ARGB,
    fill_type="solid",
)
HEADER_FONT_COLOR = "FFFFFFFF"
HYPERLINK_FONT_COLOR = "000563C1"
VISIBLE_BORDER_COLOR = "FF808080"
VISIBLE_SIDE = Side(style="thin", color=VISIBLE_BORDER_COLOR)
VISIBLE_BORDER = Border(
    left=VISIBLE_SIDE,
    right=VISIBLE_SIDE,
    top=VISIBLE_SIDE,
    bottom=VISIBLE_SIDE,
)
YELLOW_FILL = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")
TIMESTAMP_SUFFIX_RE = re.compile(r"_(\d{8}_\d{6})\.xlsx$", re.IGNORECASE)


def normalize_assignee(value):
    """Normalize case and whitespace for assignee-name comparisons."""
    return "".join(str(value or "").split()).casefold()


def himanshu_email():
    """Load Himanshu's email from environment/localconfig without exposing it."""
    value = os.environ.get("HIMANSHU_EMAIL", "").strip()
    if value:
        return value
    config_path = os.path.join(SCRIPT_DIR, "localconfig")
    try:
        with open(config_path, encoding="utf-8") as config:
            for line in config:
                key, separator, configured = line.partition("=")
                if separator and key.strip() == "HIMANSHU_EMAIL":
                    return configured.strip().strip("\"'")
    except OSError:
        pass
    return ""


def qa_owners():
    """Load the expected QA/Testing owners from environment/localconfig without hardcoding names in source."""
    value = os.environ.get("QA_OWNERS", "").strip()
    if not value:
        config_path = os.path.join(SCRIPT_DIR, "localconfig")
        try:
            with open(config_path, encoding="utf-8") as config:
                for line in config:
                    key, separator, configured = line.partition("=")
                    if separator and key.strip() == "QA_OWNERS":
                        value = configured.strip().strip("\"'")
                        break
        except OSError:
            pass
    return {name.strip() for name in value.split(",") if name.strip()}


def is_himanshu_assignee(value, blocked_name=BLOCKED_ASSIGNEE):
    if isinstance(value, dict):
        identifiers = (value.get("displayName", ""), value.get("uniqueName", ""))
    else:
        identifiers = (value,)
    target_names = {
        normalized
        for normalized in (
            normalize_assignee(blocked_name),
            normalize_assignee(himanshu_email()),
        )
        if normalized
    }
    return any(normalize_assignee(identifier) in target_names for identifier in identifiers)


def validate_input_filename(path):
    name = os.path.basename(path)
    if not name.startswith(INPUT_FILENAME_PREFIX):
        raise ValueError(
            f"Input file name '{name}' does not start with '{INPUT_FILENAME_PREFIX}'. "
            "Please supply a workbook produced by pbi_bug_task_report.py."
        )


def hours_to_number(value):
    """Return numeric hours; missing placeholders/blanks are zero, invalid values fail."""
    if value is None or value == "" or value == NO_HOURS_PLACEHOLDER:
        return 0.0
    try:
        return float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"Invalid Completed Work value {value!r}; expected numeric hours.") from exc


def is_missing_hours(value):
    return value is None or value == "" or value == NO_HOURS_PLACEHOLDER


def format_number(value):
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def _find_header_row(ws):
    """Find a row containing both the PBI ID and Task ID report headers."""
    for row in ws.iter_rows():
        values = [cell.value for cell in row]
        if "PBI ID" in values and "Task ID" in values:
            return list(row)
    raise ValueError(
        f"Could not find a header row with 'PBI ID'/'Task ID' columns in sheet '{ws.title}'."
    )


def _column_map(header_row):
    return {cell.value: cell.column for cell in header_row if cell.value}


def validate_workbook_structure(workbook):
    missing = [tab for tab in REQUIRED_TABS if tab not in workbook.sheetnames]
    if missing:
        raise ValueError(
            f"Workbook is missing required tab(s): {', '.join(missing)}. "
            f"Expected tabs: {', '.join(REQUIRED_TABS)}."
        )
    summary_ws = workbook["Summary"]
    if any(cell.value is not None for row in summary_ws.iter_rows() for cell in row):
        raise ValueError(
            "The 'Summary' tab already contains content. Supply a fresh, unprocessed "
            "PBI_Bug_Task_Report_*.xlsx workbook."
        )


def find_blocked_assignee_rows(workbook, blocked_name=BLOCKED_ASSIGNEE):
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
        for row_idx in range(header_row[0].row + 1, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=cols["Assigned To"]).value
            if not is_himanshu_assignee(value, blocked_name):
                continue
            violations.append({
                "sheet": sheet_name,
                "pbi_id": ws.cell(row_idx, cols["PBI ID"]).value if "PBI ID" in cols else None,
                "task_id": ws.cell(row_idx, cols["Task ID"]).value if "Task ID" in cols else None,
                "task_title": ws.cell(row_idx, cols["Task Title"]).value if "Task Title" in cols else None,
            })
    return violations


def read_all_tab(workbook):
    """Read All-tab rows by header names, retaining PBI hyperlinks."""
    ws = workbook["All"]
    header_row = _find_header_row(ws)
    cols = _column_map(header_row)
    required = {
        "PBI ID", "PBI Title", "PBI Story Points", "Task ID", "Task Title",
        "Completed Work (hrs)", "Task Categorization",
    }
    missing = sorted(required - set(cols))
    if missing:
        raise ValueError(f"All tab is missing required column(s): {', '.join(missing)}")

    rows = []
    for row_idx in range(header_row[0].row + 1, ws.max_row + 1):
        pbi_id = ws.cell(row_idx, cols["PBI ID"]).value
        if pbi_id is None:
            continue
        pbi_cell = ws.cell(row_idx, cols["PBI ID"])
        task_id = ws.cell(row_idx, cols["Task ID"]).value
        rows.append({
            "pbi_id": pbi_id,
            "team": ws.cell(row_idx, cols["Team"]).value if "Team" in cols else None,
            "pbi_title": ws.cell(row_idx, cols["PBI Title"]).value,
            "pbi_state": None,
            "pbi_story_points": ws.cell(row_idx, cols["PBI Story Points"]).value,
            "pbi_url": pbi_cell.hyperlink.target if pbi_cell.hyperlink else None,
            "task_id": task_id,
            "task_title": ws.cell(row_idx, cols["Task Title"]).value,
            "completed_work": ws.cell(row_idx, cols["Completed Work (hrs)"]).value,
            "categorization": ws.cell(row_idx, cols["Task Categorization"]).value,
        })
    return rows


def summarize_pbis(all_rows):
    """Compute one self-consistent effort rollup per unique PBI."""
    pbis = {}
    for row in all_rows:
        pbi_id = row["pbi_id"]
        if pbi_id not in pbis:
            pbis[pbi_id] = {
                "pbi_id": pbi_id,
                "pbi_title": row["pbi_title"],
                "pbi_state": row.get("pbi_state"),
                "pbi_story_points": row["pbi_story_points"],
                "pbi_url": row["pbi_url"],
                "tasks": [],
            }
        elif (pbis[pbi_id]["pbi_title"], pbis[pbi_id]["pbi_state"],
                            pbis[pbi_id]["pbi_story_points"]) != (
                        row["pbi_title"], row.get("pbi_state"), row["pbi_story_points"]
        ):
            raise ValueError(f"PBI {pbi_id} has inconsistent repeated PBI fields in All tab.")
        if row["task_id"] is not None:
            pbis[pbi_id]["tasks"].append(row)

    results = []
    for pbi_id in sorted(pbis):
        pbi = pbis[pbi_id]
        tasks = pbi["tasks"]
        category_hours = {category: 0.0 for category in CATEGORY_ORDER}
        total_hours = 0.0
        missing_hours = False
        for task in tasks:
            category = task["categorization"]
            if category not in category_hours:
                raise ValueError(
                    f"PBI {pbi_id}, Task/Bug {task['task_id']} ({task.get('task_title', '')}): "
                    f"unexpected Task Categorization {category!r}; expected one of "
                    f"{', '.join(CATEGORY_ORDER)}."
                )
            raw_hours = task["completed_work"]
            if is_missing_hours(raw_hours):
                missing_hours = True
            hours = hours_to_number(raw_hours)
            total_hours += hours
            category_hours[category] += hours

        category_total = sum(category_hours.values())
        if abs(category_total - total_hours) > 1e-8:
            raise ValueError(
                f"PBI {pbi_id}: category effort sums to {category_total:g}, but "
                f"Total Hours sums to {total_hours:g}."
            )
        percentages = {
            category: (hours * 100.0 / total_hours if total_hours else 0.0)
            for category, hours in category_hours.items()
        }
        results.append({
            "pbi_id": pbi_id,
            "pbi_title": pbi["pbi_title"],
            "pbi_state": pbi["pbi_state"],
            "pbi_story_points": pbi["pbi_story_points"],
            "pbi_url": pbi["pbi_url"],
            "total_hours": total_hours,
            "missing_hours": missing_hours,
            "has_children": bool(tasks),
            "category_hours": category_hours,
            "category_percentages": percentages,
        })
    return results


def fetch_live_pbi_metadata(all_rows):
    """Fetch current PBI State and assignee using read-only Azure DevOps API calls."""
    token = report.load_pat_token()
    if not token:
        raise ValueError(
            "AZURE_DEVOPS_PAT is unavailable in the environment/localconfig; "
            "PBI State cannot be populated from live Azure DevOps."
        )
    session = report.get_session(token)
    locations = {}
    for row in all_rows:
        pbi_id = row["pbi_id"]
        url = row.get("pbi_url") or ""
        parsed = urlparse(url)
        parts = [unquote(part) for part in parsed.path.strip("/").split("/")]
        if parsed.netloc.lower() != "dev.azure.com" or len(parts) < 5:
            raise ValueError(f"PBI {pbi_id} has no valid Azure DevOps work-item URL.")
        org, project = parts[0], parts[1]
        locations.setdefault((org, project), set()).add(pbi_id)

    states = {}
    for (org, project), ids in locations.items():
        for item in report.get_work_items_full(session, org, project, sorted(ids)):
            fields = item.get("fields", {})
            if fields.get("System.WorkItemType") != "Product Backlog Item":
                raise ValueError(f"Work item {item.get('id')} is not a Product Backlog Item.")
            state = fields.get("System.State")
            if not state:
                raise ValueError(f"Live Azure DevOps returned a blank System.State for PBI {item['id']}.")
            states[item["id"]] = {
                "state": state,
                "assigned_to": fields.get("System.AssignedTo"),
            }

    expected_ids = {row["pbi_id"] for row in all_rows}
    missing_ids = sorted(expected_ids - set(states))
    if missing_ids:
        raise ValueError(f"Live Azure DevOps returned no details for PBI(s): {missing_ids}.")
    return states


def exclude_himanshu_assigned_pbis(all_rows, live_pbis):
    """Drop each PBI and its rows when its live assignee is Himanshu Pathak."""
    excluded_ids = {
        pbi_id for pbi_id, item in live_pbis.items()
        if is_himanshu_assignee(item.get("assigned_to"))
    }
    return [row for row in all_rows if row["pbi_id"] not in excluded_ids], excluded_ids


def write_summary_sheet(workbook, pbi_summaries):
    """Write a flat table with one row per PBI and grouped Effort/Percentage headers."""
    ws = workbook["Summary"]
    ws.merge_cells("F1:K1")
    ws.merge_cells("L1:Q1")
    for row in range(1, 3):
        for column in range(6, 12):
            cell = ws.cell(row=row, column=column)
            cell.fill = EFFORT_HEADER_FILL
            cell.font = Font(bold=True, color=HEADER_FONT_COLOR)
        for column in range(12, 18):
            cell = ws.cell(row=row, column=column)
            cell.fill = DARK_BLUE_HEADER_FILL
            cell.font = Font(bold=True, color=HEADER_FONT_COLOR)
    for row in (1, 2):
        for column in range(1, 6):
            cell = ws.cell(row=row, column=column)
            cell.fill = DARK_BLUE_HEADER_FILL
            cell.font = Font(bold=True, color=HEADER_FONT_COLOR)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.border = VISIBLE_BORDER

    for address, label in (("F1", "Effort"), ("L1", "Percentage")):
        cell = ws[address]
        cell.value = label
        if address == "F1":
            cell.font = Font(bold=True, color=HEADER_FONT_COLOR)
            cell.fill = EFFORT_HEADER_FILL
        else:
            cell.font = Font(bold=True, color=HEADER_FONT_COLOR)
            cell.fill = DARK_BLUE_HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for start_col, end_col in ((6, 11), (12, 17)):
        for column in range(start_col, end_col + 1):
            ws.cell(row=1, column=column).border = Border(
                left=VISIBLE_SIDE if column == start_col else Side(style=None),
                right=VISIBLE_SIDE if column == end_col else Side(style=None),
                top=VISIBLE_SIDE,
                bottom=VISIBLE_SIDE,
            )

    for column, header in enumerate(SUMMARY_HEADERS, start=1):
        cell = ws.cell(row=2, column=column, value=header)
        if column <= 5:
            cell.font = Font(bold=True, color=HEADER_FONT_COLOR)
            cell.fill = DARK_BLUE_HEADER_FILL
        elif column <= 11:
            cell.font = Font(bold=True, color=HEADER_FONT_COLOR)
            cell.fill = EFFORT_HEADER_FILL
        elif column <= 17:
            cell.font = Font(bold=True, color=HEADER_FONT_COLOR)
            cell.fill = DARK_BLUE_HEADER_FILL
        else:
            cell.font = Font(bold=True)
            cell.fill = REFERENCE_HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = VISIBLE_BORDER

    widths = [12, 48, 18, 18, 38] + [18] * 12
    for column, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(column)].width = width

    for row_idx, pbi in enumerate(pbi_summaries, start=3):
        total_display = format_number(pbi["total_hours"])
        if pbi["missing_hours"]:
            total_display = f"{total_display} ({MISSING_HOURS_NOTE})"
        values = [
            pbi["pbi_id"], pbi["pbi_title"], pbi["pbi_state"],
            pbi["pbi_story_points"], total_display,
            *(format_number(pbi["category_hours"][cat]) for cat in CATEGORY_ORDER),
            *(pbi["category_percentages"][cat] for cat in CATEGORY_ORDER),
        ]
        for column, value in enumerate(values, start=1):
            cell = ws.cell(row=row_idx, column=column, value=value)
            cell.fill = REFERENCE_BODY_FILL
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.border = VISIBLE_BORDER
        pbi_cell = ws.cell(row=row_idx, column=1)
        if pbi["pbi_url"]:
            pbi_cell.hyperlink = pbi["pbi_url"]
            pbi_cell.font = Font(color=HYPERLINK_FONT_COLOR, underline="single")
        if pbi["pbi_story_points"] == NO_STORY_POINT_PLACEHOLDER:
            ws.cell(row_idx, 4).fill = YELLOW_FILL
        if pbi["missing_hours"]:
            ws.cell(row_idx, 5).fill = YELLOW_FILL
        for column in range(12, 18):
            ws.cell(row_idx, column).number_format = '0.0"%"'

    ws.freeze_panes = "F3"
    ws.auto_filter.ref = f"A2:Q{max(2, ws.max_row)}"


def build_output_filename(input_filename, timestamp):
    """Insert _Summary_<timestamp> before .xlsx, replacing an existing timestamp."""
    match = TIMESTAMP_SUFFIX_RE.search(input_filename)
    if match:
        stem = input_filename[:match.start()]
    else:
        stem, _ext = os.path.splitext(input_filename)
    return f"{stem}_Summary_{timestamp}.xlsx"


def build_summary_workbook(input_path, output_dir=OUTPUT_DIR):
    """Build a new output workbook, never modifying the source workbook."""
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
                f"Found {len(violations)} row(s) assigned to '{BLOCKED_ASSIGNEE}'. "
                f"Please remove these entries first:\n{details}"
            )
        all_rows = read_all_tab(workbook)
        live_pbis = fetch_live_pbi_metadata(all_rows)
        included_rows, _excluded_pbi_ids = exclude_himanshu_assigned_pbis(all_rows, live_pbis)
        for row in included_rows:
            row["pbi_state"] = live_pbis[row["pbi_id"]]["state"]
        pbi_summaries = summarize_pbis(included_rows)
        write_summary_sheet(workbook, pbi_summaries)
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        output_name = build_output_filename(os.path.basename(input_path), timestamp)
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, output_name)
        workbook.save(output_path)
        return output_path
    finally:
        workbook.close()


def main():
    print("=" * 70)
    print("PBI/Task Summary Report Generator")
    print("=" * 70)
    input_path = input("\nEnter the path to the PBI_Bug_Task_Report_*.xlsx file: ").strip().strip('"').strip("'")
    if not input_path:
        print("[ERROR] No file path entered. Exiting.")
        sys.exit(1)
    if not os.path.isfile(input_path):
        print(f"[ERROR] File not found: {input_path}")
        sys.exit(1)
    try:
        output_path = build_summary_workbook(input_path)
    except ValueError as error:
        print(f"\n[ERROR] {error}")
        sys.exit(1)
    print(f"\nSummary workbook written to {output_path}")


if __name__ == "__main__":
    main()
