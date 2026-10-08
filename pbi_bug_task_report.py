"""
PBI Collection & Bug/Task Association Report Generator
========================================================

For two teams (FPSO and Foundation), this script:
  1. Runs the team's supplied Azure DevOps saved query and keeps only the
     Product Backlog Item (PBI) work items it returns -- this is the PBI set
     for that team.
  2. For every included PBI, follows its parent/child hierarchy links to pull
     in ALL associated Task/Bug work items, regardless of whether those
     Task/Bug items were themselves returned by the supplied query.
  3. Writes a single Excel workbook with a "Summary" tab (left empty), a
     combined "All" tab, and one tab per team, including PBI Story Points and
     PBI Done Date columns alongside the existing per-task fields.

IMPORTANT: This script is READ-ONLY. It only issues HTTP GET requests
(saved query execution + work item detail fetches) against the Azure DevOps
REST API. It never creates, updates, or deletes anything in Azure DevOps.

QUERY SCOPE: The supplied query for a team determines only which PBIs are
included in that team's report. Task/Bug children are pulled in via each
included PBI's hierarchy links, independent of query membership. A Task/Bug
that is not a hierarchy child of any included PBI is skipped.

Usage:
    python pbi_bug_task_report.py

You will be prompted at the command prompt for:
    0. Your Azure DevOps Personal Access Token (PAT), unless the
       AZURE_DEVOPS_PAT environment variable is already set. Interactive PAT
       input is plain visible text so you can verify what was entered. It is
       never logged or written to any file.
    1. The Sprint name or number (used in the report file name)
    2. The Azure DevOps Query URL for FPSO
    3. The Azure DevOps Query URL for Foundation
"""

import re
import os
import sys
import html
import datetime
from urllib.parse import urlparse, unquote, quote

import requests
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------------------
SCRIPT_DIR = r"C:\Utilities\SprintTask"
OUTPUT_DIR = SCRIPT_DIR + r"\Reports\Pbi_bug_tasks"

API_VERSION = "7.1"

# Work item types that are treated as "PBI" for grouping purposes.
PBI_TYPES = {"Product Backlog Item"}
# Work item types that are treated as child Task/Bug rows.
CHILD_TYPES = {"Task", "Bug"}

STORY_POINTS_FIELD = "Microsoft.VSTS.Scheduling.StoryPoints"
CLOSED_DATE_FIELD = "Microsoft.VSTS.Common.ClosedDate"
ORIGINAL_ESTIMATE_FIELD = "Microsoft.VSTS.Scheduling.OriginalEstimate"
COMPLETED_WORK_FIELD = "Microsoft.VSTS.Scheduling.CompletedWork"

NO_STORY_POINT_PLACEHOLDER = "No_Story_Point"
NO_HOURS_PLACEHOLDER = "No_Hours"

# ---------------------------------------------------------------------------
# Categorization keyword lists (rule-based). Edit these lists to tune
# categorization behavior.
# ---------------------------------------------------------------------------
QA_KEYWORDS = ["test", "qa", "validation", "verify", "regression"]
REQUIREMENT_KEYWORDS = ["requirement", "analysis", "study", "research", "spec", "investigat"]
REVIEW_KEYWORDS = ["review", "demo"]
DESIGN_KEYWORDS = ["design", "document"]


def categorize_task(work_item_type, title, description):
    """Rule-based categorization of a Task/Bug work item.

    Priority order:
      1. Bug type            -> "Bug Fixing"
      2. QA/testing keywords -> "QA / Testing"
      3. Requirement keywords-> "Requirement Study"
      4. Review/demo keywords-> "Review"
      5. Design keywords     -> "Design / Documentation"
      6. Otherwise           -> "Development"
    """
    if (work_item_type or "").strip().lower() == "bug":
        return "Bug Fixing"

    text = f"{title or ''} {description or ''}".lower()

    if any(kw in text for kw in QA_KEYWORDS):
        return "QA / Testing"
    if any(kw in text for kw in REQUIREMENT_KEYWORDS):
        return "Requirement Study"
    if any(kw in text for kw in REVIEW_KEYWORDS):
        return "Review"
    if any(kw in text for kw in DESIGN_KEYWORDS):
        return "Design / Documentation"
    return "Development"


# ---------------------------------------------------------------------------
# HTTP helpers
# ---------------------------------------------------------------------------

def get_session(pat_token):
    session = requests.Session()
    session.auth = ("", pat_token)
    session.headers.update({"Content-Type": "application/json"})
    return session


def load_pat_token():
    """Load the PAT from the environment, then from local ignored config."""
    pat_token = os.environ.get("AZURE_DEVOPS_PAT", "").strip()
    if pat_token:
        return pat_token

    env_path = os.path.join(SCRIPT_DIR, "localconfig")
    try:
        with open(env_path, encoding="utf-8") as env_file:
            for line in env_file:
                key, separator, value = line.partition("=")
                if separator and key.strip() == "AZURE_DEVOPS_PAT":
                    pat_token = value.strip()
                    if len(pat_token) >= 2 and pat_token[0] == pat_token[-1]:
                        if pat_token[0] in ('"', "'"):
                            pat_token = pat_token[1:-1].strip()
                    return pat_token
    except OSError:
        pass

    return ""


def check_response(resp, context):
    if not resp.ok:
        print(f"\n[ERROR] Azure DevOps API call failed while {context}.")
        print(f"        Status: {resp.status_code}")
        try:
            print(f"        Body: {resp.text[:500]}")
        except Exception:
            pass
        resp.raise_for_status()


def parse_query_url(url):
    """Parse an Azure DevOps query URL of the form:
    https://dev.azure.com/{org}/{project}/_queries/query/?tempQueryId={guid}
    or
    https://dev.azure.com/{org}/{project}/_queries/query/{guid}
    """
    parsed = urlparse(url)
    segments = [unquote(s) for s in parsed.path.split("/") if s]

    if "_queries" not in segments:
        raise ValueError("URL does not look like a query URL (missing '_queries').")

    org = segments[0]
    project = segments[1]

    query_id = None
    # try query string first: tempQueryId=xxxx
    m = re.search(r"tempQueryId=([0-9a-fA-F-]{36})", url)
    if m:
        query_id = m.group(1)
    else:
        # try last path segment being a GUID
        for seg in reversed(segments):
            if re.match(r"^[0-9a-fA-F-]{36}$", seg):
                query_id = seg
                break

    if not query_id:
        raise ValueError("Could not find a query id (tempQueryId or GUID) in the URL.")

    return {"org": org, "project": project, "query_id": query_id}


# ---------------------------------------------------------------------------
# Azure DevOps REST API calls (all read-only: GET / query POST)
# ---------------------------------------------------------------------------

def run_wiql_query_by_id(session, org, project, query_id):
    url = (
        f"https://dev.azure.com/{quote(org)}/{quote(project)}/_apis/wit/wiql/"
        f"{query_id}?api-version={API_VERSION}"
    )
    resp = session.get(url)
    check_response(resp, f"running saved/temp query '{query_id}'")
    return parse_wiql_response(resp.json())


def parse_wiql_response(data):
    """Handle both flat ('workItems') and tree/one-hop ('workItemRelations') WIQL shapes."""
    ids = []
    if "workItems" in data:
        ids = [wi["id"] for wi in data["workItems"]]
    elif "workItemRelations" in data:
        seen = set()
        for rel in data["workItemRelations"]:
            for key in ("source", "target"):
                item = rel.get(key)
                if item and item.get("id") and item["id"] not in seen:
                    seen.add(item["id"])
                    ids.append(item["id"])
    return ids


def get_work_items_full(session, org, project, ids):
    """Batch-fetch full work item details (fields + relations) for a list of ids."""
    all_items = []
    chunk_size = 200
    for i in range(0, len(ids), chunk_size):
        chunk = ids[i:i + chunk_size]
        if not chunk:
            continue
        ids_param = ",".join(str(x) for x in chunk)
        url = (
            f"https://dev.azure.com/{quote(org)}/{quote(project)}/_apis/wit/workitems"
            f"?ids={ids_param}&$expand=all&api-version={API_VERSION}"
        )
        resp = session.get(url)
        check_response(resp, f"fetching work item details ({len(chunk)} ids)")
        data = resp.json()
        all_items.extend(data.get("value", []))
    return all_items


# ---------------------------------------------------------------------------
# Work item field helpers
# ---------------------------------------------------------------------------

def strip_html(text):
    if not text:
        return ""
    text = re.sub(r"<[^<]+?>", " ", text)
    text = html.unescape(text)
    return re.sub(r"\s+", " ", text).strip()


def get_assigned_to(fields):
    val = fields.get("System.AssignedTo")
    if isinstance(val, dict):
        return val.get("displayName", "")
    if isinstance(val, str):
        return val
    return ""


def get_parent_id(item):
    fields = item.get("fields", {})
    parent = fields.get("System.Parent")
    if parent:
        return int(parent)
    for rel in item.get("relations", []) or []:
        if rel.get("rel") == "System.LinkTypes.Hierarchy-Reverse":
            m = re.search(r"/(\d+)$", rel.get("url", ""))
            if m:
                return int(m.group(1))
    return None


def collect_child_ids(pbi_items):
    """Gather unique child work item IDs declared via each PBI's forward
    hierarchy relations (System.LinkTypes.Hierarchy-Forward), in the order
    first encountered."""
    ids = []
    seen = set()
    for item in pbi_items:
        for rel in item.get("relations", []) or []:
            if rel.get("rel") == "System.LinkTypes.Hierarchy-Forward":
                m = re.search(r"/(\d+)$", rel.get("url", ""))
                if m:
                    cid = int(m.group(1))
                    if cid not in seen:
                        seen.add(cid)
                        ids.append(cid)
    return ids


def filter_types(items, allowed_types):
    """Return only the items whose System.WorkItemType is in allowed_types."""
    return [
        item for item in items
        if item.get("fields", {}).get("System.WorkItemType", "") in allowed_types
    ]


def build_work_item_url(org, project, item_id):
    return f"https://dev.azure.com/{quote(org)}/{quote(project)}/_workitems/edit/{item_id}"


def format_story_points(value):
    """Return the Story Points value as-is, or the No_Story_Point placeholder
    when the field is genuinely empty/unset. An explicit 0 is real data."""
    if value is None or value == "":
        return NO_STORY_POINT_PLACEHOLDER
    return value


def format_hours(value):
    """Return an hours field (Original Estimate / Completed Work) as-is, or
    the No_Hours placeholder when the field is genuinely empty/unset. An
    explicit 0 is real data."""
    if value is None or value == "":
        return NO_HOURS_PLACEHOLDER
    return value


def sanitize_filename_component(text):
    """Replace characters invalid in Windows filenames with underscores."""
    cleaned = re.sub(r'[\\/:*?"<>|]', "_", (text or "").strip())
    return cleaned or "UnknownSprint"


# ---------------------------------------------------------------------------
# Row building
# ---------------------------------------------------------------------------

def build_rows(pbi_items, child_items, org, project, team_label):
    """Given the included PBI items (Step A) and the fetched Task/Bug child
    items associated with them (Step B), build the report rows: one row per
    Task/Bug under its parent PBI, and one row (with blank task columns) for
    PBIs with no Task/Bug children.

    A child whose parent is not one of the given PBIs is skipped (counted as
    an orphan) -- this can happen if a child's own parent field points
    somewhere unexpected relative to the PBI's declared forward relations.

    Returns (rows, stats), where stats describes included/skipped items.
    """
    pbis = {item["id"]: item for item in pbi_items}

    children_by_parent = {}
    orphan_count = 0
    for child in child_items:
        parent_id = get_parent_id(child)
        if parent_id is not None and parent_id in pbis:
            children_by_parent.setdefault(parent_id, []).append(child)
        else:
            orphan_count += 1

    if orphan_count:
        print(f"        Note: skipped {orphan_count} Task/Bug item(s) not linked to an included PBI.")

    rows = []
    linked_child_count = sum(len(kids) for kids in children_by_parent.values())
    pbi_without_children = len(pbis) - len(children_by_parent)

    for pbi_id in sorted(pbis.keys()):
        pbi = pbis[pbi_id]
        pbi_fields = pbi.get("fields", {})
        pbi_title = pbi_fields.get("System.Title", "")
        pbi_url = build_work_item_url(org, project, pbi_id)
        pbi_story_points = format_story_points(pbi_fields.get(STORY_POINTS_FIELD))
        pbi_closed_date = pbi_fields.get(CLOSED_DATE_FIELD, "")
        pbi_done_date = pbi_closed_date[:10] if pbi_closed_date else ""

        kids = children_by_parent.get(pbi_id, [])
        if not kids:
            rows.append({
                "pbi_id": pbi_id,
                "pbi_url": pbi_url,
                "pbi_title": pbi_title,
                "pbi_story_points": pbi_story_points,
                "pbi_done_date": pbi_done_date,
                "task_id": None,
                "task_url": None,
                "task_type": "",
                "task_title": "",
                "assigned_to": "",
                "state": "",
                "original_estimate": "",
                "completed_work": "",
                "categorization": "",
                "done_date": "",
                "description": "",
                "team": team_label,
            })
            continue

        for child in sorted(kids, key=lambda c: c["id"]):
            cf = child.get("fields", {})
            wi_type = cf.get("System.WorkItemType", "")
            title = cf.get("System.Title", "")
            description = strip_html(cf.get("System.Description", ""))
            closed_date = cf.get(CLOSED_DATE_FIELD, "")
            rows.append({
                "pbi_id": pbi_id,
                "pbi_url": pbi_url,
                "pbi_title": pbi_title,
                "pbi_story_points": pbi_story_points,
                "pbi_done_date": pbi_done_date,
                "task_id": child["id"],
                "task_url": build_work_item_url(org, project, child["id"]),
                "task_type": wi_type,
                "task_title": title,
                "assigned_to": get_assigned_to(cf),
                "state": cf.get("System.State", ""),
                "original_estimate": format_hours(cf.get(ORIGINAL_ESTIMATE_FIELD)),
                "completed_work": format_hours(cf.get(COMPLETED_WORK_FIELD)),
                "categorization": categorize_task(wi_type, title, description),
                "done_date": closed_date[:10] if closed_date else "",
                "description": description,
                "team": team_label,
            })

    return rows, {
        "pbi_count": len(pbis),
        "linked_child_count": linked_child_count,
        "orphan_count": orphan_count,
        "pbi_without_children": pbi_without_children,
    }


def validate_report_scope(rows, query_ids, team_label):
    """Ensure every PBI ID written to a team tab came from that team's
    supplied query. Task/Bug IDs are intentionally NOT checked against the
    query -- they only need to be a hierarchy child of an included PBI
    (enforced during build_rows), which is the core behavior of this script.
    """
    query_id_set = set(query_ids)
    pbi_ids_in_report = {
        row["pbi_id"] for row in rows if row.get("pbi_id") is not None
    }
    outside_query = sorted(pbi_ids_in_report - query_id_set)
    if outside_query:
        raise ValueError(
            f"{team_label} report contains PBI IDs outside the supplied query: "
            f"{outside_query}"
        )


# ---------------------------------------------------------------------------
# Fetch pipeline per team
# ---------------------------------------------------------------------------

def fetch_query_data(session, url, team_label):
    print(f"\n[{team_label}] Parsing query URL...")
    info = parse_query_url(url)
    org, project, query_id = info["org"], info["project"], info["query_id"]
    print(f"        org={org} project={project} query_id={query_id}")

    print(f"[{team_label}] Running query...")
    try:
        ids = run_wiql_query_by_id(session, org, project, query_id)
    except requests.HTTPError:
        print(f"\n[ERROR] Could not execute the {team_label} query by id '{query_id}'.")
        print("        This can happen if the query is a temporary/unsaved query")
        print("        (tempQueryId) that has expired. Please open the query in")
        print("        Azure DevOps, use 'Save As' to save it permanently, and")
        print("        re-run this script with the saved query's URL instead.")
        raise
    print(f"        found {len(ids)} work item(s)")

    if not ids:
        return [], org, project, 0

    print(f"[{team_label}] Fetching PBI details (Step A)...")
    full_items = get_work_items_full(session, org, project, ids)
    pbi_items = filter_types(full_items, PBI_TYPES)
    print(f"        {len(pbi_items)} PBI(s) found in query results")

    print(f"[{team_label}] Fetching associated Task/Bug children (Step B)...")
    child_ids = collect_child_ids(pbi_items)
    fetched_by_id = {item["id"]: item for item in full_items}
    missing_ids = [cid for cid in child_ids if cid not in fetched_by_id]
    if missing_ids:
        fetched_children = get_work_items_full(session, org, project, missing_ids)
        for item in fetched_children:
            fetched_by_id[item["id"]] = item
    candidate_children = [fetched_by_id[cid] for cid in child_ids if cid in fetched_by_id]
    child_items = filter_types(candidate_children, CHILD_TYPES)
    print(f"        {len(child_items)} Task/Bug child item(s) fetched")

    rows, stats = build_rows(pbi_items, child_items, org, project, team_label)
    validate_report_scope(rows, ids, team_label)
    print(
        f"        {stats['pbi_count']} PBI(s), "
        f"{stats['linked_child_count']} linked Task/Bug item(s), "
        f"{stats['orphan_count']} orphan Task/Bug item(s) skipped, "
        f"{stats['pbi_without_children']} PBI-only row(s), "
        f"{len(rows)} report row(s)"
    )
    return rows, org, project, stats["pbi_count"]


# ---------------------------------------------------------------------------
# Excel output
# ---------------------------------------------------------------------------

COLUMNS = [
    ("team", "Team"),
    ("pbi_id", "PBI ID"),
    ("pbi_title", "PBI Title"),
    ("pbi_story_points", "PBI Story Points"),
    ("pbi_done_date", "PBI Done Date"),
    ("task_id", "Task ID"),
    ("task_type", "Task Type"),
    ("task_title", "Task Title"),
    ("assigned_to", "Assigned To"),
    ("state", "State"),
    ("original_estimate", "Original Estimate (hrs)"),
    ("completed_work", "Completed Work (hrs)"),
    ("categorization", "Task Categorization"),
    ("done_date", "Done Date"),
    ("description", "Description"),
]


def write_sheet(wb, sheet_name, rows, source_url):
    ws = wb.create_sheet(title=sheet_name[:31])

    # Row 1: source URL
    ws.cell(row=1, column=1, value="Source URL:")
    url_cell = ws.cell(row=1, column=2, value=source_url)
    url_cell.hyperlink = source_url
    url_cell.font = Font(color="0563C1", underline="single")

    header_row = 2

    # Header row
    for col_idx, (_, header) in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=header_row, column=col_idx, value=header)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(wrap_text=True)

    data_start = header_row + 1
    for r, row in enumerate(rows, start=data_start):
        for col_idx, (key, _) in enumerate(COLUMNS, start=1):
            value = row.get(key)
            cell = ws.cell(row=r, column=col_idx, value=value)
            if key == "pbi_id" and row.get("pbi_url"):
                cell.hyperlink = row["pbi_url"]
                cell.font = Font(color="0563C1", underline="single")
            elif key == "task_id" and row.get("task_url"):
                cell.hyperlink = row["task_url"]
                cell.font = Font(color="0563C1", underline="single")

    # Reasonable column widths, keyed by column key so they stay in sync
    # even if COLUMNS is reordered.
    column_widths = {
        "team": 12,
        "pbi_id": 10,
        "pbi_title": 40,
        "pbi_story_points": 16,
        "pbi_done_date": 14,
        "task_id": 10,
        "task_type": 10,
        "task_title": 40,
        "assigned_to": 20,
        "state": 12,
        "original_estimate": 14,
        "completed_work": 14,
        "categorization": 20,
        "done_date": 12,
        "description": 50,
    }
    for i, (key, _) in enumerate(COLUMNS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = column_widths.get(key, 15)

    ws.freeze_panes = ws.cell(row=data_start, column=1)


def write_workbook(team_data, output_path):
    """team_data: list of tuples (team_label, rows, source_url)"""
    wb = Workbook()
    # remove default sheet, we'll add our own in order
    default_sheet = wb.active
    wb.remove(default_sheet)

    # "Summary" tab first -- intentionally left completely empty.
    wb.create_sheet(title="Summary")

    # "All" tab next
    all_rows = []
    for _, rows, _ in team_data:
        all_rows.extend(rows)
    write_sheet(wb, "All", all_rows, "(combined - see individual tabs for source URLs)")

    for team_label, rows, source_url in team_data:
        write_sheet(wb, team_label, rows, source_url)

    wb.save(output_path)


# ---------------------------------------------------------------------------
# Report orchestration (separated from main() so it can be exercised by
# tests without interactive input()).
# ---------------------------------------------------------------------------

def run_report(session, sprint_name, fpso_url, foundation_url, output_dir=OUTPUT_DIR):
    """Fetch FPSO + Foundation data, write the workbook, and return the output path."""
    team_data = []

    try:
        rows, _, _, _ = fetch_query_data(session, fpso_url, "FPSO")
        team_data.append(("FPSO", rows, fpso_url))
    except Exception as e:
        print(f"[ERROR] Failed to fetch FPSO data: {e}")

    try:
        rows, _, _, _ = fetch_query_data(session, foundation_url, "Foundation")
        team_data.append(("Foundation", rows, foundation_url))
    except Exception as e:
        print(f"[ERROR] Failed to fetch Foundation data: {e}")

    if not team_data:
        raise RuntimeError("No data was fetched for any team. No report was created.")

    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    os.makedirs(output_dir, exist_ok=True)
    sprint_component = sanitize_filename_component(sprint_name)
    output_path = os.path.join(
        output_dir, f"Effort_Level_Analysis_{sprint_component}_{timestamp}.xlsx"
    )

    print(f"\nWriting Excel report to {output_path} ...")
    write_workbook(team_data, output_path)
    print("Done.")
    return output_path


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    print("=" * 70)
    print("PBI Collection & Bug/Task Association Report Generator (read-only)")
    print("=" * 70)

    pat_token = load_pat_token()
    if pat_token:
        print("\nUsing Azure DevOps PAT from environment or local .env file.")
    else:
        pat_token = input("\nEnter your Azure DevOps Personal Access Token (PAT): ").strip()
        # Allow pasting the PAT wrapped in quotes, e.g. "abc123" or 'abc123'
        if len(pat_token) >= 2 and pat_token[0] == pat_token[-1] and pat_token[0] in ('"', "'"):
            pat_token = pat_token[1:-1].strip()
    if not pat_token:
        print("[ERROR] No PAT entered. Exiting.")
        sys.exit(1)

    session = get_session(pat_token)

    sprint_name = input(
        "\nEnter the Sprint name or number (used in the report file name): "
    ).strip()
    print(
        "        File name format: "
        "PBI_Bug_Task_Report_<Sprint>_<timestamp>.xlsx"
    )

    fpso_url = input("\nEnter the Azure DevOps Query URL for FPSO: ").strip()
    foundation_url = input("Enter the Azure DevOps Query URL for Foundation: ").strip()

    try:
        output_path = run_report(session, sprint_name, fpso_url, foundation_url)
    except RuntimeError as e:
        print(f"\n[ERROR] {e}")
        sys.exit(1)

    print(f"\nReport written to {output_path}")


if __name__ == "__main__":
    main()
