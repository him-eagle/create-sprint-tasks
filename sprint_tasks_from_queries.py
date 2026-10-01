"""
Sprint PBI/Task Report Generator
=================================

Pulls Product Backlog Items (PBIs) and their child Task/Bug work items from
Azure DevOps using the supplied saved query for each of three teams and writes
a single Excel workbook with a
combined "All" tab (which includes a "Team" column identifying which team
each row belongs to) plus one tab per team. The workbook is saved into the
"Reports" subfolder (created automatically if missing) alongside this
script, e.g. C:/Utilities/SprintTask/Reports/Sprint_Report_<timestamp>.xlsx.

IMPORTANT: This script is READ-ONLY. It only issues HTTP GET requests and
read-only WIQL query calls against the Azure DevOps REST API. It never
creates, updates, or deletes anything in Azure DevOps.

QUERY SCOPE: Report items are sourced only from the IDs returned by the
user-supplied query for that team. The script does not add work items through
an additional query. A Task/Bug is included only when its parent PBI is also
returned by the same query; unparented/out-of-scope children are reported as
skipped. A runtime scope validation rejects any emitted ID not in the query.

Usage:
    python sprint_tasks_from_queries.py

You will be prompted at the command prompt for:
     0. Your Azure DevOps Personal Access Token (PAT), unless the
         AZURE_DEVOPS_PAT environment variable is already set. Interactive PAT
         input is plain visible text so you can verify what was entered. It is
         never logged or written to any file.
    1. The Azure DevOps Query URL for FPSO
    2. The Azure DevOps Query URL for Foundation
    3. The Azure DevOps Query URL for CSR (Kanban board, WIQL query)
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
OUTPUT_DIR = SCRIPT_DIR + r"\Reports"

API_VERSION = "7.1"

# Work item types that are treated as "PBI" for grouping purposes.
PBI_TYPES = {"Product Backlog Item"}
# Work item types that are treated as child Task/Bug rows.
CHILD_TYPES = {"Task", "Bug"}

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


def build_work_item_url(org, project, item_id):
    return f"https://dev.azure.com/{quote(org)}/{quote(project)}/_workitems/edit/{item_id}"


# ---------------------------------------------------------------------------
# Row building
# ---------------------------------------------------------------------------

def build_rows(items, org, project, team_label):
    """Given a flat list of full work items (mix of PBI/Task/Bug/other types),
    build the report rows: one row per Task/Bug under its parent PBI, and one
    row (with blank task columns) for PBIs with no Task/Bug children.

    Returns (rows, stats), where stats describes included and skipped query items.
    """
    pbis = {}
    children = []

    for item in items:
        wi_type = item.get("fields", {}).get("System.WorkItemType", "")
        if wi_type in PBI_TYPES:
            pbis[item["id"]] = item
        elif wi_type in CHILD_TYPES:
            children.append(item)
        # other types (Feature, Epic, etc.) are intentionally ignored.

    children_by_parent = {}
    orphan_count = 0
    for child in children:
        parent_id = get_parent_id(child)
        if parent_id is not None and parent_id in pbis:
            children_by_parent.setdefault(parent_id, []).append(child)
        else:
            orphan_count += 1

    if orphan_count:
        print(f"        Note: skipped {orphan_count} Task/Bug item(s) not linked to a PBI in this scope.")

    rows = []
    linked_child_count = sum(len(kids) for kids in children_by_parent.values())
    pbi_without_children = len(pbis) - len(children_by_parent)
    for pbi_id in sorted(pbis.keys()):
        pbi = pbis[pbi_id]
        pbi_fields = pbi.get("fields", {})
        pbi_title = pbi_fields.get("System.Title", "")
        pbi_url = build_work_item_url(org, project, pbi_id)

        kids = children_by_parent.get(pbi_id, [])
        if not kids:
            rows.append({
                "pbi_id": pbi_id,
                "pbi_url": pbi_url,
                "pbi_title": pbi_title,
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
            closed_date = cf.get("Microsoft.VSTS.Common.ClosedDate", "")
            rows.append({
                "pbi_id": pbi_id,
                "pbi_url": pbi_url,
                "pbi_title": pbi_title,
                "task_id": child["id"],
                "task_url": build_work_item_url(org, project, child["id"]),
                "task_type": wi_type,
                "task_title": title,
                "assigned_to": get_assigned_to(cf),
                "state": cf.get("System.State", ""),
                "original_estimate": cf.get("Microsoft.VSTS.Scheduling.OriginalEstimate", ""),
                "completed_work": cf.get("Microsoft.VSTS.Scheduling.CompletedWork", ""),
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
    """Ensure every ID written to a team tab came from that team's query."""
    query_id_set = set(query_ids)
    report_ids = {
        value
        for row in rows
        for value in (row.get("pbi_id"), row.get("task_id"))
        if value is not None
    }
    outside_query = sorted(report_ids - query_id_set)
    if outside_query:
        raise ValueError(
            f"{team_label} report contains work item IDs outside the supplied query: "
            f"{outside_query}"
        )


# ---------------------------------------------------------------------------
# Fetch pipelines per team type
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

    print(f"[{team_label}] Fetching full work item details...")
    items = get_work_items_full(session, org, project, ids)

    rows, stats = build_rows(items, org, project, team_label)
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


def write_sheet(wb, sheet_name, rows, source_url, extra_note=None):
    ws = wb.create_sheet(title=sheet_name[:31])

    # Row 1: source URL
    ws.cell(row=1, column=1, value="Source URL:")
    url_cell = ws.cell(row=1, column=2, value=source_url)
    url_cell.hyperlink = source_url
    url_cell.font = Font(color="0563C1", underline="single")

    header_row = 2
    if extra_note:
        ws.cell(row=2, column=1, value=extra_note)
        header_row = 3

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
    """team_data: list of tuples (team_label, rows, source_url, extra_note_or_None)"""
    wb = Workbook()
    # remove default sheet, we'll add our own in order
    default_sheet = wb.active
    wb.remove(default_sheet)

    # "All" tab first
    all_rows = []
    for _, rows, _, _ in team_data:
        all_rows.extend(rows)
    write_sheet(wb, "All", all_rows, "(combined - see individual tabs for source URLs)")

    for team_label, rows, source_url, extra_note in team_data:
        write_sheet(wb, team_label, rows, source_url, extra_note)

    wb.save(output_path)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    print("=" * 70)
    print("Sprint PBI/Task Report Generator (read-only)")
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

    fpso_url = input("\nEnter the Azure DevOps Query URL for FPSO: ").strip()
    foundation_url = input("Enter the Azure DevOps Query URL for Foundation: ").strip()
    csr_url = input("Enter the Azure DevOps Query URL for CSR (Kanban board): ").strip()

    print("\n" + "-" * 70)
    print("NOTE: CSR uses a Kanban board / flat query, not an Agile sprint,")
    print("so it has no built-in sprint start/end dates.")
    print("Please verify the query and reporting period yourself before")
    print("relying on this report:")
    print(f"  CSR Query URL: {csr_url}")
    print("-" * 70)
    csr_start = input("Enter the expected reporting period START date (YYYY-MM-DD): ").strip()
    csr_end = input("Enter the expected reporting period END date (YYYY-MM-DD): ").strip()
    confirm = input("Have you verified this query and date range are correct? (y/n): ").strip().lower()
    if confirm != "y":
        print("Please verify the CSR query URL and date range, then re-run the script.")
        sys.exit(0)

    team_data = []

    # FPSO
    try:
        rows, _, _, _ = fetch_query_data(session, fpso_url, "FPSO")
        team_data.append(("FPSO", rows, fpso_url, None))
    except Exception as e:
        print(f"[ERROR] Failed to fetch FPSO data: {e}")

    # Foundation
    try:
        rows, _, _, _ = fetch_query_data(session, foundation_url, "Foundation")
        team_data.append(("Foundation", rows, foundation_url, None))
    except Exception as e:
        print(f"[ERROR] Failed to fetch Foundation data: {e}")

    # CSR
    try:
        rows, _, _, pbi_count = fetch_query_data(session, csr_url, "CSR")
        if pbi_count == 0:
            print("\n[CSR] No Product Backlog Items found in the query results.")
            print("      Skipping the CSR tab entirely (no empty tab will be created).")
        else:
            note = f"User-confirmed reporting period: {csr_start} to {csr_end}"
            team_data.append(("CSR", rows, csr_url, note))
    except Exception as e:
        print(f"[ERROR] Failed to fetch CSR data: {e}")

    if not team_data:
        print("\n[ERROR] No data was fetched for any team. Exiting without creating a report.")
        sys.exit(1)

    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    output_path = f"{OUTPUT_DIR}\\Sprint_Report_{timestamp}.xlsx"

    print(f"\nWriting Excel report to {output_path} ...")
    write_workbook(team_data, output_path)
    print("Done.")


if __name__ == "__main__":
    main()
