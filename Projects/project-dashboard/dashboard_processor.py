import pandas as pd
import re
import requests
from io import BytesIO
from datetime import datetime

# ----------------------------------------------------------------------
# 1. Name normalisation (matching your ExcelScript)
# ----------------------------------------------------------------------
CANONICAL = {
    "martin": "Martin", "martin lundy": "Martin",
    "gerard": "Gerard", "gerard de brun": "Gerard",
    "fredy": "Fredy", "samuel": "Samuel", "girish": "Girish",
    "jino": "Jino", "coral": "Coral", "jenni": "Jenni",
    "chris": "Chris", "kimberly": "Kimberly", "claudia": "Claudia",
    "roxana": "Roxana", "daniyal": "Daniyal", "jonathan": "Jonathan",
    "patrick": "Patrick", "marysol": "Marysol", "cris": "Cris",
    "christopher": "Chris", "christopher brenes": "Chris"
}

def norm_name(n):
    if not n or not isinstance(n, str):
        return ""
    t = n.strip().replace(",", "").replace(".", "").split("\n")[0].strip()
    if not t or t in ("x", "?", "-") or t.startswith("http"):
        return ""
    lo = t.lower()
    if lo in CANONICAL:
        return CANONICAL[lo]
    for key in CANONICAL:
        if lo.startswith(key):
            return CANONICAL[key]
    return t

# ----------------------------------------------------------------------
# 2. Billing logic (same as your ExcelScript)
# ----------------------------------------------------------------------
def is_invoiced_col(v):
    if not isinstance(v, str):
        return False
    s = v.strip().upper()
    return s in ("YES", "Y", "100% DISCOUNT", "INCLUDED")

def has_real_xero_invoice(xero_inv):
    inv = str(xero_inv).strip() if xero_inv else ""
    return inv != "" and inv.upper() not in ("NAN", "INCLUDED")

def has_invoice(invoiced, xero_inv):
    if is_invoiced_col(invoiced):
        return True
    if str(xero_inv).strip().upper() == "INCLUDED":
        return True
    return has_real_xero_invoice(xero_inv)

def is_paid(invoiced, xero_inv, pay_status, comments):
    if is_invoiced_col(invoiced):
        return True
    if str(xero_inv).strip().upper() == "INCLUDED":
        return True
    combined = (str(pay_status) + " " + str(comments)).lower()
    return bool(re.search(r"(paid|completed|done|received|settled)", combined)) \
           and not re.search(r"await", combined)

def is_likely_paid(invoiced, pay_status, comments, xero_inv, is_paid_flag):
    if is_paid_flag:
        return False
    if has_invoice(invoiced, xero_inv):
        return False
    comm = str(comments).lower()
    pay = str(pay_status).lower()
    return "paid" in comm and "await" not in pay and "unpaid" not in comm

def is_awaiting_payment(pay_status, comments, is_paid_flag, has_inv_flag):
    if is_paid_flag or not has_inv_flag:
        return False
    combined = (str(pay_status) + " " + str(comments)).lower()
    return "await" in combined

def smart_billing_status(invoiced, xero_inv, pay_status, comments, is_done, likely_paid_flag):
    inv_exists = has_invoice(invoiced, xero_inv)
    paid = is_paid(invoiced, xero_inv, pay_status, comments)
    if is_done and likely_paid_flag:
        return "⚠ Likely Paid – Verify"
    if not inv_exists and is_done:
        return "⚠ Not Invoiced"
    if not inv_exists and not is_done:
        return "In Progress"
    if paid:
        return "✓ Paid"
    if "await" in (str(pay_status) + " " + str(comments)).lower():
        return "⏳ Awaiting Payment"
    return "📋 Invoiced – Unconfirmed"

# ----------------------------------------------------------------------
# 3. Stuck keyword detection
# ----------------------------------------------------------------------
STUCK_KW = ["pending", "waiting", "tbc", "confirm", "follow up",
            "chase", "delay", "hold", "not start", "approval"]

def is_stuck(task_name, comments):
    combined = (str(task_name) + " " + str(comments)).lower()
    return any(k in combined for k in STUCK_KW)

# ----------------------------------------------------------------------
# 4. Extract start/end dates from Client column
# ----------------------------------------------------------------------
def extract_dates_from_client(client_text):
    """Parse strings like 'Start Date :14.01.2025 End Date : 4.02.2025'."""
    start = None
    end = None
    if not isinstance(client_text, str):
        return start, end
    # Look for "Start Date :" or "Start Date:" followed by a date
    start_match = re.search(r'Start\s*Date\s*:?\s*(\d{1,2}\.\d{1,2}\.\d{4})', client_text, re.IGNORECASE)
    if start_match:
        try:
            start = datetime.strptime(start_match.group(1), "%d.%m.%Y")
        except:
            pass
    end_match = re.search(r'End\s*Date\s*:?\s*(\d{1,2}\.\d{1,2}\.\d{4})', client_text, re.IGNORECASE)
    if end_match:
        try:
            end = datetime.strptime(end_match.group(1), "%d.%m.%Y")
        except:
            pass
    return start, end

# ----------------------------------------------------------------------
# 5. Main processing function – dynamic column mapping
# ----------------------------------------------------------------------
def process_live_projects(source, sheet_name="LIVE PROJECTS"):
    """
    source can be:
    - a local file path (string)
    - a URL (string starting with http)
    - a file-like object (e.g., from st.file_uploader)
    """
    if isinstance(source, str) and source.startswith(('http://', 'https://')):
        response = requests.get(source)
        response.raise_for_status()
        excel_data = BytesIO(response.content)
        df = pd.read_excel(excel_data, sheet_name=sheet_name, header=None)
    else:
        df = pd.read_excel(source, sheet_name=sheet_name, header=None)

    # Find header row (first cell = "Ref No:")
    header_row = None
    for i, row in df.iterrows():
        if str(row[0]).strip() == "Ref No:":
            header_row = i
            break
    if header_row is None:
        raise ValueError("Could not find header row with 'Ref No:' in first cell.")

    # Read the header row to get column names
    header = df.iloc[header_row].astype(str).str.strip().tolist()
    # Map common names to our internal fields
    col_map = {}
    for idx, name in enumerate(header):
        name_lower = name.lower()
        if "ref" in name_lower or name_lower == "project ref":
            col_map["Ref"] = idx
        elif "client" in name_lower or "project" in name_lower:
            col_map["Client"] = idx
        elif "region" in name_lower or "roi / uk" in name_lower:
            col_map["Region"] = idx
        elif "task" in name_lower:
            col_map["Task"] = idx
        elif "lead" in name_lower:
            col_map["Lead"] = idx
        elif "support" in name_lower:
            col_map["Support"] = idx
        elif "works completed" in name_lower or "works done" in name_lower:
            col_map["Works Done"] = idx
        elif "date completed" in name_lower:
            col_map["Date Completed"] = idx
        elif "invoiced" in name_lower:
            col_map["Invoiced"] = idx
        elif "xero invoice no" in name_lower or "xero inv" in name_lower:
            col_map["Xero Inv"] = idx
        elif "payment status" in name_lower:
            col_map["Pay Status"] = idx
        elif "comments" in name_lower:
            col_map["Comments"] = idx
        # Also look for "PO Number" and "Xero Quote Number" – we don't use them, but keep them

    # Ensure essential columns exist
    required = ["Ref", "Client", "Task"]
    for col in required:
        if col not in col_map:
            raise ValueError(f"Required column '{col}' not found. Found: {list(col_map.keys())}")

    # Read data rows (after header row)
    data_df = df.iloc[header_row+1:].copy()
    # Build a DataFrame with the mapped columns
    new_data = {}
    for col, idx in col_map.items():
        if idx < data_df.shape[1]:
            new_data[col] = data_df.iloc[:, idx]
    data_df = pd.DataFrame(new_data)

    # Drop rows where Ref is empty
    data_df = data_df[data_df["Ref"].notna() & (data_df["Ref"].astype(str).str.strip() != "")]

    # Fill forward Ref, Client, Region
    for col in ["Ref", "Client", "Region"]:
        if col in data_df.columns:
            data_df[col] = data_df[col].fillna(method="ffill")

    # Convert text columns to string and strip
    text_cols = ["Ref", "Client", "Region", "Task", "Lead", "Support", "Works Done", "Invoiced", "Xero Inv", "Pay Status", "Comments"]
    for col in text_cols:
        if col in data_df.columns:
            data_df[col] = data_df[col].astype(str).str.strip()

    # Extract start/end dates from Client column
    start_dates, end_dates = zip(*data_df["Client"].apply(extract_dates_from_client))
    data_df["Start Date"] = pd.to_datetime(start_dates, errors="coerce")
    data_df["End Date"] = pd.to_datetime(end_dates, errors="coerce")

    # Normalise names if Lead/Support exist
    if "Lead" in data_df.columns:
        data_df["Lead"] = data_df["Lead"].apply(norm_name)
    if "Support" in data_df.columns:
        data_df["Support"] = data_df["Support"].apply(norm_name)

    # Compute derived fields
    if "Works Done" in data_df.columns:
        data_df["IsDone"] = data_df["Works Done"].str.upper() == "YES"
    else:
        data_df["IsDone"] = False

    data_df["HasInvoice"] = data_df.apply(
        lambda row: has_invoice(row.get("Invoiced", ""), row.get("Xero Inv", "")), axis=1)
    data_df["IsPaid"] = data_df.apply(
        lambda row: is_paid(row.get("Invoiced", ""), row.get("Xero Inv", ""),
                            row.get("Pay Status", ""), row.get("Comments", "")), axis=1)
    data_df["LikelyPaid"] = data_df.apply(
        lambda row: is_likely_paid(row.get("Invoiced", ""), row.get("Pay Status", ""),
                                   row.get("Comments", ""), row.get("Xero Inv", ""),
                                   row.get("IsPaid", False)), axis=1)
    data_df["AwaitingPayment"] = data_df.apply(
        lambda row: is_awaiting_payment(row.get("Pay Status", ""), row.get("Comments", ""),
                                        row.get("IsPaid", False), row.get("HasInvoice", False)), axis=1)
    data_df["SmartStatus"] = data_df.apply(
        lambda row: smart_billing_status(row.get("Invoiced", ""), row.get("Xero Inv", ""),
                                         row.get("Pay Status", ""), row.get("Comments", ""),
                                         row.get("IsDone", False), row.get("LikelyPaid", False)), axis=1)
    data_df["IsStuck"] = data_df.apply(
        lambda row: is_stuck(row.get("Task", ""), row.get("Comments", "")) and not row.get("IsDone", False), axis=1)

    # Extract PD progress from task names (optional)
    def extract_pd(task):
        if not isinstance(task, str):
            return None, None
        m = re.search(r"PD\s+(\d+)\s+of\s+(\d+)", task, re.IGNORECASE) or re.search(r"\(\s*(\d+)\s+of\s+(\d+)\s*\)", task)
        if m:
            return int(m.group(1)), int(m.group(2))
        return None, None
    if "Task" in data_df.columns:
        data_df["PD_done"], data_df["PD_total"] = zip(*data_df["Task"].apply(extract_pd))
    else:
        data_df["PD_done"] = None
        data_df["PD_total"] = None

    # ------------------------------------------------------------------
    # Project-level aggregation
    # ------------------------------------------------------------------
    projects = {}
    for ref, group in data_df.groupby("Ref"):
        tasks = group.to_dict(orient="records")
        client = group["Client"].iloc[0] if "Client" in group else ""
        region = group["Region"].iloc[0] if "Region" in group else ""
        primary_lead = group["Lead"].iloc[0] if "Lead" in group else ""
        pd_total = max([t["PD_total"] for t in tasks if t["PD_total"] is not None], default=0)
        pd_done = max([t["PD_done"] for t in tasks if t["PD_done"] is not None], default=0)
        tasks_with_comments = [t for t in tasks if t.get("Comments") and len(str(t["Comments"])) > 5]
        latest_comment = tasks_with_comments[-1]["Comments"] if tasks_with_comments else ""
        latest_task_name = tasks_with_comments[-1]["Task"] if tasks_with_comments else ""
        total = len(tasks)
        done = sum(t.get("IsDone", False) for t in tasks)
        stuck = sum(t.get("IsStuck", False) for t in tasks)
        not_invoiced = sum(1 for t in tasks if t.get("IsDone", False) and not t.get("HasInvoice", False) and not t.get("LikelyPaid", False))
        awaiting_pay = sum(1 for t in tasks if t.get("AwaitingPayment", False))
        invoiced = sum(1 for t in tasks if t.get("HasInvoice", False))
        likely_paid = sum(1 for t in tasks if t.get("LikelyPaid", False))
        paid = sum(1 for t in tasks if t.get("IsPaid", False))
        pending = [t for t in tasks if not t.get("IsDone", False)]
        active_leads = list({t.get("Lead", "") for t in pending if t.get("Lead")})
        active_supports = list({t.get("Support", "") for t in pending if t.get("Support")})
        xero_invs = list({t.get("Xero Inv", "") for t in tasks if has_real_xero_invoice(t.get("Xero Inv", ""))})
        if pd_total > 0:
            progress = min(100, int(round(pd_done / pd_total * 100)))
        elif total > 0:
            progress = int(round(done / total * 100))
        else:
            progress = 0
        if total > 0 and (done > 0 or stuck > 0):
            if progress >= 100:
                status = "Complete"
            elif stuck > 0:
                status = "Blocked"
            elif progress > 0:
                status = "In Progress"
            else:
                status = "Not Started"
        else:
            status = "Admin"
        start_date = group["Start Date"].min() if not group["Start Date"].isna().all() else pd.NaT
        end_date = group["End Date"].max() if not group["End Date"].isna().all() else pd.NaT
        projects[ref] = {
            "ref": ref,
            "client": client,
            "region": region,
            "lead": primary_lead,
            "total": total,
            "done": done,
            "stuck": stuck,
            "pd_done": pd_done,
            "pd_total": pd_total,
            "progress": progress,
            "status": status,
            "aw_pay": awaiting_pay > 0,
            "active_leads": ", ".join(active_leads),
            "active_supports": ", ".join(active_supports),
            "not_invoiced": not_invoiced,
            "awaiting_pay_count": awaiting_pay,
            "invoiced_count": invoiced,
            "likely_paid_count": likely_paid,
            "paid_count": paid,
            "latest_update": latest_comment or latest_task_name or "—",
            "latest_pay_status": tasks[-1].get("Pay Status", "") if tasks else "",
            "latest_xero_inv": tasks[-1].get("Xero Inv", "") if tasks else "",
            "all_xero_invs": ", ".join(xero_invs[:5]),
            "start_date": start_date,
            "end_date": end_date,
            "tasks": tasks,
        }

    # ------------------------------------------------------------------
    # Person stats
    # ------------------------------------------------------------------
    person_stats = {}
    for t in data_df.to_dict(orient="records"):
        for role in ["Lead", "Support"]:
            name = t.get(role, "")
            if not name:
                continue
            if name not in person_stats:
                person_stats[name] = {
                    "name": name,
                    "lead_projects": set(),
                    "lead_done": 0,
                    "lead_pending": 0,
                    "active_set": [],
                    "support_set": [],
                    "stuck_count": 0,
                    "await_count": 0,
                }
            p = person_stats[name]
            if role == "Lead":
                if t["Ref"] not in p["lead_projects"]:
                    p["lead_projects"].add(t["Ref"])
                    proj_status = projects.get(t["Ref"], {}).get("status", "?")
                    if proj_status == "Complete":
                        p["lead_done"] += 1
                    else:
                        p["lead_pending"] += 1
                if not t.get("IsDone", False):
                    proj_status = projects.get(t["Ref"], {}).get("status", "?")
                    if proj_status != "Complete":
                        if not any(item["ref"] == t["Ref"] for item in p["active_set"]):
                            p["active_set"].append({"ref": t["Ref"], "status": proj_status, "role": "Lead"})
                if t.get("IsStuck", False):
                    p["stuck_count"] += 1
                if t.get("AwaitingPayment", False):
                    p["await_count"] += 1
            else:  # Support
                if not t.get("IsDone", False):
                    proj_status = projects.get(t["Ref"], {}).get("status", "?")
                    if proj_status != "Complete":
                        if not any(item["ref"] == t["Ref"] for item in p["active_set"]):
                            p["active_set"].append({"ref": t["Ref"], "status": proj_status, "role": "Support"})
                        if t["Ref"] not in p["support_set"]:
                            p["support_set"].append(t["Ref"])

    for name, p in person_stats.items():
        p["lead_project_count"] = len(p["lead_projects"])
        p["active_projects"] = len(p["active_set"])
        p["support_count"] = len(p["support_set"])
        p["completion_pct"] = round(p["lead_done"] / max(1, p["lead_project_count"]) * 100)
        if p["lead_pending"] == 0:
            p["availability"] = "Available"
        elif p["lead_pending"] < 20:
            p["availability"] = "Light"
        elif p["lead_pending"] < 60:
            p["availability"] = "Busy"
        else:
            p["availability"] = "Overloaded"

    # ------------------------------------------------------------------
    # Global KPIs
    # ------------------------------------------------------------------
    all_tasks = data_df.to_dict(orient="records")
    total_projects = len(projects)
    complete_projects = sum(1 for p in projects.values() if p["status"] == "Complete")
    in_progress_projects = sum(1 for p in projects.values() if p["status"] == "In Progress")
    blocked_projects = sum(1 for p in projects.values() if p["status"] == "Blocked")
    not_started_projects = sum(1 for p in projects.values() if p["status"] == "Not Started")
    admin_projects = sum(1 for p in projects.values() if p["status"] == "Admin")
    total_tasks = len(all_tasks)
    done_tasks = sum(1 for t in all_tasks if t.get("IsDone", False))
    stuck_tasks = sum(1 for t in all_tasks if t.get("IsStuck", False))
    awaiting_pay_tasks = sum(1 for t in all_tasks if t.get("AwaitingPayment", False))
    not_invoiced_tasks = sum(1 for t in all_tasks if t.get("IsDone", False) and not t.get("HasInvoice", False) and not t.get("LikelyPaid", False))
    likely_paid_tasks = sum(1 for t in all_tasks if t.get("LikelyPaid", False))
    paid_tasks = sum(1 for t in all_tasks if t.get("IsPaid", False))
    free_team = sum(1 for p in person_stats.values() if p["lead_pending"] == 0)
    overloaded_team = sum(1 for p in person_stats.values() if p["lead_pending"] >= 60)

    kpis = {
        "total_projects": total_projects,
        "complete": complete_projects,
        "in_progress": in_progress_projects,
        "blocked": blocked_projects,
        "not_started": not_started_projects,
        "admin": admin_projects,
        "total_tasks": total_tasks,
        "done_tasks": done_tasks,
        "stuck_tasks": stuck_tasks,
        "awaiting_pay_tasks": awaiting_pay_tasks,
        "not_invoiced_tasks": not_invoiced_tasks,
        "likely_paid_tasks": likely_paid_tasks,
        "paid_tasks": paid_tasks,
        "free_team": free_team,
        "overloaded_team": overloaded_team,
    }

    return {
        "projects": list(projects.values()),
        "person_stats": list(person_stats.values()),
        "tasks": all_tasks,
        "kpis": kpis,
        "dataframe": data_df
    }
