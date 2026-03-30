import pandas as pd
import re
import requests
from io import BytesIO

# ----------------------------------------------------------------------
# Name normalisation (same as your original script)
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
# Billing logic
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
# Stuck keyword detection
# ----------------------------------------------------------------------
STUCK_KW = ["pending", "waiting", "tbc", "confirm", "follow up",
            "chase", "delay", "hold", "not start", "approval"]

def is_stuck(task_name, comments):
    combined = (str(task_name) + " " + str(comments)).lower()
    return any(k in combined for k in STUCK_KW)

# ----------------------------------------------------------------------
# Main processing function
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
        # Assume it's a file path or file-like object
        df = pd.read_excel(source, sheet_name=sheet_name, header=None)

    # Find header row (where first cell is "Ref No:" or similar)
    header_row = None
    for i, row in df.iterrows():
        if str(row[0]).strip() == "Ref No:":
            header_row = i
            break
    if header_row is None:
        raise ValueError("Could not find header row in LIVE PROJECTS sheet.")

    data_df = df.iloc[header_row+1:].copy()
    if data_df.shape[1] < 16:
        raise ValueError("Not enough columns; check Excel structure.")

    # Rename columns for clarity
    col_names = [
        "Ref", "Client", "Region", "col3", "col4", "Task", "Lead", "Support",
        "col8", "Works Done", "Date Completed", "col11", "Invoiced", "Xero Inv",
        "Pay Status", "Comments"
    ]
    if data_df.shape[1] >= 18:
        col_names.extend(["Start Date", "End Date"])
    data_df.columns = col_names[:data_df.shape[1]]

    # Drop rows where Ref is empty
    data_df = data_df[data_df["Ref"].notna() & (data_df["Ref"].astype(str).str.strip() != "")]

    # Fill forward Ref, Client, Region for rows where they are missing
    data_df[["Ref", "Client", "Region"]] = data_df[["Ref", "Client", "Region"]].fillna(method="ffill")

    # Convert to string and strip
    for col in ["Ref", "Client", "Region", "Task", "Lead", "Support", "Works Done", "Invoiced", "Xero Inv", "Pay Status", "Comments"]:
        data_df[col] = data_df[col].astype(str).str.strip()

    # Normalise names
    data_df["Lead"] = data_df["Lead"].apply(norm_name)
    data_df["Support"] = data_df["Support"].apply(norm_name)

    # Parse dates if columns exist
    if "Start Date" in data_df.columns:
        data_df["Start Date"] = pd.to_datetime(data_df["Start Date"], errors="coerce")
    else:
        data_df["Start Date"] = pd.NaT
    if "End Date" in data_df.columns:
        data_df["End Date"] = pd.to_datetime(data_df["End Date"], errors="coerce")
    else:
        data_df["End Date"] = pd.NaT

    # Compute derived fields
    data_df["IsDone"] = data_df["Works Done"].str.upper() == "YES"
    data_df["HasInvoice"] = data_df.apply(lambda row: has_invoice(row["Invoiced"], row["Xero Inv"]), axis=1)
    data_df["IsPaid"] = data_df.apply(lambda row: is_paid(row["Invoiced"], row["Xero Inv"], row["Pay Status"], row["Comments"]), axis=1)
    data_df["LikelyPaid"] = data_df.apply(lambda row: is_likely_paid(row["Invoiced"], row["Pay Status"], row["Comments"], row["Xero Inv"], row["IsPaid"]), axis=1)
    data_df["AwaitingPayment"] = data_df.apply(lambda row: is_awaiting_payment(row["Pay Status"], row["Comments"], row["IsPaid"], row["HasInvoice"]), axis=1)
    data_df["SmartStatus"] = data_df.apply(lambda row: smart_billing_status(row["Invoiced"], row["Xero Inv"], row["Pay Status"], row["Comments"], row["IsDone"], row["LikelyPaid"]), axis=1)
    data_df["IsStuck"] = data_df.apply(lambda row: is_stuck(row["Task"], row["Comments"]) and not row["IsDone"], axis=1)

    # Extract PD progress from task names
    def extract_pd(task):
        m = re.search(r"PD\s+(\d+)\s+of\s+(\d+)", task, re.IGNORECASE) or re.search(r"\(\s*(\d+)\s+of\s+(\d+)\s*\)", task)
        if m:
            return int(m.group(1)), int(m.group(2))
        return None, None
    data_df["PD_done"], data_df["PD_total"] = zip(*data_df["Task"].apply(extract_pd))

    # Project-level aggregation
    projects = {}
    for ref, group in data_df.groupby("Ref"):
        tasks = group.to_dict(orient="records")
        client = group["Client"].iloc[0]
        region = group["Region"].iloc[0]
        primary_lead = group["Lead"].iloc[0]
        pd_total = max([t["PD_total"] for t in tasks if t["PD_total"] is not None], default=0)
        pd_done = max([t["PD_done"] for t in tasks if t["PD_done"] is not None], default=0)
        tasks_with_comments = [t for t in tasks if t["Comments"] and len(t["Comments"]) > 5]
        latest_comment = tasks_with_comments[-1]["Comments"] if tasks_with_comments else ""
        latest_task_name = tasks_with_comments[-1]["Task"] if tasks_with_comments else ""
        total = len(tasks)
        done = sum(t["IsDone"] for t in tasks)
        stuck = sum(t["IsStuck"] for t in tasks)
        not_invoiced = sum(1 for t in tasks if t["IsDone"] and not t["HasInvoice"] and not t["LikelyPaid"])
        awaiting_pay = sum(1 for t in tasks if t["AwaitingPayment"])
        invoiced = sum(1 for t in tasks if t["HasInvoice"])
        likely_paid = sum(1 for t in tasks if t["LikelyPaid"])
        paid = sum(1 for t in tasks if t["IsPaid"])
        pending = [t for t in tasks if not t["IsDone"]]
        active_leads = list({t["Lead"] for t in pending if t["Lead"]})
        active_supports = list({t["Support"] for t in pending if t["Support"]})
        xero_invs = list({t["Xero Inv"] for t in tasks if has_real_xero_invoice(t["Xero Inv"])})
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
            "ref": ref, "client": client, "region": region, "lead": primary_lead,
            "total": total, "done": done, "stuck": stuck,
            "pd_done": pd_done, "pd_total": pd_total, "progress": progress, "status": status,
            "aw_pay": awaiting_pay > 0,
            "active_leads": ", ".join(active_leads), "active_supports": ", ".join(active_supports),
            "not_invoiced": not_invoiced, "awaiting_pay_count": awaiting_pay,
            "invoiced_count": invoiced, "likely_paid_count": likely_paid, "paid_count": paid,
            "latest_update": latest_comment or latest_task_name or "—",
            "latest_pay_status": tasks[-1]["Pay Status"] if tasks else "",
            "latest_xero_inv": tasks[-1]["Xero Inv"] if tasks else "",
            "all_xero_invs": ", ".join(xero_invs[:5]),
            "start_date": start_date, "end_date": end_date,
            "tasks": tasks,
        }

    # Person stats
    person_stats = {}
    for t in data_df.to_dict(orient="records"):
        for role in ["Lead", "Support"]:
            name = t[role]
            if not name:
                continue
            if name not in person_stats:
                person_stats[name] = {
                    "name": name, "lead_projects": set(), "lead_done": 0, "lead_pending": 0,
                    "active_set": [], "support_set": [], "stuck_count": 0, "await_count": 0,
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
                if not t["IsDone"]:
                    proj_status = projects.get(t["Ref"], {}).get("status", "?")
                    if proj_status != "Complete":
                        if not any(item["ref"] == t["Ref"] for item in p["active_set"]):
                            p["active_set"].append({"ref": t["Ref"], "status": proj_status, "role": "Lead"})
                if t["IsStuck"]:
                    p["stuck_count"] += 1
                if t["AwaitingPayment"]:
                    p["await_count"] += 1
            else:  # Support
                if not t["IsDone"]:
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

    # Global KPIs
    all_tasks = data_df.to_dict(orient="records")
    total_projects = len(projects)
    complete_projects = sum(1 for p in projects.values() if p["status"] == "Complete")
    in_progress_projects = sum(1 for p in projects.values() if p["status"] == "In Progress")
    blocked_projects = sum(1 for p in projects.values() if p["status"] == "Blocked")
    not_started_projects = sum(1 for p in projects.values() if p["status"] == "Not Started")
    admin_projects = sum(1 for p in projects.values() if p["status"] == "Admin")
    total_tasks = len(all_tasks)
    done_tasks = sum(1 for t in all_tasks if t["IsDone"])
    stuck_tasks = sum(1 for t in all_tasks if t["IsStuck"])
    awaiting_pay_tasks = sum(1 for t in all_tasks if t["AwaitingPayment"])
    not_invoiced_tasks = sum(1 for t in all_tasks if t["IsDone"] and not t["HasInvoice"] and not t["LikelyPaid"])
    likely_paid_tasks = sum(1 for t in all_tasks if t["LikelyPaid"])
    paid_tasks = sum(1 for t in all_tasks if t["IsPaid"])
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
