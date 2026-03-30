import streamlit as st
import pandas as pd
import plotly.express as px
from dashboard_processor import process_live_projects
import tempfile
import os
import time

# -------------------------------
# Configuration
# -------------------------------
SHEET_NAME = "LIVE PROJECTS"
REFRESH_INTERVAL = 30  # seconds for auto-refresh (only when using URL)

# Try to get the Excel URL from Streamlit secrets
try:
    EXCEL_URL = st.secrets["EXCEL_URL"]
    use_url = True
except:
    use_url = False
    st.info("No Excel URL secret found. Falling back to manual file upload.")

st.set_page_config(page_title="Project Dashboard", layout="wide")
st.title("📊 Advanced Project Dashboard")

# -------------------------------
# Data loading function (cached)
# -------------------------------
@st.cache_data(ttl=REFRESH_INTERVAL if use_url else None)
def load_data_from_url():
    try:
        return process_live_projects(EXCEL_URL, SHEET_NAME)
    except Exception as e:
        st.error(f"Error loading data from URL: {e}")
        return None

@st.cache_data
def load_data_from_file(file_path):
    try:
        return process_live_projects(file_path, SHEET_NAME)
    except Exception as e:
        st.error(f"Error loading file: {e}")
        return None

# -------------------------------
# Load data based on mode
# -------------------------------
if use_url:
    st.sidebar.success("Auto‑refresh enabled – data is fetched from private URL every 30 seconds.")
    data = load_data_from_url()
    if data is None:
        st.stop()
else:
    uploaded_file = st.sidebar.file_uploader(
        "Upload LIVE PROJECTS.xlsx",
        type=["xlsx"],
        help="Upload the latest Excel file. The dashboard will use this file for all views."
    )
    if not uploaded_file:
        st.info("👈 Please upload your LIVE PROJECTS.xlsx file using the sidebar.")
        st.stop()
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(uploaded_file.getvalue())
        tmp_path = tmp.name
    data = load_data_from_file(tmp_path)
    os.unlink(tmp_path)
    if data is None:
        st.stop()

# Extract data
projects = data["projects"]
person_stats = data["person_stats"]
tasks = data["tasks"]
kpis = data["kpis"]
df_raw = data["dataframe"]

# Sidebar filters
st.sidebar.title("Filters")
all_people = sorted({t.get("Lead", "") for t in tasks if t.get("Lead")} |
                    {t.get("Support", "") for t in tasks if t.get("Support")})
person_filter = st.sidebar.selectbox("Filter by person", ["All"] + all_people)
status_filter = st.sidebar.multiselect("Project status",
                                       ["Complete", "In Progress", "Blocked", "Not Started", "Admin"],
                                       default=["In Progress", "Blocked"])
show_only_pending = st.sidebar.checkbox("Show only pending tasks")

# Apply filters to tasks
filtered_tasks = tasks.copy()
if person_filter != "All":
    filtered_tasks = [t for t in filtered_tasks if t.get("Lead") == person_filter or t.get("Support") == person_filter]
if show_only_pending:
    filtered_tasks = [t for t in filtered_tasks if not t.get("IsDone", False)]

# ----------------------------------------------------------------------
# Helper styling functions (colours)
# ----------------------------------------------------------------------
def status_color(status):
    return {
        "Complete": ("#DCFCE7", "#16A34A"),
        "In Progress": ("#CCFBF1", "#0D9488"),
        "Blocked": ("#FEE2E2", "#DC2626"),
        "Not Started": ("#FEF3C7", "#D97706"),
    }.get(status, ("#F1F5F9", "#475569"))

def availability_color(avail):
    return {
        "Available": ("#DCFCE7", "#16A34A"),
        "Light": ("#CCFBF1", "#0D9488"),
        "Busy": ("#FEF3C7", "#D97706"),
        "Overloaded": ("#FEE2E2", "#DC2626"),
    }.get(avail, ("#F1F5F9", "#475569"))

def billing_status_color(status):
    if status.startswith("✓"):
        return "#DCFCE7", "#16A34A"
    if status.startswith("⏳"):
        return "#FEF3C7", "#D97706"
    if status.startswith("⚠ Not Invoiced"):
        return "#FEE2E2", "#DC2626"
    if status.startswith("⚠ Likely Paid"):
        return "#FFEDD5", "#EA580C"
    if status.startswith("📋"):
        return "#DBEAFE", "#2563EB"
    if status == "In Progress":
        return "#CCFBF1", "#0D9488"
    return None

# ----------------------------------------------------------------------
# Tabs
# ----------------------------------------------------------------------
tab1, tab2, tab3, tab4, tab5 = st.tabs(["📈 Overview", "📋 Projects", "👥 Tasks by Person", "💰 Billing", "📅 Timeline"])

# -------------------- TAB 1: OVERVIEW --------------------
with tab1:
    st.header("Key Metrics")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Projects", kpis["total_projects"])
        st.metric("Active Projects", kpis["in_progress"] + kpis["blocked"])
    with col2:
        st.metric("Tasks Completed", f"{kpis['done_tasks']}/{kpis['total_tasks']}",
                  f"{kpis['done_tasks']/max(1,kpis['total_tasks']):.0%}")
        st.metric("Stuck Tasks", kpis["stuck_tasks"])
    with col3:
        st.metric("Paid Tasks", kpis["paid_tasks"])
        st.metric("Awaiting Payment", kpis["awaiting_pay_tasks"])
    with col4:
        st.metric("Not Invoiced", kpis["not_invoiced_tasks"])
        st.metric("Likely Paid", kpis["likely_paid_tasks"])

    with st.expander("Team Availability", expanded=True):
        if person_stats:
            team_df = pd.DataFrame(person_stats)
            team_df["Availability"] = team_df["availability"]
            team_df["Pending Projects"] = team_df["lead_pending"]
            team_df["Active Projects"] = team_df["active_projects"]
            st.dataframe(team_df[["name", "Availability", "Pending Projects", "Active Projects", "lead_done", "completion_pct"]],
                         use_container_width=True)

# -------------------- TAB 2: PROJECTS --------------------
with tab2:
    st.header("All Projects")
    if projects:
        df_projects = pd.DataFrame(projects)
        df_projects = df_projects[df_projects["status"].isin(status_filter)]
        if "start_date" in df_projects.columns:
            df_projects["start_date"] = df_projects["start_date"].dt.strftime("%Y-%m-%d")
        if "end_date" in df_projects.columns:
            df_projects["end_date"] = df_projects["end_date"].dt.strftime("%Y-%m-%d")
        st.dataframe(df_projects, use_container_width=True)

# -------------------- TAB 3: TASKS BY PERSON --------------------
with tab3:
    st.header("Tasks by Person")
    if person_filter != "All":
        st.subheader(f"Tasks for **{person_filter}**")
    if filtered_tasks:
        df_tasks = pd.DataFrame(filtered_tasks)
        cols = ["ref", "client", "task", "Lead", "Support", "IsDone", "SmartStatus", "IsStuck", "Comments"]
        df_display = df_tasks[[c for c in cols if c in df_tasks.columns]].copy()
        if "IsDone" in df_display:
            df_display["IsDone"] = df_display["IsDone"].apply(lambda x: "✓ YES" if x else "NO")
        st.dataframe(df_display, use_container_width=True)
    else:
        st.info("No tasks match the current filters.")

# -------------------- TAB 4: BILLING --------------------
with tab4:
    st.header("Billing Intelligence")
    unpaid_projects = [p for p in projects if p["not_invoiced"] > 0 or p["awaiting_pay_count"] > 0 or p["likely_paid_count"] > 0]
    if unpaid_projects:
        with st.expander("Clients with unpaid or pending billing", expanded=True):
            df_unpaid = pd.DataFrame(unpaid_projects)
            st.dataframe(df_unpaid[["client", "ref", "not_invoiced", "awaiting_pay_count", "likely_paid_count", "paid_count", "all_xero_invs"]],
                         use_container_width=True)
    else:
        st.success("All clients are fully paid and invoiced.")

    bill_tasks = [t for t in tasks if (t.get("IsDone", False) and not t.get("HasInvoice", False) and not t.get("LikelyPaid", False)) or t.get("AwaitingPayment", False) or t.get("LikelyPaid", False)]
    if bill_tasks:
        with st.expander("Tasks requiring billing action", expanded=True):
            df_bill_tasks = pd.DataFrame(bill_tasks)
            st.dataframe(df_bill_tasks[["ref", "client", "task", "SmartStatus", "Xero Inv", "Pay Status", "Comments"]],
                         use_container_width=True)

# -------------------- TAB 5: TIMELINE --------------------
with tab5:
    st.header("Project Timeline")
    # Gantt chart using plotly.express.timeline
    if projects and any(p.get("start_date") and p.get("end_date") for p in projects):
        gantt_data = []
        for p in projects:
            if p["start_date"] and p["end_date"] and not pd.isna(p["start_date"]) and not pd.isna(p["end_date"]):
                gantt_data.append(dict(
                    Task=p["ref"],
                    Start=p["start_date"],
                    Finish=p["end_date"],
                    Status=p["status"],
                    Client=p["client"]
                ))
        if gantt_data:
            df_gantt = pd.DataFrame(gantt_data)
            fig = px.timeline(df_gantt, x_start="Start", x_end="Finish", y="Task",
                              color="Status", hover_data=["Client"],
                              title="Project Timeline")
            fig.update_yaxes(categoryorder="total ascending")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("No projects have valid start and end dates.")
    else:
        st.warning("Start and end dates are not available in the source data. Please add them to the Excel file (embedded in Client column).")

    if projects and any(p.get("end_date") for p in projects):
        today = pd.Timestamp.today()
        upcoming = [p for p in projects if p.get("end_date") and p["end_date"] >= today and p["status"] != "Complete"]
        if upcoming:
            with st.expander("Upcoming Deadlines", expanded=True):
                upcoming_df = pd.DataFrame(upcoming)
                upcoming_df = upcoming_df[["ref", "client", "end_date", "status"]].sort_values("end_date")
                upcoming_df["end_date"] = upcoming_df["end_date"].dt.strftime("%Y-%m-%d")
                st.dataframe(upcoming_df, use_container_width=True)

# Optional: show last refresh time in sidebar
if use_url:
    st.sidebar.caption(f"Auto-refresh every {REFRESH_INTERVAL} seconds")
    if st.sidebar.button("Force refresh now"):
        st.cache_data.clear()
        st.rerun()
