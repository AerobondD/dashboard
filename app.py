import hmac
from pathlib import Path
from dataclasses import dataclass
import streamlit as st
import pandas as pd
import plotly.graph_objects as go

st.set_page_config(page_title="Aerobond Dashboard", layout="wide", initial_sidebar_state="expanded")

PRIMARY_RED = "#D71920"
CHARCOAL = "#2B2B2B"
BG = "#F7F8FA"
CARD = "#FFFFFF"
GRAY = "#7A7F87"
GREEN = "#2E8B57"
AMBER = "#D99A00"
RED = "#C62828"

BASE_DIR = Path(__file__).resolve().parent

st.markdown(
    f"""
    <style>
    .stApp {{ background: {BG}; color: {CHARCOAL}; }}
    section[data-testid="stSidebar"] {{ background: #111111; }}
    section[data-testid="stSidebar"] * {{ color: white; }}
    .brand-box {{
        padding: 1rem;
        border-radius: 14px;
        background: {CARD};
        border: 1px solid #e8e8e8;
        margin-bottom: 1rem;
    }}
    .kpi-card {{
        background: {CARD};
        border: 1px solid #eaeaea;
        border-radius: 16px;
        padding: 1rem;
        box-shadow: 0 2px 10px rgba(0,0,0,0.04);
        min-height: 118px;
    }}
    .kpi-title {{ font-size: 0.85rem; color: {GRAY}; margin-bottom: 0.35rem; }}
    .kpi-value {{ font-size: 2rem; font-weight: 700; color: {CHARCOAL}; line-height: 1.1; }}
    .kpi-delta {{ font-size: 0.85rem; color: {GRAY}; margin-top: 0.3rem; }}
    .section-card {{
        background: {CARD};
        border: 1px solid #eaeaea;
        border-radius: 16px;
        padding: 1rem;
        box-shadow: 0 2px 10px rgba(0,0,0,0.04);
    }}
    </style>
    """,
    unsafe_allow_html=True,
)

def img_path(name):
    return BASE_DIR / name

def show_image(name, width=None):
    p = img_path(name)
    if p.exists():
        st.image(str(p), use_container_width=True)
    else:
        st.write("Aerobond Dashboard")

def init_state():
    if "auth" not in st.session_state:
        st.session_state.auth = False
    if "snapshots" not in st.session_state:
        st.session_state.snapshots = []
    if "uploaded_names" not in st.session_state:
        st.session_state.uploaded_names = set()

def check_password():
    secret_pw = st.secrets.get("password", "")
    with st.sidebar:
        st.markdown("### Login")
        password = st.text_input("Password", type="password", key="password")
        if st.button("Sign in", use_container_width=True):
            if password == secret_pw and secret_pw != "":
                st.session_state.auth = True
                st.success("Signed in")
            else:
                st.error("Invalid password")
    return st.session_state.auth

def read_workbook(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    sheets = {}
    for name in xls.sheet_names:
        sheets[name] = pd.read_excel(uploaded_file, sheet_name=name)
    return sheets

def safe_first(df, col, default=None):
    if df is None or df.empty or col not in df.columns:
        return default
    vals = df[col].dropna()
    return vals.iloc[0] if len(vals) else default

def parse_snapshot(uploaded_file):
    sheets = read_workbook(uploaded_file)
    md = sheets.get("Meeting Details", pd.DataFrame())
    summary = sheets.get("Summary", pd.DataFrame())
    actions = sheets.get("Actions Items", pd.DataFrame())
    risks = sheets.get("Risk Review", pd.DataFrame())
    trans = sheets.get("Document Transmittal", pd.DataFrame())
    next_steps = sheets.get("Next Steps", pd.DataFrame())
    return {
        "filename": uploaded_file.name,
        "project_name": safe_first(md, "Project Name", "Unknown"),
        "project_no": safe_first(md, "Project No.", ""),
        "report_date": pd.to_datetime(safe_first(md, "Date", pd.Timestamp.today()), errors="coerce"),
        "reporting_staff": safe_first(md, "Reporting Staff", ""),
        "summary": summary,
        "actions": actions,
        "risks": risks,
        "transmittal": trans,
        "next_steps": next_steps,
        "sheets": sheets,
    }

def compute_kpis(snapshot, previous=None):
    actions = snapshot["actions"]
    risks = snapshot["risks"]
    summary = snapshot["summary"]
    report_date = snapshot["report_date"]

    open_actions = 0
    overdue_actions = 0
    due_next_7_days = 0

    if not actions.empty:
        if "Status" in actions.columns:
            status = actions["Status"].astype(str).str.lower()
            open_actions = int((status != "closed").sum())
            if "Due Date" in actions.columns:
                due = pd.to_datetime(actions["Due Date"], errors="coerce")
                overdue_actions = int(((due < report_date) & (status != "closed")).sum())
                due_next_7_days = int(((due >= report_date) & (due <= report_date + pd.Timedelta(days=7))).sum())

    open_risks = 0
    high_risks = 0
    if not risks.empty:
        if "Status" in risks.columns:
            open_risks = int((risks["Status"].astype(str).str.lower() != "closed").sum())
        if "Risk Level" in risks.columns:
            high_risks = int((risks["Risk Level"].astype(str).str.lower() == "high").sum())

    progress_value = None
    if not summary.empty:
        for col in ["Progress ex.gst", "Progress ex GST", "Progress"]:
            if col in summary.columns:
                progress_value = pd.to_numeric(summary[col], errors="coerce").fillna(0).sum()
                break

    previous_progress = previous.get("weekly_progress") if previous else None
    weekly_delta = None
    if progress_value is not None and previous_progress is not None:
        weekly_delta = progress_value - previous_progress

    program_health = "Green"
    if overdue_actions > 3 or high_risks >= 1:
        program_health = "Red"
    elif overdue_actions >= 1 or open_risks >= 1:
        program_health = "Amber"

    return {
        "program_health": program_health,
        "open_actions": open_actions,
        "overdue_actions": overdue_actions,
        "open_risks": open_risks,
        "high_risks": high_risks,
        "weekly_progress": progress_value,
        "weekly_delta": weekly_delta,
        "due_next_7_days": due_next_7_days,
    }

@dataclass
class Snapshot:
    filename: str
    project_name: str
    project_no: str
    report_date: object
    reporting_staff: str
    summary: pd.DataFrame
    actions: pd.DataFrame
    risks: pd.DataFrame
    transmittal: pd.DataFrame
    next_steps: pd.DataFrame
    sheets: dict
    kpis: dict

def add_snapshot(parsed, kpis):
    st.session_state.snapshots.append(
        Snapshot(
            filename=parsed["filename"],
            project_name=parsed["project_name"],
            project_no=parsed["project_no"],
            report_date=parsed["report_date"],
            reporting_staff=parsed["reporting_staff"],
            summary=parsed["summary"],
            actions=parsed["actions"],
            risks=parsed["risks"],
            transmittal=parsed["transmittal"],
            next_steps=parsed["next_steps"],
            sheets=parsed["sheets"],
            kpis=kpis,
        )
    )
    st.session_state.uploaded_names.add(parsed["filename"])

def latest_snapshot():
    return st.session_state.snapshots[-1] if st.session_state.snapshots else None

def kpi_card(title, value, delta=None, accent=PRIMARY_RED):
    delta_text = "" if delta is None else f"<div class='kpi-delta'>{delta}</div>"
    st.markdown(
        f"""
        <div class="kpi-card" style="border-top: 4px solid {accent};">
            <div class="kpi-title">{title}</div>
            <div class="kpi-value">{value}</div>
            {delta_text}
        </div>
        """,
        unsafe_allow_html=True,
    )

def render_header():
    snap = latest_snapshot()
    c1, c2, c3 = st.columns([1.2, 2.6, 1.2])
    with c1:
        show_image("AEROBOND-logo_tagline-AV-DEF-SP_RGB-2.png")
    with c2:
        st.markdown(
            f"""
            <div class="brand-box">
                <div style="font-size: 1.2rem; font-weight: 700; color: {PRIMARY_RED};">Aerobond Defence Program Dashboard</div>
                <div style="color: {GRAY};">Weekly reporting, actions, risks, production and history</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with c3:
        if snap:
            st.markdown(
                f"""
                <div class="brand-box">
                    <div style="font-size: 0.85rem; color: {GRAY};">Current snapshot</div>
                    <div style="font-weight: 700;">{snap.filename}</div>
                    <div style="color: {GRAY};">{snap.project_name}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

def overview_page():
    snap = latest_snapshot()
    if not snap:
        st.info("Upload one or more Excel reports to start.")
        return

    k = snap.kpis
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    with c1: kpi_card("Program Health", k["program_health"], accent=GREEN if k["program_health"] == "Green" else AMBER if k["program_health"] == "Amber" else RED)
    with c2: kpi_card("Weekly Progress", k["weekly_progress"] if k["weekly_progress"] is not None else "—", f"Δ {k['weekly_delta']}" if k["weekly_delta"] is not None else None)
    with c3: kpi_card("Open Actions", k["open_actions"], f"Overdue: {k['overdue_actions']}", accent=AMBER if k["overdue_actions"] else GREEN)
    with c4: kpi_card("Open Risks", k["open_risks"], f"High: {k['high_risks']}", accent=RED if k["high_risks"] else GREEN)
    with c5: kpi_card("Due Next 7 Days", k["due_next_7_days"], accent=PRIMARY_RED)
    with c6: kpi_card("Project", snap.project_name, accent=PRIMARY_RED)

    st.markdown("<div class='section-card'>", unsafe_allow_html=True)
    st.subheader("Weekly trend")
    hist = pd.DataFrame([
        {"date": s.report_date, "progress": s.kpis.get("weekly_progress")}
        for s in st.session_state.snapshots
    ]).dropna(subset=["date"])
    if len(hist) >= 2:
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=hist["date"], y=hist["progress"], mode="lines+markers", name="Progress", line=dict(color=PRIMARY_RED, width=3)))
        fig.update_layout(height=360, margin=dict(l=10, r=10, t=20, b=10), plot_bgcolor="white", paper_bgcolor="white")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Add at least two uploads to see trends.")
    st.markdown("</div>", unsafe_allow_html=True)

    a, r, p = st.columns(3)
    with a:
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        st.subheader("Actions")
        st.dataframe(snap.actions, use_container_width=True, height=320)
        st.markdown("</div>", unsafe_allow_html=True)
    with r:
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        st.subheader("Risks")
        st.dataframe(snap.risks, use_container_width=True, height=320)
        st.markdown("</div>", unsafe_allow_html=True)
    with p:
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        st.subheader("Production")
        st.dataframe(snap.summary, use_container_width=True, height=320)
        st.markdown("</div>", unsafe_allow_html=True)

def uploads_page():
    st.subheader("Upload reports")
    uploaded = st.file_uploader("Upload one or more Excel files", type=["xlsx"], accept_multiple_files=True)
    if uploaded:
        for f in uploaded:
            if f.name in st.session_state.uploaded_names:
                st.warning(f"{f.name} already uploaded in this session.")
                continue
            try:
                parsed = parse_snapshot(f)
                prev = latest_snapshot()
                kpis = compute_kpis(parsed, previous=prev.kpis if prev else None)
                add_snapshot(parsed, kpis)
                st.success(f"Loaded {f.name}")
            except Exception as e:
                st.error(f"Failed to read {f.name}: {e}")

    if st.session_state.snapshots:
        rows = []
        for s in st.session_state.snapshots:
            rows.append({
                "filename": s.filename,
                "project_name": s.project_name,
                "report_date": s.report_date,
                "program_health": s.kpis.get("program_health"),
                "open_actions": s.kpis.get("open_actions"),
                "open_risks": s.kpis.get("open_risks"),
            })
        st.dataframe(pd.DataFrame(rows), use_container_width=True)

def simple_table_page(title, df):
    st.subheader(title)
    st.dataframe(df, use_container_width=True, height=600)

def history_page():
    if not st.session_state.snapshots:
        st.info("No snapshots yet.")
        return
    rows = []
    for s in st.session_state.snapshots:
        rows.append({
            "filename": s.filename,
            "project_name": s.project_name,
            "report_date": s.report_date,
            "weekly_progress": s.kpis.get("weekly_progress"),
            "open_actions": s.kpis.get("open_actions"),
            "open_risks": s.kpis.get("open_risks"),
        })
    st.subheader("Snapshot history")
    st.dataframe(pd.DataFrame(rows), use_container_width=True)

def main():
    init_state()
    with st.sidebar:
        show_image("AEROBOND-logo_tagline-AV-DEF-SP_RGB-2.png")
        if not st.session_state.auth:
            check_password()
        else:
            st.success("Authenticated")
        st.divider()
        page = st.radio("Navigate", ["Overview", "Uploads", "Actions", "Risks", "Production", "History"])
        st.divider()
        if st.button("Sign out"):
            st.session_state.auth = False
            st.rerun()

    if not st.session_state.auth:
        st.title("Aerobond Defence Dashboard")
        st.info("Enter the password in the sidebar to continue.")
        return

    render_header()

    snap = latest_snapshot()
    if page == "Overview":
        overview_page()
    elif page == "Uploads":
        uploads_page()
    elif page == "Actions":
        if snap: simple_table_page("Actions", snap.actions)
        else: st.info("Upload a report first.")
    elif page == "Risks":
        if snap: simple_table_page("Risks", snap.risks)
        else: st.info("Upload a report first.")
    elif page == "Production":
        if snap: simple_table_page("Production", snap.summary)
        else: st.info("Upload a report first.")
    elif page == "History":
        history_page()

if __name__ == "__main__":
    main()
