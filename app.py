"""
Job Flow Tracker — Streamlit app
Run:  streamlit run app.py
Team: deploy to a shared server or Streamlit Community Cloud (see README.md)
"""

import sqlite3
import uuid
import io
from datetime import date, datetime

import pandas as pd
import streamlit as st

# ─── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Job Flow Tracker",
    page_icon="🖨️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── Stage definitions (from Flow_planning.xlsx) ──────────────────────────────
STAGES = [
    {"id": "job_creation",    "label": "Job Creation",     "icon": "📋", "fields": [
        {"key": "job_number",        "label": "Job Number",       "type": "text", "required": True},
        {"key": "job_creation_date", "label": "Creation Date",    "type": "date"},
        {"key": "job_created_by",    "label": "Created By",       "type": "text"},
        {"key": "date_of_delivery",  "label": "Date of Delivery", "type": "date"},
    ]},
    {"id": "planning",        "label": "Planning",          "icon": "🗓️", "fields": [
        {"key": "date_of_printing", "label": "Date of Printing", "type": "date"},
    ]},
    {"id": "prepress",        "label": "Pre-Press",         "icon": "🎨", "fields": [
        {"key": "pdf_conformity",      "label": "PDF Conformity Date", "type": "date"},
        {"key": "ctp_submission_date", "label": "CTP Submission Date", "type": "date"},
        {"key": "prepress_user",       "label": "Prepress User",       "type": "text"},
    ]},
    {"id": "workflow",        "label": "Workflow",          "icon": "⚙️", "fields": [
        {"key": "plates_manufacturing", "label": "Plates Manufacturing Date", "type": "date"},
        {"key": "plates_sent_to_press", "label": "Plates Sent to Press",      "type": "date"},
    ]},
    {"id": "store_logistics", "label": "Store & Logistics", "icon": "📦", "fields": [
        {"key": "paper_store_issue",             "label": "Paper Store Issue Date",        "type": "date"},
        {"key": "paper_delivered_to_production", "label": "Paper Delivered to Production", "type": "date"},
    ]},
    {"id": "press",     "label": "Press",     "icon": "🖨️", "fields": [
        {"key": "printing_date", "label": "Printing Date", "type": "date"},
    ]},
    {"id": "cutting",   "label": "Cutting",   "icon": "✂️", "fields": [
        {"key": "cutting_date", "label": "Cutting Date", "type": "date"},
    ]},
    {"id": "finishing", "label": "Finishing", "icon": "🏁", "fields": [
        {"key": "finishing_date", "label": "Finishing Date", "type": "date"},
    ]},
    {"id": "delivery",  "label": "Delivery",  "icon": "🚚", "fields": [
        {"key": "delivered_to_store",  "label": "Delivered to Store",  "type": "date"},
        {"key": "delivered_to_client", "label": "Delivered to Client", "type": "date"},
    ]},
    {"id": "invoice",   "label": "Invoice",   "icon": "🧾", "fields": [
        {"key": "invoice_date", "label": "Invoice Date", "type": "date"},
    ]},
]

ALL_FIELDS  = [f["key"] for s in STAGES for f in s["fields"]]
DATE_FIELDS = {f["key"] for s in STAGES for f in s["fields"] if f["type"] == "date"}

# Excel column → field key mapping (case-insensitive)
EXCEL_COL_MAP = {
    "job number": "job_number", "job creation date": "job_creation_date",
    "job created by": "job_created_by", "date of delivery": "date_of_delivery",
    "date of printing": "date_of_printing", "pdf conformity": "pdf_conformity",
    "ctp submission date": "ctp_submission_date", "prepress user": "prepress_user",
    "plates manufacturing": "plates_manufacturing", "plates sent to press dept": "plates_sent_to_press",
    "paper store issue": "paper_store_issue", "paper delivered to production": "paper_delivered_to_production",
    "printing date": "printing_date", "cutting date": "cutting_date",
    "finishing date": "finishing_date", "delivered to store": "delivered_to_store",
    "delivered to client": "delivered_to_client", "invoice date": "invoice_date",
}

# ─── Database ──────────────────────────────────────────────────────────────────
DB_PATH = "jobs.db"

@st.cache_resource
def get_db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    extra = ", ".join(f"{f} TEXT DEFAULT ''" for f in ALL_FIELDS if f != "job_number")
    conn.execute(f"""
        CREATE TABLE IF NOT EXISTS jobs (
            id TEXT PRIMARY KEY,
            job_number TEXT UNIQUE NOT NULL,
            {extra},
            created_at TEXT DEFAULT (datetime('now'))
        )
    """)
    conn.commit()
    return conn

def load_jobs():
    return [dict(r) for r in get_db().execute("SELECT * FROM jobs ORDER BY created_at DESC").fetchall()]

def upsert_job(job: dict):
    db   = get_db()
    keys = ["id", "job_number"] + [k for k in ALL_FIELDS if k != "job_number" and k in job]
    ph   = ", ".join(f":{k}" for k in keys)
    ups  = ", ".join(f"{k}=:{k}" for k in keys if k not in ("id","job_number"))
    db.execute(
        f"INSERT INTO jobs ({','.join(keys)}) VALUES ({ph}) ON CONFLICT(job_number) DO UPDATE SET {ups}",
        {k: job.get(k,"") for k in keys}
    )
    db.commit()

def delete_job(job_id: str):
    db = get_db()
    db.execute("DELETE FROM jobs WHERE id=?", (job_id,))
    db.commit()

# ─── Helpers ──────────────────────────────────────────────────────────────────
def stage_progress(job: dict) -> int:
    for i in range(len(STAGES)-1, -1, -1):
        if any(job.get(f["key"],"") for f in STAGES[i]["fields"]):
            return i
    return -1

def status_label(job: dict) -> str:
    i = stage_progress(job)
    if i == len(STAGES)-1: return "✅ Complete"
    if i >= 0:              return f"🔵 {STAGES[i]['label']}"
    return "⚪ New"

def progress_pct(job: dict) -> float:
    i = stage_progress(job)
    return 0.0 if i < 0 else (i+1) / len(STAGES)

def parse_date(val) -> date | None:
    if not val or str(val).strip() in ("","nan","NaT","None"): return None
    if isinstance(val, datetime): return val.date()
    if isinstance(val, date):     return val
    try: return datetime.strptime(str(val)[:10], "%Y-%m-%d").date()
    except: return None

def to_str(val) -> str:
    if val is None: return ""
    if isinstance(val, (date, datetime)): return val.strftime("%Y-%m-%d")
    return str(val)

# ─── CSS ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stAppViewContainer"] > .main { background: #F0EDE8; }
[data-testid="stSidebar"] { background: #1A2130 !important; }
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] span { color: #C8C4BC !important; }
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 { color: #E8E4DC !important; }
[data-testid="stSidebar"] [data-testid="stMetricValue"] { color: #FFFFFF !important; }
[data-testid="stSidebar"] div[role="radiogroup"] label { color: #E8E4DC !important; }
.stTabs [data-baseweb="tab"] { font-size: 0.78rem; padding: 6px 10px; }
button[kind="primary"] { background-color: #1C5FA5 !important; }
[data-testid="stDataFrame"] { border-radius: 8px; }
</style>
""", unsafe_allow_html=True)

# ─── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🖨️ Job Flow Tracker")
    st.divider()
    page = st.radio("nav", ["📋  All Jobs", "➕  New Job", "📥  Import Excel", "📤  Export"],
                    label_visibility="collapsed")
    st.divider()
    jobs = load_jobs()
    col_a, col_b = st.columns(2)
    col_a.metric("Total", len(jobs))
    col_b.metric("Complete", sum(1 for j in jobs if stage_progress(j)==len(STAGES)-1))
    st.metric("In Progress", sum(1 for j in jobs if 0 <= stage_progress(j) < len(STAGES)-1))

# ─── All Jobs ─────────────────────────────────────────────────────────────────
if "All Jobs" in page:
    st.title("All Jobs")

    c1, c2 = st.columns([3,1])
    search  = c1.text_input("Search", placeholder="Search by job number or creator…", label_visibility="collapsed")
    f_stat  = c2.selectbox("Filter", ["All","New","In Progress","Complete"], label_visibility="collapsed")

    filtered = jobs
    if search:
        s = search.lower()
        filtered = [j for j in filtered if s in (j.get("job_number","") or "").lower()
                    or s in (j.get("job_created_by","") or "").lower()]
    if f_stat == "New":         filtered = [j for j in filtered if stage_progress(j)==-1]
    elif f_stat == "In Progress": filtered = [j for j in filtered if 0 <= stage_progress(j) < len(STAGES)-1]
    elif f_stat == "Complete":  filtered = [j for j in filtered if stage_progress(j)==len(STAGES)-1]

    if not filtered:
        st.info("No jobs found. Use ➕ New Job or 📥 Import Excel to get started.")
    else:
        st.dataframe(pd.DataFrame([{
            "Job Number":  j.get("job_number",""),
            "Created By":  j.get("job_created_by","—") or "—",
            "Delivery":    j.get("date_of_delivery","—") or "—",
            "Status":      status_label(j),
            "Progress":    f"{round(progress_pct(j)*100)}%",
        } for j in filtered]), use_container_width=True, hide_index=True)

        st.divider()
        st.subheader("Update a Job")

        sel_num = st.selectbox("Select job to edit", ["— select —"] + [j.get("job_number","") for j in filtered],
                               label_visibility="collapsed")
        if sel_num and sel_num != "— select —":
            job = next((j for j in jobs if j.get("job_number")==sel_num), None)
            if job:
                pct   = progress_pct(job)
                s_idx = stage_progress(job)
                st.markdown(f"**{job.get('job_number','')}** &nbsp;—&nbsp; {status_label(job)}")
                st.progress(pct, text=f"{round(pct*100)}% complete · "
                            f"{'Stage '+str(s_idx+1)+'/'+str(len(STAGES)) if s_idx>=0 else 'Not started'}")
                st.write("")

                tabs = st.tabs([f"{s['icon']} {s['label']}" for s in STAGES])

                for stage, tab in zip(STAGES, tabs):
                    done = all(job.get(f["key"],"") for f in stage["fields"])
                    with tab:
                        if done:
                            st.success("Stage complete ✓")

                        with st.form(f"frm_{job['id']}_{stage['id']}"):
                            vals   = {}
                            ncols  = min(len(stage["fields"]), 2)
                            cols   = st.columns(ncols)
                            for i, field in enumerate(stage["fields"]):
                                lbl = ("\\* " if field.get("required") else "") + field["label"]
                                wkey = f"{job['id']}_{stage['id']}_{field['key']}"
                                with cols[i % ncols]:
                                    if field["type"] == "date":
                                        vals[field["key"]] = st.date_input(
                                            lbl, value=parse_date(job.get(field["key"])),
                                            format="DD/MM/YYYY", key=wkey)
                                    else:
                                        vals[field["key"]] = st.text_input(
                                            lbl, value=job.get(field["key"],"") or "", key=wkey)

                            do_save = st.form_submit_button("💾 Save stage", type="primary", use_container_width=True)
                            if do_save:
                                updated = {**job, **{k: to_str(v) if isinstance(v, date) else (v or "") for k,v in vals.items()}}
                                upsert_job(updated)
                                st.success("Saved!")
                                st.rerun()

                # Delete zone (outside form)
                with st.expander("⚠️ Danger zone"):
                    st.warning(f"Delete **{job.get('job_number','')}**? This cannot be undone.")
                    if st.button("Yes, delete this job", key=f"del_{job['id']}"):
                        delete_job(job["id"])
                        st.success("Job deleted.")
                        st.rerun()

# ─── New Job ──────────────────────────────────────────────────────────────────
elif "New Job" in page:
    st.title("Create New Job")

    with st.form("new_job_form"):
        c1, c2 = st.columns(2)
        job_number       = c1.text_input("* Job Number",      placeholder="e.g. JOB-2026-001")
        job_created_by   = c1.text_input("Created By",        placeholder="Your name")
        job_creation_date = c2.date_input("Creation Date",    value=date.today(), format="DD/MM/YYYY")
        date_of_delivery  = c2.date_input("Date of Delivery", value=None, format="DD/MM/YYYY")

        if st.form_submit_button("Create Job →", type="primary", use_container_width=True):
            if not job_number.strip():
                st.error("Job Number is required.")
            else:
                try:
                    upsert_job({
                        "id": str(uuid.uuid4()),
                        "job_number": job_number.strip(),
                        "job_created_by": job_created_by,
                        "job_creation_date": to_str(job_creation_date),
                        "date_of_delivery": to_str(date_of_delivery),
                    })
                    st.success(f"✅ Job **{job_number}** created! Go to All Jobs to fill in the stages.")
                except Exception as e:
                    st.error("Job number already exists." if "UNIQUE" in str(e) else str(e))

# ─── Import Excel ─────────────────────────────────────────────────────────────
elif "Import" in page:
    st.title("Import from Excel")
    st.info("Upload your Flow Planning Excel file. The column headers should be in row 2 "
            "(with stage group labels in row 1), matching the original template.")

    uploaded = st.file_uploader("Choose an Excel file (.xlsx / .xls)", type=["xlsx","xls"])
    if uploaded:
        try:
            raw = pd.read_excel(uploaded, header=None)

            # Find the row that contains "Job Number"
            header_row = None
            for idx, row in raw.iterrows():
                if any(str(c).strip().lower() == "job number" for c in row):
                    header_row = idx
                    break

            if header_row is None:
                st.error("Cannot find a 'Job Number' column header. "
                         "Make sure your file matches the Flow Planning template format.")
            else:
                df = pd.read_excel(uploaded, header=header_row)
                df.columns = df.columns.map(lambda c: str(c).strip().lower())

                # Map & filter columns
                df = df.rename(columns={k: v for k,v in EXCEL_COL_MAP.items() if k in df.columns})
                df = df[[c for c in df.columns if c in ALL_FIELDS]]

                if "job_number" not in df.columns:
                    st.error("No Job Number column found. Check the template headers.")
                    st.stop()

                df["job_number"] = df["job_number"].astype(str).str.strip()
                df = df[df["job_number"].isin([v for v in df["job_number"] if v not in ("","nan")])]

                # Convert dates
                def safe_date(v):
                    if pd.isna(v) or str(v).strip() in ("","nan","NaT","None"): return ""
                    if isinstance(v, (date, datetime)):
                        return (v if isinstance(v, datetime) else datetime.combine(v, datetime.min.time())).strftime("%Y-%m-%d")
                    try: return pd.to_datetime(v).strftime("%Y-%m-%d")
                    except: return str(v).strip()

                for col in DATE_FIELDS:
                    if col in df.columns:
                        df[col] = df[col].apply(safe_date)

                df = df.fillna("").astype(str).replace({"nan":"","NaT":"","None":""})

                # Duplicate detection
                existing = {j.get("job_number","") for j in load_jobs()}
                df["⚠️ Exists"] = df["job_number"].isin(existing)
                n_new = (~df["⚠️ Exists"]).sum()
                n_dup = df["⚠️ Exists"].sum()

                st.write(f"**{len(df)} jobs found** — {n_new} new, {n_dup} already in the tracker")

                preview_cols = [c for c in ["job_number","job_created_by","date_of_delivery","⚠️ Exists"] if c in df.columns]
                st.dataframe(
                    df[preview_cols].rename(columns={
                        "job_number":"Job #","job_created_by":"Created By",
                        "date_of_delivery":"Delivery Date"
                    }),
                    use_container_width=True, hide_index=True
                )

                dup_mode = "Skip"
                if n_dup > 0:
                    choice = st.radio(
                        f"**{n_dup} duplicate(s) found** — what should happen?",
                        ["Skip duplicates (keep existing data)", "Replace with imported data"],
                        horizontal=True,
                    )
                    dup_mode = "Skip" if "Skip" in choice else "Replace"

                if st.button("✅ Confirm Import", type="primary"):
                    imported = 0
                    for _, row in df.iterrows():
                        if row.get("⚠️ Exists") and dup_mode == "Skip":
                            continue
                        data = {k: v for k,v in row.items() if k in ALL_FIELDS and v not in ("","nan","NaT","None")}
                        data["id"] = str(uuid.uuid4())
                        try:
                            upsert_job(data)
                            imported += 1
                        except Exception as ex:
                            st.warning(f"Skipped {row.get('job_number','?')}: {ex}")
                    st.success(f"✅ Imported {imported} job(s).")
                    st.rerun()

        except Exception as e:
            st.error(f"Failed to read file: {e}")

# ─── Export ───────────────────────────────────────────────────────────────────
elif "Export" in page:
    st.title("Export Jobs")

    if not jobs:
        st.info("No jobs to export yet.")
    else:
        label_map = {f["key"]: f["label"] for s in STAGES for f in s["fields"]}
        rows = []
        for j in jobs:
            row = {label_map[k]: j.get(k,"") or "" for k in ALL_FIELDS if k in label_map}
            row["Status"]     = status_label(j).replace("✅ ","").replace("🔵 ","").replace("⚪ ","")
            row["Progress %"] = round(progress_pct(j)*100)
            rows.append(row)

        df_out = pd.DataFrame(rows)
        st.dataframe(df_out, use_container_width=True, hide_index=True)

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df_out.to_excel(writer, index=False, sheet_name="Job Tracker")
        buf.seek(0)

        st.download_button(
            "📥 Download as Excel",
            data=buf.getvalue(),
            file_name=f"job_tracker_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
