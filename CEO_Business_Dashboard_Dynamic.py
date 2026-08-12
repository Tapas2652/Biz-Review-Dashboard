
import os
import re
from pathlib import Path
from datetime import datetime

import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# ============================================================
# CEO BUSINESS DASHBOARD
# Direct Excel connection + automatic refresh
#
# Put this .py file in the same folder as the Excel workbook,
# OR set the environment variable:
#   CEO_DASHBOARD_XLSX = r"C:\...\CEO's Review dashboard (1).xlsx"
#
# The dashboard checks the Excel file timestamp every 60 sec.
# When the workbook is saved/updated, the next refresh reloads
# the data automatically.
# ============================================================

st.set_page_config(
    page_title="CEO Business Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

try:
    from streamlit_autorefresh import st_autorefresh
    AUTO_REFRESH_AVAILABLE = True
except Exception:
    AUTO_REFRESH_AVAILABLE = False


# -----------------------------
# Configuration
# -----------------------------
DEFAULT_FILE = Path(r"C:\Users\E36250360\OneDrive - JoulestoWatts Business Solutions Pvt Ltd\CEO's Review dashboard (1).xlsx")
EXCEL_PATH = Path(os.getenv("CEO_DASHBOARD_XLSX", str(DEFAULT_FILE)))

DISPLAY_BH = [
    "Sadhna",
    "Mehr",
    "Prathap",
    "Anuradha",
    "Deepak",
]

BH_ALIASES = {
    "sadhna": "Sadhna",
    "sadhna shukla": "Sadhna",
    "mehr": "Mehr",
    "mehr hashim": "Mehr",
    "prathap": "Prathap",
    "prathap sagar": "Prathap",
    "anuradha": "Anuradha",
    "anuradha murthy": "Anuradha",
    "deepak": "Deepak",
    "deepak desai": "Deepak",
    "deepak review": "Deepak",
}


# -----------------------------
# Styling
# -----------------------------
st.markdown(
    """
<style>
    .block-container {padding-top: 1rem; padding-bottom: 2rem;}
    .hero {
        background: linear-gradient(135deg,#0d2f4f,#173e69);
        color:white; padding:22px 26px; border-radius:16px;
        margin-bottom:18px;
    }
    .hero h1 {margin:0; font-size:30px;}
    .hero p {margin:6px 0 0; color:#d9e7f3;}
    .small-muted {color:#6b7280; font-size:12px;}
    .section {
        font-size:19px; font-weight:800; color:#0d2f4f;
        margin:20px 0 8px;
    }
    div[data-testid="stMetric"] {
        background:#fff; border:1px solid #d8dee6;
        border-radius:13px; padding:12px;
        box-shadow:0 1px 4px rgba(0,0,0,.04);
    }
    .bh-card {
        background:#fff; border:1px solid #d8dee6;
        border-radius:14px; padding:14px;
    }
    .bh-name {font-weight:850; color:#0d2f4f; font-size:16px;}
    .bh-value {font-weight:900; color:#0d2f4f; font-size:26px;}
    .positive {color:#16723b;}
    .negative {color:#a52626;}
</style>
""",
    unsafe_allow_html=True,
)


# -----------------------------
# Helpers
# -----------------------------
def clean_text(v):
    if pd.isna(v):
        return ""
    return re.sub(r"\s+", " ", str(v)).strip()


def clean_key(v):
    return clean_text(v).lower()


def normalize_bh(v):
    key = clean_key(v)
    return BH_ALIASES.get(key, clean_text(v))


def numeric(s):
    return pd.to_numeric(s, errors="coerce").fillna(0)


MONTH_RE = re.compile(
    r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*['\-]?\s*(\d{2,4})",
    re.I,
)


def normalize_month_value(v):
    if pd.isna(v):
        return None

    if isinstance(v, pd.Timestamp):
        return v.strftime("%b'%y")

    if isinstance(v, datetime):
        return v.strftime("%b'%y")

    s = clean_text(v)
    m = MONTH_RE.search(s)
    if m:
        yy = m.group(2)
        if len(yy) == 4:
            yy = yy[-2:]
        return f"{m.group(1).title()}'{yy}"

    dt = pd.to_datetime(s, errors="coerce")
    if not pd.isna(dt):
        return dt.strftime("%b'%y")

    return s


def month_sort_key(label):
    try:
        return datetime.strptime(label, "%b'%y")
    except Exception:
        return datetime(1900, 1, 1)


def month_order(values):
    vals = [clean_text(x) for x in values if clean_text(x)]
    return sorted(set(vals), key=month_sort_key)


def fmt_l(x):
    return f"₹{float(x):,.1f}L"


def fmt_num(x):
    return f"{float(x):,.0f}"


def safe_pct(num, den):
    return (num / den * 100) if den else 0


def standardize_columns(df):
    df = df.copy()
    df.columns = [clean_text(c) for c in df.columns]
    return df


def first_existing(df, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None


def read_selected_sheet(xl, sheet, columns):
    """Read only the columns needed for the dashboard."""
    header = pd.read_excel(xl, sheet_name=sheet, nrows=0)
    available = [clean_text(c) for c in header.columns]
    wanted = [c for c in columns if c in available]
    if not wanted:
        return pd.DataFrame()
    return standardize_columns(
        pd.read_excel(xl, sheet_name=sheet, usecols=wanted)
    )


def add_common_fields(
    df,
    client_candidates,
    bh_candidates=("BH", "Business Head"),
    month_candidates=("Month",),
    date_candidates=(),
):
    df = df.copy()
    if df.empty:
        return df

    client_col = first_existing(df, client_candidates)
    bh_col = first_existing(df, bh_candidates)
    month_col = first_existing(df, month_candidates)
    date_col = first_existing(df, date_candidates)

    df["_client"] = df[client_col].map(clean_text) if client_col else ""
    df["_client_key"] = df["_client"].map(clean_key)
    df["_bh"] = df[bh_col].map(normalize_bh) if bh_col else ""
    df["_month"] = (
        df[month_col].map(normalize_month_value)
        if month_col
        else (df[date_col].map(normalize_month_value) if date_col else None)
    )
    return df[
        df["_bh"].isin(DISPLAY_BH)
        & df["_client_key"].ne("")
    ].copy()


# -----------------------------
# Excel loading
# -----------------------------
@st.cache_data(show_spinner="Loading CEO dashboard data from Excel...")
def load_workbook_data(path_str, file_mtime, file_size):
    path = Path(path_str)
    xl = pd.ExcelFile(path)

    data = {}

    data["Demand"] = add_common_fields(
        read_selected_sheet(
            xl,
            "Demand",
            ["company_name", "Created_at", "Month", "BH", "no_of_opening"],
        ),
        ["company_name"],
        date_candidates=["Created_at"],
    )

    data["Submissions"] = add_common_fields(
        read_selected_sheet(
            xl, "Submissions", ["client", "date", "Month", "BH"]
        ),
        ["client"],
        date_candidates=["date"],
    )

    data["Interviews"] = add_common_fields(
        read_selected_sheet(
            xl,
            "Interviews",
            ["company_name", "interview_date", "Month", "BH"],
        ),
        ["company_name"],
        date_candidates=["interview_date"],
    )

    data["Selections"] = add_common_fields(
        read_selected_sheet(
            xl,
            "Selections",
            [
                "candidate", "joining_date", "selection_date",
                "company_name", "po", "margin", "Month", "BH"
            ],
        ),
        ["company_name"],
        date_candidates=["selection_date", "joining_date"],
    )

    data["Onboarding"] = add_common_fields(
        read_selected_sheet(
            xl,
            "Onboarding",
            [
                "full_name", "display_date", "p_o_value", "margin",
                "company_name", "Selection_date", "Month", "BH"
            ],
        ),
        ["company_name"],
        date_candidates=["display_date", "Selection_date"],
    )

    data["Exit"] = add_common_fields(
        read_selected_sheet(
            xl,
            "Exit",
            [
                "full_name", "joining_date", "p_o_value", "margin",
                "company_name", "last_work_day", "created_at", "Month", "BH"
            ],
        ),
        ["company_name"],
        date_candidates=["last_work_day", "created_at"],
    )

    data["Exit in Progress"] = add_common_fields(
        read_selected_sheet(
            xl,
            "Exit in Progress",
            [
                "full_name", "joining_date", "p_o_value", "margin",
                "company_name", "tentative_exit_date", "Month", "BH"
            ],
        ),
        ["company_name"],
        date_candidates=["tentative_exit_date", "joining_date"],
    )

    data["OB Pipeline"] = add_common_fields(
        read_selected_sheet(
            xl,
            "OB Pipeline",
            [
                "full_name", "display_date", "p_o_value", "margin",
                "Company_name", "Selection_date", "Month", "BH"
            ],
        ),
        ["Company_name", "company_name"],
        date_candidates=["display_date", "Selection_date"],
    )

    active = read_selected_sheet(
        xl,
        "Active Head Count",
        [
            "recruiter_name1", "Lead1", "manager_name1", "full_name",
            "employee_type", "ctc", "contact_phone", "email", "employee_id",
            "display_date", "p_o_value", "margin", "company_name", "Status",
            "designation", "po_end_date", "total_experience", "name",
            "Month", "Domain", "Domain1", "BH", "HRBP", "Margin %", "Bucket"
        ],
    )
    active = add_common_fields(
        active,
        ["company_name"],
        date_candidates=["display_date"],
    )
    if not active.empty:
        emp_col = first_existing(active, ["employee_id", "full_name", "name"])
        active["_employee_key"] = (
            active[emp_col].map(clean_key) if emp_col else
            pd.Series(range(len(active)), index=active.index).astype(str)
        )
    data["Active Head Count"] = active

    closure = read_selected_sheet(
        xl,
        "Contract Closure Rawa Data",
        [
            "Emp ID", "Impacted Headcount", "DOJ", "LWD Status",
            "Exit Receievd", "Reason for Closure ", "PO Value(MRR Impact)",
            "Key Account names", "Business Head", "Relationship Manager",
            "HRBP", "PO Start Date", "Current PO End Date",
            "Extensions Start date", "Extensions END date", "Month End Date",
            "Client Type", "Finance Team", "Final Status (As of today)", "Status"
        ],
    )
    if not closure.empty:
        closure["_client"] = closure["Key Account names"].map(clean_text)
        closure["_client_key"] = closure["_client"].map(clean_key)
        closure["_bh"] = closure["Business Head"].map(normalize_bh)
        closure["_month"] = closure["Month End Date"].map(normalize_month_value)
        closure = closure[
            closure["_bh"].isin(DISPLAY_BH) & closure["_client_key"].ne("")
        ].copy()
    data["Contract Closure"] = closure

    # Signed Clients is a wide monthly snapshot. Convert it to a long table.
    signed = standardize_columns(pd.read_excel(xl, sheet_name="Signed Clients Data"))
    signed_long = []
    if not signed.empty:
        client_col = signed.columns[0]
        bh_col = signed.columns[6] if len(signed.columns) > 6 else None

        for col_idx in range(7, min(len(signed.columns), 64), 3):
            if col_idx + 2 >= len(signed.columns):
                break
            month_label = normalize_month_value(
                signed.columns[col_idx - 1] if False else
                None
            )

        # Month labels are in row 1 of the workbook, so read the sheet without header.
        raw_signed = pd.read_excel(xl, sheet_name="Signed Clients Data", header=None)
        for start in range(7, 64, 3):
            if start + 2 >= raw_signed.shape[1]:
                break

            # The month name sits in the first row at the start-1 position.
            month_label = normalize_month_value(raw_signed.iloc[0, start])
            if not month_label:
                # Search nearby header cells.
                for j in range(max(0, start - 1), min(raw_signed.shape[1], start + 1)):
                    month_label = normalize_month_value(raw_signed.iloc[0, j])
                    if month_label:
                        break
            if not month_label:
                continue

            temp = pd.DataFrame({
                "_client": raw_signed.iloc[2:, 0].map(clean_text),
                "_bh": raw_signed.iloc[2:, 6].map(normalize_bh),
                "signed_hc": pd.to_numeric(raw_signed.iloc[2:, start], errors="coerce").fillna(0),
                "signed_po": pd.to_numeric(raw_signed.iloc[2:, start + 1], errors="coerce").fillna(0) / 100000,
                "signed_margin": pd.to_numeric(raw_signed.iloc[2:, start + 2], errors="coerce").fillna(0) / 100000,
            })
            temp["_client_key"] = temp["_client"].map(clean_key)
            temp["_month"] = month_label
            temp = temp[
                temp["_bh"].isin(DISPLAY_BH) & temp["_client_key"].ne("")
            ]
            signed_long.append(temp)

    data["Signed Clients"] = (
        pd.concat(signed_long, ignore_index=True)
        if signed_long else pd.DataFrame()
    )

    # Cost sheet contains two sections: Delivery Cost and Total Cost.
    cost_raw = pd.read_excel(xl, sheet_name="Cost", header=None)
    cost_rows = []
    section = None
    if not cost_raw.empty:
        for _, row in cost_raw.iterrows():
            first = clean_text(row.iloc[0]) if len(row) else ""
            if first.lower() == "delivery cost":
                section = "Delivery Cost"
                continue
            if first.lower() == "total cost":
                section = "Total Cost"
                continue

            bh = normalize_bh(first)
            if section in {"Delivery Cost", "Total Cost"} and bh in DISPLAY_BH:
                for j in range(1, len(row)):
                    month = normalize_month_value(cost_raw.iloc[0, j])
                    if not month:
                        # Cost header may be on the section row.
                        month = normalize_month_value(cost_raw.iloc[0, j])
                    value = pd.to_numeric(row.iloc[j], errors="coerce")
                    if pd.notna(value):
                        cost_rows.append(
                            {
                                "_bh": bh,
                                "_month": month,
                                "cost_type": section,
                                "cost_l": float(value),
                            }
                        )

    # Re-read Cost with its actual month headers when possible.
    if not cost_raw.empty:
        header_candidates = []
        for j in range(1, cost_raw.shape[1]):
            header_candidates.append(normalize_month_value(cost_raw.iloc[0, j]))
        # Current workbook has months on row 0; rebuild cleanly.
        clean_cost = []
        section = None
        for i in range(1, len(cost_raw)):
            first = clean_text(cost_raw.iloc[i, 0])
            if first.lower() == "total cost":
                section = "Total Cost"
                continue
            if first.lower() == "delivery cost":
                section = "Delivery Cost"
                continue
            bh = normalize_bh(first)
            if section and bh in DISPLAY_BH:
                for j in range(1, cost_raw.shape[1]):
                    month = normalize_month_value(cost_raw.iloc[0, j])
                    val = pd.to_numeric(cost_raw.iloc[i, j], errors="coerce")
                    if month and pd.notna(val):
                        clean_cost.append(
                            {"_bh": bh, "_month": month, "cost_type": section, "cost_l": float(val)}
                        )
        data["Cost"] = pd.DataFrame(clean_cost)
    else:
        data["Cost"] = pd.DataFrame()

    return data


# -----------------------------
# Load + refresh
# -----------------------------
if not EXCEL_PATH.exists():
    st.error(
        f"Excel workbook not found:\n\n{EXCEL_PATH}\n\n"
        "Place the workbook beside this Python file or set "
        "CEO_DASHBOARD_XLSX to the exact Excel path."
    )
    st.stop()

file_stat = EXCEL_PATH.stat()
file_mtime = file_stat.st_mtime
file_size = file_stat.st_size

if AUTO_REFRESH_AVAILABLE:
    st_autorefresh(interval=60_000, key="ceo_dashboard_refresh")

with st.sidebar:
    st.markdown("### Dashboard Controls")
    if st.button("🔄 Refresh Excel Data", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    st.caption(f"Workbook: {EXCEL_PATH.name}")
    st.caption(
        "Last file update: "
        + datetime.fromtimestamp(file_mtime).strftime("%d-%b-%Y %I:%M %p")
    )
    st.caption("Auto refresh: every 60 seconds")

data = load_workbook_data(str(EXCEL_PATH), file_mtime, file_size)


# -----------------------------
# Derived fact tables
# -----------------------------
def count_fact(df, metric_name):
    if df.empty:
        return pd.DataFrame(
            columns=["_bh", "_client", "_client_key", "_month", metric_name]
        )
    x = (
        df.groupby(["_bh", "_client", "_client_key", "_month"], dropna=False)
        .size()
        .reset_index(name=metric_name)
    )
    return x


def financial_fact(df, po_col, margin_col, prefix):
    if df.empty:
        return pd.DataFrame(
            columns=["_bh", "_client", "_client_key", "_month",
                     f"{prefix}_hc", f"{prefix}_po", f"{prefix}_margin"]
        )

    x = df.copy()
    x[po_col] = numeric(x[po_col]) / 100000 if po_col in x.columns else 0
    x[margin_col] = numeric(x[margin_col]) / 100000 if margin_col in x.columns else 0

    return (
        x.groupby(["_bh", "_client", "_client_key", "_month"], dropna=False)
        .agg(
            **{
                f"{prefix}_hc": ("_client_key", "size"),
                f"{prefix}_po": (po_col, "sum"),
                f"{prefix}_margin": (margin_col, "sum"),
            },
        )
        .reset_index()
    )


facts = [
    count_fact(data["Demand"], "demand"),
    count_fact(data["Submissions"], "submissions"),
    count_fact(data["Interviews"], "interviews"),
    count_fact(data["Selections"], "selections"),
    financial_fact(data["Onboarding"], "p_o_value", "margin", "ob"),
    financial_fact(data["OB Pipeline"], "p_o_value", "margin", "pipeline"),
    financial_fact(data["Exit"], "p_o_value", "margin", "exit"),
    count_fact(data["Exit in Progress"], "eip"),
]

fact = None
for f in facts:
    fact = f if fact is None else fact.merge(
        f,
        on=["_bh", "_client", "_client_key", "_month"],
        how="outer",
    )

if fact is None:
    fact = pd.DataFrame()

numeric_cols = [
    c for c in fact.columns
    if c not in {"_bh", "_client", "_client_key", "_month"}
]
for c in numeric_cols:
    fact[c] = pd.to_numeric(fact[c], errors="coerce").fillna(0)

for c in [
    "demand", "submissions", "interviews", "selections",
    "ob_hc", "ob_po", "ob_margin",
    "pipeline_hc", "pipeline_po", "pipeline_margin",
    "exit_hc", "exit_po", "exit_margin", "eip"
]:
    if c not in fact.columns:
        fact[c] = 0

fact["net_hc"] = fact["ob_hc"] - fact["exit_hc"]
fact["net_po"] = fact["ob_po"] - fact["exit_po"]
fact["net_margin"] = fact["ob_margin"] - fact["exit_margin"]


# Active headcount / current book
active = data["Active Head Count"].copy()
if not active.empty:
    active["p_o_value"] = numeric(active.get("p_o_value", 0)) / 100000
    active["margin"] = numeric(active.get("margin", 0)) / 100000

    active_group = (
        active.groupby(["_bh", "_client", "_client_key"], dropna=False)
        .agg(
            active_hc=("_employee_key", "nunique"),
            active_po=("p_o_value", "sum"),
            active_margin=("margin", "sum"),
        )
        .reset_index()
    )
else:
    active_group = pd.DataFrame(
        columns=["_bh", "_client", "_client_key",
                 "active_hc", "active_po", "active_margin"]
    )


# Contract closure
closure = data["Contract Closure"].copy()
if not closure.empty:
    closure["closure_po"] = numeric(closure["PO Value(MRR Impact)"]) / 100000
    closure_group = (
        closure.groupby(["_bh", "_client", "_client_key"], dropna=False)
        .agg(
            closure_hc=("_client_key", "size"),
            closure_po=("closure_po", "sum"),
        )
        .reset_index()
    )
else:
    closure_group = pd.DataFrame(
        columns=["_bh", "_client", "_client_key", "closure_hc", "closure_po"]
    )


# Signed clients
signed = data["Signed Clients"].copy()
if not signed.empty:
    signed_group = (
        signed.groupby(["_bh", "_client", "_client_key", "_month"], dropna=False)
        .agg(
            signed_hc=("signed_hc", "sum"),
            signed_po=("signed_po", "sum"),
            signed_margin=("signed_margin", "sum"),
        )
        .reset_index()
    )
else:
    signed_group = pd.DataFrame()


# -----------------------------
# UI Header
# -----------------------------
st.markdown(
    """
<div class="hero">
    <h1>CEO Business Performance Dashboard</h1>
    <p>5-BH executive view | Funnel • Active Book • Margin • Cost • Closures • Signed Clients • Client Drilldown</p>
</div>
""",
    unsafe_allow_html=True,
)

available_months = month_order(
    list(fact["_month"].dropna().unique())
    + (list(signed["_month"].dropna().unique()) if not signed.empty else [])
)

if not available_months:
    st.error("No month values were found in the source data.")
    st.stop()

latest_month = available_months[-1]

c1, c2 = st.columns([1, 3])
with c1:
    selected_month = st.selectbox(
        "Reporting Month",
        available_months,
        index=available_months.index(latest_month),
    )
with c2:
    st.info(
        "Click a BH row or Client row to drill down. "
        "The client drilldown exposes operational raw records by source."
    )


# -----------------------------
# Executive KPIs
# -----------------------------
m = fact[fact["_month"] == selected_month].copy()

total_active_hc = active_group["active_hc"].sum()
total_active_po = active_group["active_po"].sum()
total_active_margin = active_group["active_margin"].sum()
total_active_margin_pct = safe_pct(total_active_margin, total_active_po)

total_demand = m["demand"].sum()
total_subs = m["submissions"].sum()
total_intv = m["interviews"].sum()
total_sel = m["selections"].sum()
total_ob_hc = m["ob_hc"].sum()
total_ob_po = m["ob_po"].sum()
total_ob_margin = m["ob_margin"].sum()
total_pipeline_hc = m["pipeline_hc"].sum()
total_pipeline_po = m["pipeline_po"].sum()
total_pipeline_margin = m["pipeline_margin"].sum()
total_exit_hc = m["exit_hc"].sum()
total_exit_po = m["exit_po"].sum()
total_exit_margin = m["exit_margin"].sum()
total_eip = m["eip"].sum()

closure_hc = closure_group["closure_hc"].sum()
closure_po = closure_group["closure_po"].sum()

if not signed_group.empty:
    signed_m = signed_group[signed_group["_month"] == selected_month]
    signed_hc = signed_m["signed_hc"].sum()
    signed_po = signed_m["signed_po"].sum()
    signed_margin = signed_m["signed_margin"].sum()
else:
    signed_hc = signed_po = signed_margin = 0

st.markdown('<div class="section">Executive Snapshot</div>', unsafe_allow_html=True)

k = st.columns(6)
k[0].metric("Active HC", fmt_num(total_active_hc))
k[1].metric("Active PO", fmt_l(total_active_po))
k[2].metric("Active Margin", fmt_l(total_active_margin))
k[3].metric("Margin / PO", f"{total_active_margin_pct:.1f}%")
k[4].metric("Onboarding HC", fmt_num(total_ob_hc))
k[5].metric("Onboarding Margin", fmt_l(total_ob_margin))

k2 = st.columns(6)
k2[0].metric("Demand", fmt_num(total_demand))
k2[1].metric("Submissions", fmt_num(total_subs))
k2[2].metric("Interviews", fmt_num(total_intv))
k2[3].metric("Selections", fmt_num(total_sel))
k2[4].metric("OB Pipeline HC", fmt_num(total_pipeline_hc))
k2[5].metric("Exit HC", fmt_num(total_exit_hc))

# -----------------------------
# BH scorecard
# -----------------------------
st.markdown('<div class="section">BH Performance — Click a Row to Drill Down</div>', unsafe_allow_html=True)

bh_rows = []
for bh in DISPLAY_BH:
    fm = m[m["_bh"] == bh]
    ag = active_group[active_group["_bh"] == bh]
    cg = closure_group[closure_group["_bh"] == bh]

    row = {
        "BH": bh,
        "Active HC": ag["active_hc"].sum(),
        "Active PO (L)": ag["active_po"].sum(),
        "Active Margin (L)": ag["active_margin"].sum(),
        "Margin %": safe_pct(ag["active_margin"].sum(), ag["active_po"].sum()),
        "Demand": fm["demand"].sum(),
        "Submissions": fm["submissions"].sum(),
        "Interviews": fm["interviews"].sum(),
        "Selections": fm["selections"].sum(),
        "OB HC": fm["ob_hc"].sum(),
        "OB PO (L)": fm["ob_po"].sum(),
        "OB Margin (L)": fm["ob_margin"].sum(),
        "Pipeline HC": fm["pipeline_hc"].sum(),
        "Exit HC": fm["exit_hc"].sum(),
        "EIP HC": fm["eip"].sum(),
        "Closure HC": cg["closure_hc"].sum(),
        "Closure PO (L)": cg["closure_po"].sum(),
    }
    bh_rows.append(row)

bh_df = pd.DataFrame(bh_rows)

bh_event = st.dataframe(
    bh_df,
    use_container_width=True,
    hide_index=True,
    height=280,
    on_select="rerun",
    selection_mode="single-row",
    column_config={
        "Active PO (L)": st.column_config.NumberColumn(format="₹ %.1fL"),
        "Active Margin (L)": st.column_config.NumberColumn(format="₹ %.1fL"),
        "Margin %": st.column_config.NumberColumn(format="%.1f%%"),
        "OB PO (L)": st.column_config.NumberColumn(format="₹ %.1fL"),
        "OB Margin (L)": st.column_config.NumberColumn(format="₹ %.1fL"),
        "Closure PO (L)": st.column_config.NumberColumn(format="₹ %.1fL"),
    },
)

selected_bh = st.session_state.get("selected_bh", "Sadhna")
try:
    selected_rows = bh_event.selection.rows
    if selected_rows:
        selected_bh = bh_df.iloc[selected_rows[0]]["BH"]
        st.session_state["selected_bh"] = selected_bh
except Exception:
    pass


# -----------------------------
# BH trend charts
# -----------------------------
st.markdown('<div class="section">BH Margin & Business Trend</div>', unsafe_allow_html=True)

trend = (
    fact[fact["_bh"] == selected_bh]
    .groupby("_month", as_index=False)
    .agg(
        demand=("demand", "sum"),
        submissions=("submissions", "sum"),
        interviews=("interviews", "sum"),
        selections=("selections", "sum"),
        ob_hc=("ob_hc", "sum"),
        ob_po=("ob_po", "sum"),
        ob_margin=("ob_margin", "sum"),
        exit_hc=("exit_hc", "sum"),
        exit_po=("exit_po", "sum"),
        exit_margin=("exit_margin", "sum"),
        net_hc=("net_hc", "sum"),
        net_po=("net_po", "sum"),
        net_margin=("net_margin", "sum"),
    )
)
trend["sort"] = trend["_month"].map(month_sort_key)
trend = trend.sort_values("sort")
trend["margin_pct"] = trend.apply(
    lambda r: safe_pct(r["net_margin"], r["net_po"]), axis=1
)

t1, t2 = st.columns(2)
with t1:
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=trend["_month"], y=trend["net_margin"],
        mode="lines+markers", name="Net Margin (L)"
    ))
    fig.add_trace(go.Scatter(
        x=trend["_month"], y=trend["ob_margin"],
        mode="lines+markers", name="OB Margin (L)"
    ))
    fig.add_trace(go.Scatter(
        x=trend["_month"], y=-trend["exit_margin"],
        mode="lines+markers", name="Exit Margin (L)"
    ))
    fig.update_layout(
        title=f"{selected_bh} — Margin Trend",
        xaxis_title="", yaxis_title="₹ Lakhs",
        height=360, margin=dict(l=20,r=20,t=55,b=20)
    )
    st.plotly_chart(fig, use_container_width=True)

with t2:
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=trend["_month"], y=trend["net_hc"],
        mode="lines+markers", name="Net HC"
    ))
    fig.add_trace(go.Scatter(
        x=trend["_month"], y=trend["ob_hc"],
        mode="lines+markers", name="OB HC"
    ))
    fig.add_trace(go.Scatter(
        x=trend["_month"], y=-trend["exit_hc"],
        mode="lines+markers", name="Exit HC"
    ))
    fig.update_layout(
        title=f"{selected_bh} — HC Trend",
        xaxis_title="", yaxis_title="Headcount",
        height=360, margin=dict(l=20,r=20,t=55,b=20)
    )
    st.plotly_chart(fig, use_container_width=True)


# -----------------------------
# Cost trend
# -----------------------------
st.markdown('<div class="section">Delivery Cost & Total Cost Trend</div>', unsafe_allow_html=True)

cost = data["Cost"]
if not cost.empty:
    cst = cost[cost["_bh"] == selected_bh].copy()
    cst["sort"] = cst["_month"].map(month_sort_key)
    cst = cst.sort_values("sort")

    fig = px.line(
        cst,
        x="_month",
        y="cost_l",
        color="cost_type",
        markers=True,
        title=f"{selected_bh} — Cost Trend",
        labels={"_month": "", "cost_l": "₹ Lakhs", "cost_type": ""},
    )
    fig.update_layout(height=330)
    st.plotly_chart(fig, use_container_width=True)
else:
    st.warning("Cost sheet did not return usable cost records.")


# -----------------------------
# Client performance
# -----------------------------
st.markdown(
    f'<div class="section">{selected_bh} — Client Performance — Click a Row for Raw Data</div>',
    unsafe_allow_html=True,
)

client_month = m[m["_bh"] == selected_bh].copy()

client_perf = (
    client_month.groupby(["_client", "_client_key"], as_index=False)
    .agg(
        Demand=("demand", "sum"),
        Submissions=("submissions", "sum"),
        Interviews=("interviews", "sum"),
        Selections=("selections", "sum"),
        OB_HC=("ob_hc", "sum"),
        OB_PO=("ob_po", "sum"),
        OB_Margin=("ob_margin", "sum"),
        Pipeline_HC=("pipeline_hc", "sum"),
        Pipeline_PO=("pipeline_po", "sum"),
        Pipeline_Margin=("pipeline_margin", "sum"),
        Exit_HC=("exit_hc", "sum"),
        Exit_PO=("exit_po", "sum"),
        Exit_Margin=("exit_margin", "sum"),
        EIP_HC=("eip", "sum"),
        Net_HC=("net_hc", "sum"),
        Net_PO=("net_po", "sum"),
        Net_Margin=("net_margin", "sum"),
    )
)

if not active_group.empty:
    client_perf = client_perf.merge(
        active_group[active_group["_bh"] == selected_bh][
            ["_client_key", "active_hc", "active_po", "active_margin"]
        ],
        on="_client_key",
        how="outer",
    )
else:
    client_perf["active_hc"] = 0
    client_perf["active_po"] = 0
    client_perf["active_margin"] = 0

if not closure_group.empty:
    client_perf = client_perf.merge(
        closure_group[closure_group["_bh"] == selected_bh][
            ["_client_key", "closure_hc", "closure_po"]
        ],
        on="_client_key",
        how="left",
    )
else:
    client_perf["closure_hc"] = 0
    client_perf["closure_po"] = 0

client_perf = client_perf.fillna(0)
client_perf["Active Margin %"] = client_perf.apply(
    lambda r: safe_pct(r["active_margin"], r["active_po"]), axis=1
)
client_perf["Net Margin %"] = client_perf.apply(
    lambda r: safe_pct(r["Net_Margin"], r["Net_PO"]), axis=1
)

# Keep display order clean.
display_client = client_perf.rename(
    columns={
        "_client": "Client",
        "active_hc": "Active HC",
        "active_po": "Active PO (L)",
        "active_margin": "Active Margin (L)",
        "closure_hc": "Closure HC",
        "closure_po": "Closure PO (L)",
        "OB_HC": "OB HC",
        "OB_PO": "OB PO (L)",
        "OB_Margin": "OB Margin (L)",
        "Pipeline_HC": "Pipeline HC",
        "Pipeline_PO": "Pipeline PO (L)",
        "Pipeline_Margin": "Pipeline Margin (L)",
        "Exit_HC": "Exit HC",
        "Exit_PO": "Exit PO (L)",
        "Exit_Margin": "Exit Margin (L)",
        "EIP_HC": "EIP HC",
        "Net_HC": "Net HC",
        "Net_PO": "Net PO (L)",
        "Net_Margin": "Net Margin (L)",
    }
)

client_cols = [
    "Client", "Active HC", "Active PO (L)", "Active Margin (L)", "Active Margin %",
    "Demand", "Submissions", "Interviews", "Selections",
    "OB HC", "OB PO (L)", "OB Margin (L)",
    "Pipeline HC", "Pipeline PO (L)", "Pipeline Margin (L)",
    "Exit HC", "Exit PO (L)", "Exit Margin (L)",
    "EIP HC", "Net HC", "Net PO (L)", "Net Margin (L)", "Net Margin %",
    "Closure HC", "Closure PO (L)",
]
client_cols = [c for c in client_cols if c in display_client.columns]
display_client = display_client[client_cols].sort_values(
    ["Active PO (L)", "OB PO (L)", "Demand"],
    ascending=False,
)

client_event = st.dataframe(
    display_client,
    use_container_width=True,
    hide_index=True,
    height=470,
    on_select="rerun",
    selection_mode="single-row",
    column_config={
        c: st.column_config.NumberColumn(format="₹ %.1fL")
        for c in [
            "Active PO (L)", "Active Margin (L)",
            "OB PO (L)", "OB Margin (L)",
            "Pipeline PO (L)", "Pipeline Margin (L)",
            "Exit PO (L)", "Exit Margin (L)",
            "Net PO (L)", "Net Margin (L)",
            "Closure PO (L)",
        ]
        if c in display_client.columns
    },
)

selected_client = st.session_state.get("selected_client")

try:
    selected_rows = client_event.selection.rows
    if selected_rows:
        selected_client = display_client.iloc[selected_rows[0]]["Client"]
        st.session_state["selected_client"] = selected_client
except Exception:
    pass


# -----------------------------
# Client margin trend
# -----------------------------
if selected_client:
    st.markdown(
        f'<div class="section">{selected_client} — Client Drilldown</div>',
        unsafe_allow_html=True,
    )

    cc_key = clean_key(selected_client)
    client_trend = (
        fact[
            (fact["_bh"] == selected_bh)
            & (fact["_client_key"] == cc_key)
        ]
        .groupby("_month", as_index=False)
        .agg(
            OB_HC=("ob_hc", "sum"),
            OB_PO=("ob_po", "sum"),
            OB_Margin=("ob_margin", "sum"),
            Exit_HC=("exit_hc", "sum"),
            Exit_PO=("exit_po", "sum"),
            Exit_Margin=("exit_margin", "sum"),
            Net_HC=("net_hc", "sum"),
            Net_PO=("net_po", "sum"),
            Net_Margin=("net_margin", "sum"),
        )
    )
    client_trend["sort"] = client_trend["_month"].map(month_sort_key)
    client_trend = client_trend.sort_values("sort")

    a, b = st.columns(2)
    with a:
        fig = px.line(
            client_trend,
            x="_month",
            y=["OB_Margin", "Exit_Margin", "Net_Margin"],
            markers=True,
            title="Client Margin Trend",
            labels={"value": "₹ Lakhs", "_month": "", "variable": ""},
        )
        st.plotly_chart(fig, use_container_width=True)

    with b:
        fig = px.line(
            client_trend,
            x="_month",
            y=["OB_HC", "Exit_HC", "Net_HC"],
            markers=True,
            title="Client HC Trend",
            labels={"value": "Headcount", "_month": "", "variable": ""},
        )
        st.plotly_chart(fig, use_container_width=True)

    # -------------------------
    # Raw client-level records
    # -------------------------
    st.markdown("### Client-Level Raw Data")

    def raw_table(source, df, extra_filter=None):
        if df.empty:
            st.info(f"No records found in {source}.")
            return

        x = df[
            (df["_bh"] == selected_bh)
            & (df["_client_key"] == cc_key)
        ].copy()

        if extra_filter:
            x = x[extra_filter(x)]

        hidden = [c for c in x.columns if c.startswith("_")]
        x = x.drop(columns=hidden, errors="ignore")

        st.caption(f"{source}: {len(x):,} records")
        st.dataframe(
            x,
            use_container_width=True,
            hide_index=True,
            height=300,
        )

    tabs = st.tabs([
        "Active HC",
        "Demand",
        "Submissions",
        "Interviews",
        "Selections",
        "Onboarding",
        "OB Pipeline",
        "Exit",
        "Exit In Progress",
        "Contract Closure",
        "Signed Clients",
    ])

    with tabs[0]:
        raw_table("Active Head Count", active)

    with tabs[1]:
        raw_table("Demand", data["Demand"])

    with tabs[2]:
        raw_table("Submissions", data["Submissions"])

    with tabs[3]:
        raw_table("Interviews", data["Interviews"])

    with tabs[4]:
        raw_table("Selections", data["Selections"])

    with tabs[5]:
        raw_table("Onboarding", data["Onboarding"])

    with tabs[6]:
        raw_table("OB Pipeline", data["OB Pipeline"])

    with tabs[7]:
        raw_table("Exit", data["Exit"])

    with tabs[8]:
        raw_table("Exit In Progress", data["Exit in Progress"])

    with tabs[9]:
        x = closure[
            (closure["_bh"] == selected_bh)
            & (closure["_client_key"] == cc_key)
        ].drop(columns=[c for c in closure.columns if c.startswith("_")], errors="ignore")
        st.caption(f"Contract Closure: {len(x):,} records")
        st.dataframe(x, use_container_width=True, hide_index=True, height=300)

    with tabs[10]:
        if signed.empty:
            st.info("No Signed Clients records found.")
        else:
            x = signed[
                (signed["_bh"] == selected_bh)
                & (signed["_client_key"] == cc_key)
            ].drop(
                columns=[c for c in signed.columns if c.startswith("_")],
                errors="ignore",
            )
            st.caption(f"Signed Clients monthly records: {len(x):,}")
            st.dataframe(x, use_container_width=True, hide_index=True, height=300)


# -----------------------------
# Signed clients trend
# -----------------------------
st.markdown('<div class="section">Signed Client Portfolio Trend</div>', unsafe_allow_html=True)

if not signed_group.empty:
    s = (
        signed_group[signed_group["_bh"] == selected_bh]
        .groupby("_month", as_index=False)
        .agg(
            HC=("signed_hc", "sum"),
            PO=("signed_po", "sum"),
            Margin=("signed_margin", "sum"),
        )
    )
    s["sort"] = s["_month"].map(month_sort_key)
    s = s.sort_values("sort")

    x1, x2 = st.columns(2)
    with x1:
        fig = px.line(
            s, x="_month", y=["PO", "Margin"], markers=True,
            title=f"{selected_bh} — Signed Client PO & Margin",
            labels={"value": "₹ Lakhs", "_month": "", "variable": ""},
        )
        st.plotly_chart(fig, use_container_width=True)
    with x2:
        fig = px.line(
            s, x="_month", y="HC", markers=True,
            title=f"{selected_bh} — Signed Client HC",
            labels={"HC": "Headcount", "_month": ""},
        )
        st.plotly_chart(fig, use_container_width=True)
else:
    st.info("Signed Clients Data could not be converted into a monthly trend.")


# -----------------------------
# Contract closure view
# -----------------------------
st.markdown('<div class="section">Contract Closure Exposure</div>', unsafe_allow_html=True)

if not closure.empty:
    closure_view = (
        closure.groupby(["_bh", "_client"], as_index=False)
        .agg(
            HC=("Emp ID", "count"),
            PO_MRR_L=("closure_po", "sum"),
        )
        .sort_values("PO_MRR_L", ascending=False)
    )
    st.dataframe(
        closure_view.rename(columns={"_bh": "BH", "_client": "Client"}),
        use_container_width=True,
        hide_index=True,
        height=300,
        column_config={
            "PO_MRR_L": st.column_config.NumberColumn(format="₹ %.1fL")
        },
    )
else:
    st.info("No contract closure records found for the five selected BHs.")


# -----------------------------
# Data coverage / source health
# -----------------------------
with st.expander("Data Source & Coverage"):
    source_rows = []
    for source, df in data.items():
        source_rows.append({
            "Source": source,
            "Rows": len(df),
            "Latest Month": (
                max(df["_month"].dropna().tolist(), key=month_sort_key)
                if "_month" in df.columns and df["_month"].notna().any()
                else "—"
            ),
        })
    st.dataframe(
        pd.DataFrame(source_rows),
        use_container_width=True,
        hide_index=True,
    )

st.caption(
    "Source: connected Excel workbook. Dashboard is restricted to Sadhna, Mehr, "
    "Prathap, Anuradha and Deepak. Figures are calculated directly from the raw sheets; "
    "no static dashboard values are embedded."
)
