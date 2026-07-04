import streamlit as st
import pandas as pd
from pathlib import Path
import plotly.express as px

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="Resource Planning Dashboard", layout="wide")

# =========================
# LOAD FILE
# =========================
BASE_DIR = Path(__file__).resolve().parent
EXCEL_PATH = BASE_DIR / "resource_planning_sample.xlsx"

st.title("📊 Resource Planning Dashboard")

# Debug path
st.write("📁 App folder:", BASE_DIR)
st.write("📄 Excel path:", EXCEL_PATH)
st.write("✅ File exists:", EXCEL_PATH.exists())

if not EXCEL_PATH.exists():
    st.error("❌ Không tìm thấy file Excel trong cùng thư mục với app.py")
    st.stop()

# =========================
# LOAD DATA (CACHE)
# =========================
@st.cache_data
def load_data():
    resource_df = pd.read_excel(EXCEL_PATH, sheet_name="Resource Plan", header=3)
    summary_df = pd.read_excel(EXCEL_PATH, sheet_name="Summary", header=2)
    return resource_df, summary_df

try:
    resource_df, summary_df = load_data()
except Exception as e:
    st.error(f"❌ Lỗi đọc file Excel: {e}")
    st.stop()

# =========================
# CLEAN DATA
# =========================
resource_df = resource_df.dropna(subset=["Employee"]).copy()
resource_df["Monthly Capacity (hrs)"] = pd.to_numeric(resource_df["Monthly Capacity (hrs)"], errors="coerce")
resource_df["Allocation %"] = pd.to_numeric(resource_df["Allocation %"], errors="coerce")

# The sample workbook stores allocation as a percentage, not as direct hrs/status values.
# Derive the missing values so charts and KPIs still render correctly.
resource_df["Allocated Hrs"] = pd.to_numeric(resource_df["Allocated Hrs"], errors="coerce")
resource_df["Allocated Hrs"] = resource_df["Allocated Hrs"].fillna(
    resource_df["Monthly Capacity (hrs)"] * resource_df["Allocation %"]
)

if "Utilization Status" in resource_df.columns:
    resource_df["Utilization Status"] = resource_df["Utilization Status"].fillna(pd.NA)


def derive_status(allocation_pct):
    if pd.isna(allocation_pct):
        return "Unknown"
    if allocation_pct >= 1.0:
        return "Overallocated"
    if allocation_pct >= 0.8:
        return "Near Capacity"
    if allocation_pct >= 0.6:
        return "Healthy"
    return "Bench Risk"

resource_df["Utilization Status"] = resource_df["Utilization Status"].fillna(
    resource_df["Allocation %"].apply(derive_status)
)

# =========================
# DERIVE SUMMARY DATA
# =========================
# The Summary sheet also has mostly empty derived columns.
# Recalculate from the Resource Plan data so the summary table renders.
if not resource_df.empty and "Project" in resource_df.columns:
    project_summary = (
        resource_df.groupby("Project", as_index=False)
        .agg({
            "Employee": "count",
            "Monthly Capacity (hrs)": "sum",
            "Allocated Hrs": "sum",
            "Allocation %": "mean",
        })
        .rename(columns={
            "Employee": "Team Members",
            "Monthly Capacity (hrs)": "Capacity Hrs",
            "Allocation %": "Avg Allocation %",
        })
    )
    project_summary["Available Hrs"] = (
        project_summary["Capacity Hrs"] - project_summary["Allocated Hrs"]
    )
    
    # Merge computed values into summary_df
    if not summary_df.empty and "Project" in summary_df.columns:
        for col in ["Team Members", "Capacity Hrs", "Allocated Hrs", "Avg Allocation %", "Available Hrs"]:
            if col in summary_df.columns:
                summary_df[col] = summary_df["Project"].map(
                    project_summary.set_index("Project")[col]
                )


# =========================
# SIDEBAR FILTER
# =========================
st.sidebar.header("🔍 Filter")

roles = st.sidebar.multiselect(
    "Role",
    options=resource_df["Role"].dropna().unique(),
    default=resource_df["Role"].dropna().unique()
)

projects = st.sidebar.multiselect(
    "Project",
    options=resource_df["Project"].dropna().unique(),
    default=resource_df["Project"].dropna().unique()
)

statuses = st.sidebar.multiselect(
    "Utilization Status",
    options=resource_df["Utilization Status"].dropna().unique(),
    default=resource_df["Utilization Status"].dropna().unique()
)

filtered_df = resource_df[
    (resource_df["Role"].isin(roles)) &
    (resource_df["Project"].isin(projects)) &
    (resource_df["Utilization Status"].isin(statuses))
]

# =========================
# KPI SECTION
# =========================
st.subheader("📌 KPI Overview")

col1, col2, col3, col4 = st.columns(4)

total_capacity = filtered_df["Monthly Capacity (hrs)"].sum()
total_allocated = filtered_df["Allocated Hrs"].sum()

utilization = (
    total_allocated / total_capacity * 100
    if total_capacity > 0 else 0
)

bench_count = (filtered_df["Utilization Status"] == "Bench Risk").sum()

col1.metric("Total Capacity", int(total_capacity) if pd.notnull(total_capacity) else 0)
col2.metric("Allocated Hours", int(total_allocated) if pd.notnull(total_allocated) else 0)
col3.metric("Utilization %", f"{utilization:.1f}%")
col4.metric("Bench Risk", int(bench_count))

# =========================
# CHART 1 - Allocation by Project
# =========================
st.subheader("📊 Allocation by Project")

project_df = (
    filtered_df.groupby("Project")["Allocated Hrs"]
    .sum()
    .reset_index()
)

fig1 = px.bar(
    project_df,
    x="Project",
    y="Allocated Hrs",
    color="Project",
    text_auto=True
)

st.plotly_chart(fig1, use_container_width=True)

# =========================
# CHART 2 - Utilization Status
# =========================
st.subheader("📊 Utilization Status")

status_df = (
    filtered_df["Utilization Status"]
    .value_counts()
    .reset_index()
)

status_df.columns = ["Status", "Count"]

fig2 = px.pie(
    status_df,
    names="Status",
    values="Count"
)

st.plotly_chart(fig2, use_container_width=True)

# =========================
# CHART 3 - Role Distribution
# =========================
st.subheader("📊 Resource by Role")

role_df = (
    filtered_df["Role"]
    .value_counts()
    .reset_index()
)

role_df.columns = ["Role", "Count"]

fig3 = px.bar(
    role_df,
    x="Role",
    y="Count",
    color="Role",
    text_auto=True
)

st.plotly_chart(fig3, use_container_width=True)

# =========================
# TABLE
# =========================
st.subheader("📋 Resource Details")
st.dataframe(filtered_df, use_container_width=True)

# =========================
# SUMMARY
# =========================
st.subheader("📈 Project Summary")
st.dataframe(summary_df, use_container_width=True)
