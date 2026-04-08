import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="AWP Availability Dashboard",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =========================
# Custom CSS
# =========================
st.markdown("""
<style>
    .main {
        padding-top: 1.2rem;
    }

    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
        padding-left: 2rem;
        padding-right: 2rem;
    }

    .dashboard-title {
        font-size: 2.4rem;
        font-weight: 700;
        margin-bottom: 0.2rem;
    }

    .dashboard-subtitle {
        color: #6b7280;
        font-size: 1rem;
        margin-bottom: 1.2rem;
    }

    .section-title {
        font-size: 1.2rem;
        font-weight: 700;
        margin-top: 1rem;
        margin-bottom: 0.6rem;
    }

    .card {
        border-radius: 16px;
        padding: 18px 20px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.06);
        border: 1px solid #e5e7eb;
        background-color: white;
        margin-bottom: 0.8rem;
    }

    .card-label {
        font-size: 0.95rem;
        color: #6b7280;
        margin-bottom: 0.4rem;
    }

    .card-value {
        font-size: 2rem;
        font-weight: 700;
        color: #111827;
        line-height: 1.1;
    }

    .card-green {
        border-left: 6px solid #16a34a;
    }

    .card-yellow {
        border-left: 6px solid #eab308;
    }

    .card-blue {
        border-left: 6px solid #2563eb;
    }

    .card-red {
        border-left: 6px solid #dc2626;
    }

    .card-dark {
        border-left: 6px solid #0f766e;
    }

    .note-box {
        border-radius: 14px;
        padding: 14px 16px;
        background: #f8fafc;
        border: 1px solid #e5e7eb;
        margin-bottom: 1rem;
    }

    .small-note {
        color: #6b7280;
        font-size: 0.92rem;
    }
</style>
""", unsafe_allow_html=True)

# =========================
# Load Data
# =========================
@st.cache_data
def load_data():
    df = pd.read_excel("data.xlsx", sheet_name="clean_data")
    df.columns = df.columns.str.strip()

    numeric_cols = ["UK Stock", "On Water", "Ordered", "Next Order"]
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    date_cols = ["On Water ETA", "Machines Order Date", "EGRD", "Ordered ETA", "Next Order Date", "Next ETA"]
    for col in date_cols:
        df[col] = pd.to_datetime(df[col], errors="coerce")

    df["Model"] = df["Model"].fillna("").astype(str).str.strip()
    df["Product Group"] = df["Product Group"].fillna("").astype(str).str.strip()

    df = df[df["Model"] != ""]

    df["Total Visible Supply"] = (
        df["UK Stock"] + df["On Water"] + df["Ordered"] + df["Next Order"]
    )

    def get_status(row):
        if row["UK Stock"] > 0:
            return "In Stock"
        elif row["On Water"] > 0:
            return "Incoming"
        elif row["Ordered"] > 0 or row["Next Order"] > 0:
            return "Future Supply"
        else:
            return "No Supply"

    df["Status"] = df.apply(get_status, axis=1)

    return df

df = load_data()

# =========================
# Sidebar
# =========================
st.sidebar.title("Filters")

view_mode = st.sidebar.radio("View Mode", ["Sales View", "Full View"])

group_options = ["All"] + sorted(df["Product Group"].dropna().unique().tolist())
selected_group = st.sidebar.selectbox("Product Group", group_options)

status_options = ["All", "In Stock", "Incoming", "Future Supply", "No Supply"]
selected_status = st.sidebar.selectbox("Status", status_options)

show_only_risk = st.sidebar.checkbox("Show only risk models")

# =========================
# Apply Filters
# =========================
filtered_df = df.copy()

if selected_group != "All":
    filtered_df = filtered_df[filtered_df["Product Group"] == selected_group]

if selected_status != "All":
    filtered_df = filtered_df[filtered_df["Status"] == selected_status]

if show_only_risk:
    filtered_df = filtered_df[filtered_df["Status"].isin(["Incoming", "Future Supply", "No Supply"])]

# =========================
# Header
# =========================
st.markdown('<div class="dashboard-title">AWP Availability Dashboard</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="dashboard-subtitle">Supply visibility for UK stock, on-water units, confirmed orders, and next supply pipeline.</div>',
    unsafe_allow_html=True
)

st.markdown(
    f"""
    <div class="note-box">
        <div class="small-note">
            Current view: <b>{view_mode}</b> &nbsp;&nbsp;|&nbsp;&nbsp;
            Product Group: <b>{selected_group}</b> &nbsp;&nbsp;|&nbsp;&nbsp;
            Status: <b>{selected_status}</b>
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

# =========================
# KPI Cards - Supply
# =========================
st.markdown('<div class="section-title">Supply Overview</div>', unsafe_allow_html=True)

k1, k2, k3, k4 = st.columns(4)

with k1:
    st.markdown(
        f"""
        <div class="card card-dark">
            <div class="card-label">UK Stock</div>
            <div class="card-value">{int(filtered_df["UK Stock"].sum())}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with k2:
    st.markdown(
        f"""
        <div class="card card-blue">
            <div class="card-label">On Water</div>
            <div class="card-value">{int(filtered_df["On Water"].sum())}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with k3:
    st.markdown(
        f"""
        <div class="card card-yellow">
            <div class="card-label">Ordered</div>
            <div class="card-value">{int(filtered_df["Ordered"].sum())}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with k4:
    st.markdown(
        f"""
        <div class="card card-green">
            <div class="card-label">Total Visible Supply</div>
            <div class="card-value">{int(filtered_df["Total Visible Supply"].sum())}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

# =========================
# KPI Cards - Model Status
# =========================
st.markdown('<div class="section-title">Model Status Overview</div>', unsafe_allow_html=True)

in_stock_count = (filtered_df["Status"] == "In Stock").sum()
incoming_count = (filtered_df["Status"] == "Incoming").sum()
future_count = (filtered_df["Status"] == "Future Supply").sum()
no_supply_count = (filtered_df["Status"] == "No Supply").sum()

s1, s2, s3, s4 = st.columns(4)

with s1:
    st.markdown(
        f"""
        <div class="card card-green">
            <div class="card-label">In Stock Models</div>
            <div class="card-value">{int(in_stock_count)}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with s2:
    st.markdown(
        f"""
        <div class="card card-yellow">
            <div class="card-label">Incoming Models</div>
            <div class="card-value">{int(incoming_count)}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with s3:
    st.markdown(
        f"""
        <div class="card card-blue">
            <div class="card-label">Future Supply Models</div>
            <div class="card-value">{int(future_count)}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with s4:
    st.markdown(
        f"""
        <div class="card card-red">
            <div class="card-label">No Supply Models</div>
            <div class="card-value">{int(no_supply_count)}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

# =========================
# Prepare Display Data
# =========================
display_df = filtered_df.copy()

date_cols = ["On Water ETA", "Machines Order Date", "EGRD", "Ordered ETA", "Next Order Date", "Next ETA"]
for col in date_cols:
    display_df[col] = display_df[col].dt.strftime("%d-%b-%y")
    display_df[col] = display_df[col].fillna("")

status_map = {
    "In Stock": "✅ In Stock",
    "Incoming": "🟡 Incoming",
    "Future Supply": "🔵 Future Supply",
    "No Supply": "❌ No Supply",
}
display_df["Status Display"] = display_df["Status"].map(status_map)

# Status priority for sorting
status_priority = {
    "No Supply": 1,
    "Future Supply": 2,
    "Incoming": 3,
    "In Stock": 4,
}

display_df["Status Priority"] = display_df["Status"].map(status_priority)

# Sort main table: risk first, then lower supply first
display_df = display_df.sort_values(
    by=["Status Priority", "Total Visible Supply", "UK Stock"],
    ascending=[True, True, True]
)

# =========================
# Top Attention Models
# =========================
st.markdown('<div class="section-title">Top Attention Models</div>', unsafe_allow_html=True)

attention_df = display_df[display_df["Status"].isin(["No Supply", "Future Supply", "Incoming"])].copy()

if attention_df.empty:
    st.success("No attention models in the current filter.")
else:
    attention_columns = [
        "Model",
        "Product Group",
        "UK Stock",
        "On Water",
        "Ordered",
        "Next Order",
        "Total Visible Supply",
        "Status Display",
    ]
    st.dataframe(
        attention_df[attention_columns].head(8),
        width="stretch",
        hide_index=True
    )

# =========================
# Main Table
# =========================
st.markdown('<div class="section-title">Model Availability</div>', unsafe_allow_html=True)

if view_mode == "Sales View":
    sales_columns = [
        "Model",
        "Product Group",
        "UK Stock",
        "On Water",
        "Ordered",
        "Next Order",
        "Next Order Date",
        "Total Visible Supply",
        "Status Display",
    ]
    st.dataframe(
        display_df[sales_columns],
        width="stretch",
        hide_index=True
    )
else:
    full_columns = [
        "Model",
        "Product Group",
        "UK Stock",
        "On Water",
        "On Water ETA",
        "Ordered",
        "Machines Order Date",
        "EGRD",
        "Ordered ETA",
        "Next Order",
        "Next Order Date",
        "Next ETA",
        "Total Visible Supply",
        "Status Display",
    ]
    st.dataframe(
        display_df[full_columns],
        width="stretch",
        hide_index=True
    )

# =========================
# Risk / Attention Section
# =========================
st.markdown('<div class="section-title">Risk / Attention Models</div>', unsafe_allow_html=True)

risk_df = display_df[display_df["Status"].isin(["Incoming", "Future Supply", "No Supply"])]

if risk_df.empty:
    st.success("No risk models in the current filter.")
else:
    risk_columns = [
        "Model",
        "Product Group",
        "UK Stock",
        "On Water",
        "Ordered",
        "Next Order",
        "Total Visible Supply",
        "Status Display",
    ]
    st.dataframe(
        risk_df[risk_columns],
        width="stretch",
        hide_index=True
    )

# =========================
# Footer Note
# =========================
st.markdown(
    """
    <div class="note-box">
        <div class="small-note">
            Note: Total Visible Supply = UK Stock + On Water + Ordered + Next Order.
            This is a gross supply visibility view and does not deduct rolling sales consumption.
        </div>
    </div>
    """,
    unsafe_allow_html=True
)