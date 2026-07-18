import hashlib
import hmac
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st


APP_DIR = Path(__file__).parent
EXCEL_PATH = APP_DIR / "data.xlsx"
TABLE_NAME = "awp_inventory"

COLUMNS = [
    "Model",
    "Material Number",
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
    "Remark",
]
NUMERIC_COLUMNS = ["UK Stock", "On Water", "Ordered", "Next Order"]
DATE_COLUMNS = [
    "On Water ETA",
    "Machines Order Date",
    "EGRD",
    "Ordered ETA",
    "Next Order Date",
    "Next ETA",
]
DB_TO_APP = {
    "id": "_id",
    "model": "Model",
    "material_number": "Material Number",
    "product_group": "Product Group",
    "uk_stock": "UK Stock",
    "on_water": "On Water",
    "on_water_eta": "On Water ETA",
    "ordered": "Ordered",
    "machines_order_date": "Machines Order Date",
    "egrd": "EGRD",
    "ordered_eta": "Ordered ETA",
    "next_order": "Next Order",
    "next_order_date": "Next Order Date",
    "next_eta": "Next ETA",
    "remark": "Remark",
    "sort_order": "_sort_order",
}
APP_TO_DB = {value: key for key, value in DB_TO_APP.items()}


st.set_page_config(
    page_title="AWP Availability Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .block-container {padding-top: 1.5rem; padding-bottom: 2rem;}
    .dashboard-title {font-size: 2.35rem; font-weight: 750; margin-bottom: .15rem;}
    .dashboard-subtitle {color: #6b7280; margin-bottom: 1.2rem;}
    .section-title {font-size: 1.2rem; font-weight: 700; margin: 1rem 0 .6rem;}
    .card {border-radius: 16px; padding: 18px 20px; border: 1px solid #e5e7eb;
           background: white; box-shadow: 0 2px 10px rgba(0,0,0,.06);}
    .card-label {font-size: .95rem; color: #6b7280; margin-bottom: .35rem;}
    .card-value {font-size: 2rem; font-weight: 750; color: #111827;}
    .card-green {border-left: 6px solid #16a34a;}
    .card-yellow {border-left: 6px solid #eab308;}
    .card-blue {border-left: 6px solid #2563eb;}
    .card-red {border-left: 6px solid #dc2626;}
    .card-dark {border-left: 6px solid #0f766e;}
    .note-box {border-radius: 14px; padding: 14px 16px; background: #f8fafc;
               border: 1px solid #e5e7eb; margin-bottom: 1rem;}
    .small-note {color: #6b7280; font-size: .92rem;}
    </style>
    """,
    unsafe_allow_html=True,
)


def get_secret(name, default=None):
    try:
        return st.secrets.get(name, default)
    except Exception:
        return default


def cloud_is_configured():
    return bool(get_secret("SUPABASE_URL") and get_secret("SUPABASE_KEY"))


@st.cache_resource
def get_supabase_client():
    from supabase import create_client

    return create_client(get_secret("SUPABASE_URL"), get_secret("SUPABASE_KEY"))


def normalise_dataframe(df):
    df = df.copy()
    for column in COLUMNS:
        if column not in df.columns:
            df[column] = None
    for column in ["Model", "Material Number", "Product Group", "Remark"]:
        df[column] = df[column].fillna("").astype(str).str.strip()
    for column in NUMERIC_COLUMNS:
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0).clip(lower=0).astype(int)
    for column in DATE_COLUMNS:
        df[column] = pd.to_datetime(df[column], errors="coerce")
    df = df[df["Model"] != ""].reset_index(drop=True)
    return df


def load_local_data():
    return normalise_dataframe(pd.read_excel(EXCEL_PATH, sheet_name=0))


def load_cloud_data():
    response = (
        get_supabase_client()
        .table(TABLE_NAME)
        .select("*")
        .order("sort_order")
        .execute()
    )
    df = pd.DataFrame(response.data)
    if df.empty:
        return pd.DataFrame(columns=COLUMNS + ["_id", "_sort_order"])
    return normalise_dataframe(df.rename(columns=DB_TO_APP))


@st.cache_data(ttl=30, show_spinner=False)
def load_data():
    if cloud_is_configured():
        return load_cloud_data()
    return load_local_data()


def dataframe_to_excel(df):
    output = BytesIO()
    export_df = df[COLUMNS].copy()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export_df.to_excel(writer, sheet_name="clean_data", index=False)
    output.seek(0)
    return output.getvalue()


def app_row_to_db(row, sort_order):
    record = {}
    for column in COLUMNS:
        value = row.get(column)
        key = APP_TO_DB[column]
        if column in DATE_COLUMNS:
            record[key] = None if pd.isna(value) else pd.Timestamp(value).date().isoformat()
        elif column in NUMERIC_COLUMNS:
            record[key] = int(0 if pd.isna(value) else value)
        else:
            record[key] = "" if pd.isna(value) else str(value).strip()
    record["sort_order"] = sort_order
    row_id = row.get("_id")
    if pd.notna(row_id) and str(row_id).strip():
        record["id"] = int(row_id)
    return record


def save_cloud_data(edited_df, original_df):
    client = get_supabase_client()
    original_ids = {
        int(value)
        for value in original_df.get("_id", pd.Series(dtype=float)).dropna().tolist()
    }
    kept_ids = {
        int(value)
        for value in edited_df.get("_id", pd.Series(dtype=float)).dropna().tolist()
    }
    deleted_ids = original_ids - kept_ids
    for row_id in deleted_ids:
        client.table(TABLE_NAME).delete().eq("id", row_id).execute()

    existing_records = []
    new_records = []
    for index, (_, row) in enumerate(edited_df.iterrows()):
        record = app_row_to_db(row, index)
        if "id" in record:
            existing_records.append(record)
        else:
            new_records.append(record)
    if existing_records:
        client.table(TABLE_NAME).upsert(existing_records).execute()
    if new_records:
        client.table(TABLE_NAME).insert(new_records).execute()


def save_local_data(edited_df):
    # Replace only the data sheet so any other workbook sheets are preserved.
    with pd.ExcelWriter(
        EXCEL_PATH,
        engine="openpyxl",
        mode="a",
        if_sheet_exists="replace",
    ) as writer:
        edited_df[COLUMNS].to_excel(writer, sheet_name="clean_data", index=False)


def replace_all_data(new_df, current_df):
    new_df = normalise_dataframe(new_df)
    if cloud_is_configured():
        client = get_supabase_client()
        current_ids = current_df.get("_id", pd.Series(dtype=float)).dropna().tolist()
        for row_id in current_ids:
            client.table(TABLE_NAME).delete().eq("id", int(row_id)).execute()
        records = [app_row_to_db(row, index) for index, (_, row) in enumerate(new_df.iterrows())]
        for record in records:
            record.pop("id", None)
        if records:
            client.table(TABLE_NAME).insert(records).execute()
    else:
        save_local_data(new_df)


def add_derived_columns(df):
    df = normalise_dataframe(df)
    df["Total Visible Supply"] = df[NUMERIC_COLUMNS].sum(axis=1)

    def status(row):
        if row["UK Stock"] > 0:
            return "In Stock"
        if row["On Water"] > 0:
            return "Incoming"
        if row["Ordered"] > 0 or row["Next Order"] > 0:
            return "Future Supply"
        return "No Supply"

    df["Status"] = df.apply(status, axis=1)
    df["Status Display"] = df["Status"].map(
        {
            "In Stock": "✅ In Stock",
            "Incoming": "🟡 Incoming",
            "Future Supply": "🔵 Future Supply",
            "No Supply": "🔴 No Supply",
        }
    )
    return df


def metric_card(column, label, value, colour):
    with column:
        st.markdown(
            f'<div class="card {colour}"><div class="card-label">{label}</div>'
            f'<div class="card-value">{int(value)}</div></div>',
            unsafe_allow_html=True,
        )


def render_dashboard(raw_df):
    df = add_derived_columns(raw_df)
    st.markdown('<div class="dashboard-title">AWP Availability Dashboard</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="dashboard-subtitle">UK stock and incoming supply visibility.</div>',
        unsafe_allow_html=True,
    )

    with st.sidebar:
        st.markdown("### Filters")
        view_mode = st.radio("View Mode", ["Sales View", "Full View"])
        groups = ["All"] + sorted(item for item in df["Product Group"].unique() if item)
        selected_group = st.selectbox("Product Group", groups)
        statuses = ["All", "In Stock", "Incoming", "Future Supply", "No Supply"]
        selected_status = st.selectbox("Status", statuses)
        only_risk = st.checkbox("Show only risk models")

    filtered = df.copy()
    if selected_group != "All":
        filtered = filtered[filtered["Product Group"] == selected_group]
    if selected_status != "All":
        filtered = filtered[filtered["Status"] == selected_status]
    if only_risk:
        filtered = filtered[filtered["Status"] != "In Stock"]

    st.markdown('<div class="section-title">Supply Overview</div>', unsafe_allow_html=True)
    cols = st.columns(4)
    metric_card(cols[0], "UK Stock", filtered["UK Stock"].sum(), "card-dark")
    metric_card(cols[1], "On Water", filtered["On Water"].sum(), "card-blue")
    metric_card(cols[2], "Ordered", filtered["Ordered"].sum(), "card-yellow")
    metric_card(cols[3], "Total Visible Supply", filtered["Total Visible Supply"].sum(), "card-green")

    st.markdown('<div class="section-title">Model Status Overview</div>', unsafe_allow_html=True)
    cols = st.columns(4)
    metric_card(cols[0], "In Stock Models", (filtered["Status"] == "In Stock").sum(), "card-green")
    metric_card(cols[1], "Incoming Models", (filtered["Status"] == "Incoming").sum(), "card-yellow")
    metric_card(cols[2], "Future Supply Models", (filtered["Status"] == "Future Supply").sum(), "card-blue")
    metric_card(cols[3], "No Supply Models", (filtered["Status"] == "No Supply").sum(), "card-red")

    display = filtered.copy()
    priority = {"No Supply": 1, "Future Supply": 2, "Incoming": 3, "In Stock": 4}
    display["_priority"] = display["Status"].map(priority)
    display = display.sort_values(["_priority", "Total Visible Supply", "UK Stock"])
    for column in DATE_COLUMNS:
        display[column] = display[column].dt.strftime("%d-%b-%y").fillna("")

    attention = display[display["Status"] != "In Stock"]
    st.markdown('<div class="section-title">Top Attention Models</div>', unsafe_allow_html=True)
    attention_columns = [
        "Model", "Material Number", "Product Group", "UK Stock", "On Water",
        "Ordered", "Next Order", "Total Visible Supply", "Status Display", "Remark",
    ]
    if attention.empty:
        st.success("No attention models in the current filter.")
    else:
        st.dataframe(attention[attention_columns].head(8), width="stretch", hide_index=True)

    st.markdown('<div class="section-title">Model Availability</div>', unsafe_allow_html=True)
    if view_mode == "Sales View":
        shown_columns = [
            "Model", "Material Number", "Product Group", "UK Stock", "On Water",
            "Ordered", "Next Order", "Next Order Date", "Total Visible Supply",
            "Status Display", "Remark",
        ]
    else:
        shown_columns = [
            "Model", "Material Number", "Product Group", "UK Stock", "On Water",
            "On Water ETA", "Ordered", "Machines Order Date", "EGRD", "Ordered ETA",
            "Next Order", "Next Order Date", "Next ETA", "Total Visible Supply",
            "Status Display", "Remark",
        ]
    st.dataframe(display[shown_columns], width="stretch", hide_index=True)


def password_matches(password):
    configured_hash = get_secret("ADMIN_PASSWORD_SHA256")
    configured_password = get_secret("ADMIN_PASSWORD")
    if configured_hash:
        actual = hashlib.sha256(password.encode("utf-8")).hexdigest()
        return hmac.compare_digest(actual, configured_hash)
    if configured_password:
        return hmac.compare_digest(password, configured_password)
    return False


def render_login():
    st.title("Data Management")
    if not get_secret("ADMIN_PASSWORD") and not get_secret("ADMIN_PASSWORD_SHA256"):
        st.warning(
            "Admin password is not configured. Add ADMIN_PASSWORD to .streamlit/secrets.toml "
            "before enabling online editing."
        )
        return False
    with st.form("admin_login"):
        password = st.text_input("Admin password", type="password")
        submitted = st.form_submit_button("Log in", type="primary")
    if submitted:
        if password_matches(password):
            st.session_state.admin_authenticated = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    return False


def render_data_management(raw_df):
    if not st.session_state.get("admin_authenticated"):
        return render_login()

    title_col, logout_col = st.columns([6, 1])
    with title_col:
        st.title("Data Management")
    with logout_col:
        if st.button("Log out"):
            st.session_state.admin_authenticated = False
            st.rerun()

    storage_label = "Supabase cloud database" if cloud_is_configured() else "local data.xlsx"
    st.info(f"Current storage: **{storage_label}**")
    st.caption("Edit cells directly. Add a row at the bottom, or select rows and press Delete. Click Save changes when finished.")

    editor_df = raw_df.copy()
    editor_columns = COLUMNS.copy()
    if "_id" in editor_df.columns:
        editor_columns = ["_id"] + editor_columns

    column_config = {
        "_id": st.column_config.NumberColumn("ID", disabled=True),
        "Model": st.column_config.TextColumn("Model *", required=True),
        "Material Number": st.column_config.TextColumn("Material Number"),
        "Product Group": st.column_config.TextColumn("Product Group"),
        "UK Stock": st.column_config.NumberColumn("UK Stock", min_value=0, step=1),
        "On Water": st.column_config.NumberColumn("On Water", min_value=0, step=1),
        "Ordered": st.column_config.NumberColumn("Ordered", min_value=0, step=1),
        "Next Order": st.column_config.NumberColumn("Next Order", min_value=0, step=1),
    }
    for column in DATE_COLUMNS:
        column_config[column] = st.column_config.DateColumn(column, format="DD-MMM-YYYY")

    edited = st.data_editor(
        editor_df[editor_columns],
        column_config=column_config,
        num_rows="dynamic",
        width="stretch",
        hide_index=True,
        key="inventory_editor",
    )

    save_col, download_col, refresh_col = st.columns([1, 1, 4])
    with save_col:
        if st.button("Save changes", type="primary"):
            cleaned = normalise_dataframe(edited)
            if cleaned["Model"].duplicated().any():
                st.error("Model names must be unique. Please remove the duplicate model.")
            else:
                with st.spinner("Saving..."):
                    if cloud_is_configured():
                        save_cloud_data(cleaned, raw_df)
                    else:
                        save_local_data(cleaned)
                    load_data.clear()
                st.success("Changes saved.")
                st.rerun()
    with download_col:
        st.download_button(
            "Download Excel",
            dataframe_to_excel(raw_df),
            file_name="awp_inventory.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with refresh_col:
        if st.button("Discard / refresh"):
            load_data.clear()
            st.rerun()

    st.divider()
    st.subheader("Replace data from Excel")
    st.caption("The first worksheet must contain the same column headings. This replaces all current records.")
    uploaded = st.file_uploader("Choose an .xlsx file", type=["xlsx"])
    confirm = st.checkbox("I understand this will replace all current records")
    if uploaded and st.button("Import and replace", disabled=not confirm):
        imported = pd.read_excel(uploaded, sheet_name=0)
        missing = [column for column in COLUMNS if column not in imported.columns]
        if missing:
            st.error("Missing columns: " + ", ".join(missing))
        else:
            with st.spinner("Importing..."):
                replace_all_data(imported, raw_df)
                load_data.clear()
            st.success("Excel data imported.")
            st.rerun()


page = st.sidebar.radio("Navigation", ["Dashboard", "Data Management"])

try:
    data = load_data()
except Exception as exc:
    st.error("Unable to load inventory data.")
    st.exception(exc)
    st.stop()

if page == "Dashboard":
    render_dashboard(data)
else:
    render_data_management(data)
