"""
app.py – Streamlit web UI for the Power Query scheduling logic.

Run:  streamlit run app.py
"""

import io
import os
import pandas as pd
import streamlit as st

from pq_logic import (
    DEFAULT_CAPS,
    DEFAULT_COL_QTY,
    DEFAULT_COL_MRP,
    DEFAULT_COL_REQ_DATE,
    DEFAULT_COL_COMMIT_DATE,
    build_cap_flat,
    run_query,
)

WORKSPACE_FILE  = "MPS.xlsx"
DEFAULT_DATA_SHEET = "Entry Open Orders"
DEFAULT_CAP_SHEET  = "Capacity Table"
DEFAULT_CAP_HEADER = 1   # blank row above headers in MPS.xlsx

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="MPS Capacity Scheduler",
    page_icon="🏭",
    layout="wide",
)

st.title("🏭 MPS – Capacity Schedule Tester")

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Settings")

    base_week = st.number_input(
        "Base Week (ISO)", min_value=1, max_value=53, value=24, step=1,
        help="Matches BaseWeek = 24 in the M code.",
    )

    st.subheader("Column names")
    col_qty         = st.text_input("Qty column",            value=DEFAULT_COL_QTY)
    col_mrp         = st.text_input("MRP column",            value=DEFAULT_COL_MRP)
    col_req_date    = st.text_input("Requested date column", value=DEFAULT_COL_REQ_DATE)
    col_commit_date = st.text_input("Plan/Commit date column", value=DEFAULT_COL_COMMIT_DATE,
                                    help="Used first; falls back to Req. date if empty.")

    st.subheader("Default capacities (fallback)")
    for lg in list(DEFAULT_CAPS):
        DEFAULT_CAPS[lg] = st.number_input(lg, value=DEFAULT_CAPS[lg], step=500)

    st.divider()
    st.caption("M code → `Excel.xlsx`")

# ── Data source selector ──────────────────────────────────────────────────────
st.subheader("1 · Data source")

src_option = st.radio(
    "Choose data source",
    options=["📂 Use MPS.xlsx from workspace", "⬆️ Upload a different file"],
    horizontal=True,
)

use_workspace = src_option.startswith("📂")

if use_workspace:
    if not os.path.exists(WORKSPACE_FILE):
        st.error(f"`{WORKSPACE_FILE}` not found in the workspace. Upload a file instead.")
        st.stop()
    file_source = WORKSPACE_FILE
    xl_bytes    = open(WORKSPACE_FILE, "rb").read()
    st.success(f"Using **{WORKSPACE_FILE}** ({os.path.getsize(WORKSPACE_FILE):,} bytes) from workspace.")
else:
    uploaded = st.file_uploader("Upload an Excel file (.xlsx)", type=["xlsx"])
    if uploaded is None:
        st.info("Upload an Excel file to continue.")
        st.stop()
    file_source = uploaded.name
    xl_bytes    = uploaded.read()
    st.success(f"Uploaded: **{uploaded.name}**")

# ── Sheet selection ───────────────────────────────────────────────────────────
st.subheader("2 · Sheet selection")

xl          = pd.ExcelFile(io.BytesIO(xl_bytes))
sheet_names = xl.sheet_names

col_a, col_b, col_c = st.columns(3)

with col_a:
    def _idx(names, *candidates):
        for c in candidates:
            if c in names:
                return names.index(c)
        return 0

    data_sheet = st.selectbox(
        "Data sheet", sheet_names,
        index=_idx(sheet_names, DEFAULT_DATA_SHEET, "Data"),
    )

with col_b:
    cap_options   = ["(none)"] + sheet_names
    cap_sheet     = st.selectbox(
        "Capacity sheet", cap_options,
        index=_idx(cap_options, DEFAULT_CAP_SHEET, "Capacity"),
    )

with col_c:
    cap_header = st.number_input(
        "Capacity header row", min_value=0, max_value=10,
        value=DEFAULT_CAP_HEADER,
        help="0 = first row is header. Set to 1 if there's a blank row above the headers.",
    )

# ── Load and preview data ─────────────────────────────────────────────────────
buf = io.BytesIO(xl_bytes)
data_df = pd.read_excel(buf, sheet_name=data_sheet)

buf = io.BytesIO(xl_bytes)
cap_df = (
    pd.read_excel(buf, sheet_name=cap_sheet, header=int(cap_header))
    if cap_sheet != "(none)"
    else None
)

with st.expander(f"📋 Preview — {data_sheet}  ({len(data_df):,} rows)", expanded=False):
    st.dataframe(data_df.head(100), use_container_width=True)

if cap_df is not None:
    with st.expander(f"📋 Preview — {cap_sheet}  ({len(cap_df):,} rows)", expanded=False):
        st.dataframe(cap_df, use_container_width=True)

    with st.expander("🔍 Parsed capacity buckets", expanded=False):
        cap_flat = build_cap_flat(cap_df)
        if cap_flat:
            cdf = pd.DataFrame(cap_flat)
            st.dataframe(
                cdf.pivot(index="_LG", columns="_WkNum", values="_Cap").fillna(0),
                use_container_width=True,
            )
        else:
            st.warning("No capacity rows parsed. Check that 'Line Group' column and date headers exist.")

# ── Run ───────────────────────────────────────────────────────────────────────
st.subheader("3 · Run")

if st.button("▶ Run Query", type="primary", use_container_width=True):
    with st.spinner("Running scheduling logic…"):
        try:
            result = run_query(
                data_df,
                cap_df,
                base_week=int(base_week),
                col_qty=col_qty,
                col_mrp=col_mrp,
                col_req_date=col_req_date,
                col_commit_date=col_commit_date,
            )
        except Exception as e:
            st.error(f"Error: {e}")
            st.exception(e)
            st.stop()

    # ── Metrics ───────────────────────────────────────────────────────────────
    st.success(f"✅ Done — {len(result):,} rows")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Output rows",      f"{len(result):,}")
    m2.metric("Split orders",     f"{(result['SplitFlag']=='SPLIT').sum():,}")
    m3.metric("Groups",           result['LineGroup'].nunique())
    m4.metric("Total Sched. Qty", f"{result['ScheduledQty'].sum():,.0f}")

    # ── Summary pivot ─────────────────────────────────────────────────────────
    with st.expander("📊 Summary — LineGroup × Production Week", expanded=True):
        summ = (
            result.groupby(["LineGroup", "ProductionWeek"], sort=True)
            .agg(
                Orders          = ("ScheduledQty", "count"),
                ScheduledQty    = ("ScheduledQty", "sum"),
                Splits          = ("SplitFlag", lambda s: (s == "SPLIT").sum()),
            )
            .reset_index()
        )
        # Pivot for compact view
        pivot = summ.pivot_table(
            index="LineGroup", columns="ProductionWeek",
            values="ScheduledQty", aggfunc="sum", fill_value=0,
        )
        st.dataframe(pivot, use_container_width=True)
        with st.expander("Detail table"):
            st.dataframe(summ, use_container_width=True)

    # ── Full result ───────────────────────────────────────────────────────────
    with st.expander("📄 Full result", expanded=False):
        def _highlight(row):
            if row.get("SplitFlag") == "SPLIT":
                return ["background-color: #fff3cd"] * len(row)
            return [""] * len(row)
        st.dataframe(result.style.apply(_highlight, axis=1),
                     use_container_width=True, height=420)

    # ── Download ──────────────────────────────────────────────────────────────
    st.subheader("4 · Download")
    c1, c2 = st.columns(2)

    buf_xl = io.BytesIO()
    with pd.ExcelWriter(buf_xl, engine="openpyxl") as w:
        result.to_excel(w, sheet_name="Scheduled", index=False)
        summ.to_excel(w,   sheet_name="Summary",   index=False)
        pivot.reset_index().to_excel(w, sheet_name="Pivot", index=False)

    c1.download_button(
        "⬇️ Excel (.xlsx)",
        data=buf_xl.getvalue(),
        file_name="scheduled_result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    buf_csv = io.StringIO()
    result.to_csv(buf_csv, index=False)
    c2.download_button(
        "⬇️ CSV",
        data=buf_csv.getvalue(),
        file_name="scheduled_result.csv",
        mime="text/csv",
        use_container_width=True,
    )

# ── M code viewer ─────────────────────────────────────────────────────────────
with st.expander("📝 Current M code (Excel.xlsx)", expanded=False):
    try:
        st.code(open("Excel.xlsx").read(), language="plaintext")
    except Exception:
        st.warning("Could not read Excel.xlsx.")
