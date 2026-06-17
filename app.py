"""
app.py – Streamlit web UI for testing the Power Query scheduling logic.

Run with:
    streamlit run app.py
Then open the forwarded port 8501 in your browser (VS Code will show
a "Open in Browser" notification automatically in Codespaces).
"""

import io
import math
import textwrap

import pandas as pd
import streamlit as st

from pq_logic import (
    DEFAULT_CAPS,
    DEFAULT_COL_QTY,
    DEFAULT_COL_MRP,
    DEFAULT_COL_REQ_DATE,
    DEFAULT_COL_COMMIT_DATE,
    build_cap_flat,
    get_buckets,
    run_query,
    normalize_group,
)

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Power Query Tester",
    page_icon="⚙️",
    layout="wide",
)

st.title("⚙️ Power Query – Capacity Schedule Tester")
st.caption(
    "Upload your Excel workbook, pick the sheets, and run the scheduling query "
    "directly here — no need to open Excel."
)

# ── Sidebar: settings ─────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Settings")
    base_week = st.number_input(
        "Base Week (ISO)", min_value=1, max_value=53, value=24, step=1,
        help="The starting week number. Matches BaseWeek = 24 in the M code."
    )
    col_qty         = st.text_input("Order quantity column", value=DEFAULT_COL_QTY)
    col_mrp         = st.text_input("MRP column",            value=DEFAULT_COL_MRP)
    col_req_date    = st.text_input("Req. Date column",       value=DEFAULT_COL_REQ_DATE)
    col_commit_date = st.text_input("Commit/Plan Date column", value=DEFAULT_COL_COMMIT_DATE,
                                    help="Used first; falls back to Req. Date if empty.")

    st.divider()
    st.subheader("Default capacities (fallback)")
    default_cap_inputs = {}
    for lg, cap in DEFAULT_CAPS.items():
        default_cap_inputs[lg] = st.number_input(lg, value=cap, step=500)

    st.divider()
    st.markdown(
        "**M code file:** `Excel.xlsx`  \n"
        "Edit it in VS Code — changes are picked up on the next Run."
    )

# ── File upload ───────────────────────────────────────────────────────────────
st.subheader("1 · Upload your workbook")

uploaded = st.file_uploader(
    "Drop an Excel file here (.xlsx)",
    type=["xlsx"],
    help="The file must contain at least a Data sheet. "
         "Add a Capacity sheet for dynamic weekly capacities.",
)

if uploaded is None:
    st.info(
        "👆 Upload your Excel file to get started.  \n"
        "The workbook should have:\n"
        "- **Data** sheet — orders / demand\n"
        "- **Capacity** sheet *(optional)* — LineGroup + weekly date columns"
    )
    st.stop()

# ── Load sheets ───────────────────────────────────────────────────────────────
xl = pd.ExcelFile(uploaded)
sheet_names = xl.sheet_names

st.subheader("2 · Select sheets")
col_a, col_b = st.columns(2)

with col_a:
    default_data = sheet_names.index("Entry Open Orders") if "Entry Open Orders" in sheet_names else (
        sheet_names.index("Data") if "Data" in sheet_names else 0
    )
    data_sheet = st.selectbox("Data sheet", sheet_names, index=default_data)

with col_b:
    cap_options = ["(none)"] + sheet_names
    default_cap_sheet = (
        cap_options.index("Capacity Table") if "Capacity Table" in cap_options else
        (cap_options.index("Capacity") if "Capacity" in cap_options else 0)
    )
    cap_sheet = st.selectbox("Capacity sheet", cap_options, index=default_cap_sheet)

cap_header_row = st.number_input(
    "Capacity sheet header row (0 = first row)", min_value=0, max_value=10, value=1,
    help="Set to 1 if the Capacity sheet has a blank row above the headers (like MPS.xlsx)."
)

data_df = pd.read_excel(uploaded, sheet_name=data_sheet)
cap_df  = (
    pd.read_excel(uploaded, sheet_name=cap_sheet, header=int(cap_header_row))
    if cap_sheet != "(none)"
    else None
)

# ── Preview input data ────────────────────────────────────────────────────────
with st.expander("📋 Preview: Data sheet", expanded=False):
    st.dataframe(data_df, use_container_width=True)

if cap_df is not None:
    with st.expander("📋 Preview: Capacity sheet", expanded=False):
        st.dataframe(cap_df, use_container_width=True)

    with st.expander("🔍 Parsed capacity buckets (what the query sees)", expanded=False):
        cap_flat = build_cap_flat(cap_df)
        if cap_flat:
            st.dataframe(pd.DataFrame(cap_flat), use_container_width=True)
        else:
            st.warning(
                "No capacity rows could be parsed from the Capacity sheet. "
                "Check that the 'Line Group' column exists and date headers are valid dates."
            )

# ── Run query ─────────────────────────────────────────────────────────────────
st.subheader("3 · Run")

if st.button("▶ Run Query", type="primary", use_container_width=True):

    # Patch DEFAULT_CAPS with sidebar values
    import pq_logic
    for lg, v in default_cap_inputs.items():
        pq_logic.DEFAULT_CAPS[lg] = v

    with st.spinner("Running scheduling logic…"):
        try:
            result = run_query(
                data_df,
                cap_df,
                base_week=base_week,
                col_qty=col_qty,
                col_mrp=col_mrp,
                col_req_date=col_req_date,
                col_commit_date=col_commit_date,
            )
        except Exception as e:
            st.error(f"Error running query: {e}")
            st.exception(e)
            st.stop()

    # ── Results ───────────────────────────────────────────────────────────────
    st.success(f"✅ Done — {len(result):,} rows produced.")

    # Summary metrics
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total rows", f"{len(result):,}")
    m2.metric(
        "Split rows",
        f"{(result['SplitFlag'] == 'SPLIT').sum():,}",
    )
    m3.metric(
        "Rows with Qty > 0",
        f"{(result['ScheduledQty'] > 0).sum():,}",
    )
    m4.metric(
        "Scheduled Qty total",
        f"{result['ScheduledQty'].sum():,.0f}",
    )

    # Per-group summary
    with st.expander("📊 Summary by LineGroup + ProductionWeek", expanded=True):
        summary = (
            result.groupby(["LineGroup", "ProductionWeek"], sort=True)
            .agg(
                Orders=("ScheduledQty", "count"),
                ScheduledQty=("ScheduledQty", "sum"),
                Split_Rows=("SplitFlag", lambda s: (s == "SPLIT").sum()),
            )
            .reset_index()
        )
        st.dataframe(summary, use_container_width=True)

    # Full results
    with st.expander("📄 Full result table", expanded=True):
        # Highlight split rows
        def highlight_split(row):
            if row.get("SplitFlag") == "SPLIT":
                return ["background-color: #fff3cd"] * len(row)
            return [""] * len(row)

        styled = result.style.apply(highlight_split, axis=1)
        st.dataframe(styled, use_container_width=True, height=450)

    # ── Download ──────────────────────────────────────────────────────────────
    st.subheader("4 · Download result")
    col_dl1, col_dl2 = st.columns(2)

    with col_dl1:
        buf_xlsx = io.BytesIO()
        with pd.ExcelWriter(buf_xlsx, engine="openpyxl") as writer:
            result.to_excel(writer, index=False, sheet_name="Result")
            summary.to_excel(writer, index=False, sheet_name="Summary")
        st.download_button(
            "⬇️ Download as Excel (.xlsx)",
            data=buf_xlsx.getvalue(),
            file_name="scheduled_result.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with col_dl2:
        buf_csv = io.StringIO()
        result.to_csv(buf_csv, index=False)
        st.download_button(
            "⬇️ Download as CSV",
            data=buf_csv.getvalue(),
            file_name="scheduled_result.csv",
            mime="text/csv",
            use_container_width=True,
        )

# ── Footer: show current M code ───────────────────────────────────────────────
with st.expander("📝 Current Power Query M code (Excel.xlsx)", expanded=False):
    try:
        with open("Excel.xlsx", "r") as f:
            mcode = f.read()
        st.code(mcode, language="plaintext")
    except Exception:
        st.warning("Could not read Excel.xlsx M code file.")
