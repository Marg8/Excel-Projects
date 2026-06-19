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
    DEFAULT_COL_LINE,
    DEFAULT_COL_REQ_DATE,
    DEFAULT_COL_COMMIT_DATE,
    DEFAULT_COL_STD_PACK,
    apply_capacity_overrides,
    build_cap_flat,
    build_hrs_lookup,
    build_cost_lookup,
    build_std_pack_lookup,
    build_ma3_summary,
    capacity_comparison,
    data_quality_report,
    run_query,
)
from bom_analysis import (
    build_cost_lookup_dated,
    cost_date_summary,
    parse_bom,
    explode_bom_demand,
    component_demand_pivot,
    compute_build_capability,
    build_capability_summary,
    shortage_report,
    run_build_analysis,
    build_rm_coverage_table,
    rm_coverage_to_excel,
    build_line_output_plan,
    line_output_to_excel,
)

WORKSPACE_FILE  = "MPS.xlsx"
DEFAULT_DATA_SHEET = "Entry Open Orders"
DEFAULT_CAP_SHEET  = "Capacity Table"
DEFAULT_CAP_HEADER = 1   # blank row above headers in MPS.xlsx
DEFAULT_HRS_SHEET  = "Hrs"
DEFAULT_COST_SHEET = "Cost"
DEFAULT_SP_SHEET   = "Std Pack"
DEFAULT_BOM_SHEET  = "BOM"
DEFAULT_STOCK_SHEET = "Stock"
DEFAULT_RM_PO_SHEET = "RM_PO"

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="MPS Capacity Scheduler",
    page_icon="🏭",
    layout="wide",
)

# ── Global CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
a { text-decoration: none; color: #464feb; }
tr th, tr td { border: 1px solid #e6e6e6; }
tr th { background-color: #f5f5f5; }
[data-testid="metric-container"] {
    background: #f8f9ff;
    border: 1px solid #e0e4ff;
    border-radius: 8px;
    padding: 8px 12px;
}
[data-testid="metric-container"] label { color: #555; font-size: 12px; }
.section-header {
    background: linear-gradient(90deg, #464feb15, transparent);
    border-left: 4px solid #464feb;
    padding: 6px 14px;
    border-radius: 0 6px 6px 0;
    margin-bottom: 12px;
}
</style>
""", unsafe_allow_html=True)

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
    col_line        = st.text_input("Line column",           value=DEFAULT_COL_LINE,
                                    help="Primary planning key. Example: HC7, H2J, HC4_A9R")
    col_req_date    = st.text_input("Requested date column", value=DEFAULT_COL_REQ_DATE)
    col_commit_date = st.text_input("Plan/Commit date column", value=DEFAULT_COL_COMMIT_DATE,
                                    help="Used first; falls back to Req. date if empty.")
    col_std_pack    = st.text_input("Std Pack column",       value=DEFAULT_COL_STD_PACK)

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

# Master-table sheets
col_d, col_e, col_f = st.columns(3)
with col_d:
    hrs_options = ["(none)"] + sheet_names
    hrs_sheet = st.selectbox(
        "Hrs sheet", hrs_options,
        index=_idx(hrs_options, DEFAULT_HRS_SHEET, "Hrs", "Hours"),
    )
with col_e:
    cost_options = ["(none)"] + sheet_names
    cost_sheet = st.selectbox(
        "Cost sheet", cost_options,
        index=_idx(cost_options, DEFAULT_COST_SHEET, "Cost"),
    )
with col_f:
    sp_options = ["(none)"] + sheet_names
    sp_master_sheet = st.selectbox(
        "Std Pack master sheet", sp_options,
        index=_idx(sp_options, DEFAULT_SP_SHEET, "Std Pack", "StdPack"),
    )

col_g, _, _ = st.columns(3)
with col_g:
    bom_options = ["(none)"] + sheet_names
    bom_sheet = st.selectbox(
        "BOM sheet", bom_options,
        index=_idx(bom_options, DEFAULT_BOM_SHEET, "BOM", "Bill of Materials"),
    )

col_h, col_i, _ = st.columns(3)
with col_h:
    stock_options = ["(none)"] + sheet_names
    stock_sheet = st.selectbox(
        "Stock sheet", stock_options,
        index=_idx(stock_options, DEFAULT_STOCK_SHEET, "Stock"),
    )
with col_i:
    rm_po_options = ["(none)"] + sheet_names
    rm_po_sheet = st.selectbox(
        "RM_PO sheet", rm_po_options,
        index=_idx(rm_po_options, DEFAULT_RM_PO_SHEET, "RM_PO", "RM PO"),
    )

capacity_line_col = "Line"

# ── Load and preview data ─────────────────────────────────────────────────────
buf = io.BytesIO(xl_bytes)
data_df = pd.read_excel(buf, sheet_name=data_sheet)

buf = io.BytesIO(xl_bytes)
cap_df = (
    pd.read_excel(buf, sheet_name=cap_sheet, header=int(cap_header))
    if cap_sheet != "(none)"
    else None
)

buf = io.BytesIO(xl_bytes)
hrs_df = pd.read_excel(buf, sheet_name=hrs_sheet) if hrs_sheet != "(none)" else None

buf = io.BytesIO(xl_bytes)
cost_df = pd.read_excel(buf, sheet_name=cost_sheet) if cost_sheet != "(none)" else None

buf = io.BytesIO(xl_bytes)
sp_master_df = pd.read_excel(buf, sheet_name=sp_master_sheet) if sp_master_sheet != "(none)" else None

buf = io.BytesIO(xl_bytes)
bom_df = pd.read_excel(buf, sheet_name=bom_sheet) if bom_sheet != "(none)" else None

buf = io.BytesIO(xl_bytes)
stock_df = pd.read_excel(buf, sheet_name=stock_sheet) if stock_sheet != "(none)" else None

buf = io.BytesIO(xl_bytes)
rm_po_df = pd.read_excel(buf, sheet_name=rm_po_sheet) if rm_po_sheet != "(none)" else None

with st.expander(f"📋 Preview — {data_sheet}  ({len(data_df):,} rows)", expanded=False):
    st.dataframe(data_df.head(100), use_container_width=True)

if cap_df is not None:
    with st.expander(f"📋 Preview — {cap_sheet}  ({len(cap_df):,} rows)", expanded=False):
        st.dataframe(cap_df, use_container_width=True)

    cap_key_options = [str(c) for c in cap_df.columns]
    if cap_key_options:
        default_key_idx = cap_key_options.index("Line") if "Line" in cap_key_options else 0
        capacity_line_col = st.selectbox(
            "Capacity key column (must match order Line values)",
            cap_key_options,
            index=default_key_idx,
            help="Choose the capacity column that contains the same line codes as your order Line column.",
        )

    line_values_hint = set(data_df[col_line].astype(str).str.strip().str.upper().tolist()) if col_line in data_df.columns else None
    base_cap_flat = build_cap_flat(
        cap_df,
        preferred_line_col=capacity_line_col,
        valid_lines=line_values_hint,
    )

    with st.expander("🔍 Parsed capacity buckets", expanded=False):
        if base_cap_flat:
            cdf = pd.DataFrame(base_cap_flat)
            if "_LineName" in cdf.columns:
                cdf["LineLabel"] = cdf["_Line"] + " | " + cdf["_LineName"].astype(str)
            else:
                cdf["LineLabel"] = cdf["_Line"]
            st.dataframe(
                cdf.pivot(index="LineLabel", columns="_WeekDate", values="_Cap").fillna(0),
                use_container_width=True,
            )
        else:
            st.warning("No capacity rows parsed. Check that 'Line' column and date headers exist.")

    st.subheader("3 · Scenario simulation (horizontal, like Capacity table)")
    st.caption("Edit capacities directly in a Line x Week matrix. Only modified cells are treated as overrides.")

    if base_cap_flat:
        cdf = pd.DataFrame(base_cap_flat)
        line_names = cdf.groupby("_Line")["_LineName"].first().to_dict() if "_LineName" in cdf.columns else {}
        base_wide = (
            cdf.pivot_table(index="_Line", columns="_WeekDate", values="_Cap", aggfunc="sum", fill_value=0)
            .sort_index(axis=0)
            .sort_index(axis=1)
        )

        def _week_label(d):
            return f"{d.month}/{d.day}/{d.year}"

        week_map = {d: _week_label(d) for d in base_wide.columns}
        reverse_week_map = {v: k for k, v in week_map.items()}

        matrix_df = base_wide.rename(columns=week_map).reset_index().rename(columns={"_Line": "Line"})
        matrix_df.insert(1, "Line Name", matrix_df["Line"].map(lambda x: line_names.get(x, x)))

        column_cfg = {
            "Line": st.column_config.TextColumn("Line", disabled=True),
            "Line Name": st.column_config.TextColumn("Line Name", disabled=True),
        }
        for c in matrix_df.columns:
            if c not in {"Line", "Line Name"}:
                column_cfg[c] = st.column_config.NumberColumn(c, min_value=0, step=100)

        edited_matrix = st.data_editor(
            matrix_df,
            key="scenario_overrides_horizontal",
            num_rows="fixed",
            use_container_width=True,
            column_config=column_cfg,
            hide_index=True,
        )

        # Build override rows from changed cells only.
        original_long = matrix_df.melt(id_vars=["Line", "Line Name"], var_name="WeekLabel", value_name="Original capacity")
        edited_long = edited_matrix.melt(id_vars=["Line", "Line Name"], var_name="WeekLabel", value_name="Adjusted capacity")

        compare_long = original_long.merge(edited_long, on=["Line", "Line Name", "WeekLabel"], how="inner")
        compare_long["Original capacity"] = compare_long["Original capacity"].fillna(0).astype(float)
        compare_long["Adjusted capacity"] = compare_long["Adjusted capacity"].fillna(0).astype(float)
        compare_long["Week"] = compare_long["WeekLabel"].map(reverse_week_map)
        compare_long["Delta"] = compare_long["Adjusted capacity"] - compare_long["Original capacity"]

        changed_long = compare_long[compare_long["Delta"] != 0].copy()
        sim_df = changed_long[["Line", "Week", "Adjusted capacity"]].rename(
            columns={"Adjusted capacity": "New Capacity"}
        )

        adjusted_cap_flat = apply_capacity_overrides(base_cap_flat, sim_df)
        cmp_df = capacity_comparison(base_cap_flat, adjusted_cap_flat)

        with st.expander("📊 Capacity comparison (Original vs Adjusted vs Delta)", expanded=False):
            st.dataframe(changed_long[["Line", "Line Name", "Week", "Original capacity", "Adjusted capacity", "Delta"]], use_container_width=True)

            adjusted_pivot = (
                pd.DataFrame(adjusted_cap_flat)
                .pivot_table(index="_Line", columns="_WeekDate", values="_Cap", aggfunc="sum", fill_value=0)
                .sort_index(axis=0)
                .sort_index(axis=1)
            )
            adjusted_pivot = adjusted_pivot.reset_index().rename(columns={"_Line": "Line"})
            adjusted_pivot.insert(1, "Line Name", adjusted_pivot["Line"].map(lambda x: line_names.get(x, x)))
            adjusted_pivot = adjusted_pivot.set_index(["Line", "Line Name"])
            adjusted_pivot = adjusted_pivot.rename(columns=week_map)
            st.caption("Adjusted capacity table")
            st.dataframe(adjusted_pivot, use_container_width=True)
    else:
        sim_df = pd.DataFrame(columns=["Line", "Week", "New Capacity"])
        cmp_df = pd.DataFrame(columns=["Line", "Week", "Original capacity", "Adjusted capacity", "Delta"])
else:
    base_cap_flat = []
    sim_df = pd.DataFrame(columns=["Line", "Week", "New Capacity"])
    cmp_df = pd.DataFrame(columns=["Line", "Week", "Original capacity", "Adjusted capacity", "Delta"])

# ── Data Quality Check ────────────────────────────────────────────────────────
st.subheader("3b · Data Quality Check")
dq = data_quality_report(
    data_df,
    capacity_df=cap_df,
    sp_df=sp_master_df,
    col_qty=col_qty,
    col_line=col_line,
    col_std_pack=col_std_pack,
    col_part="Product code",
    capacity_line_col=capacity_line_col,
)
if dq["issues"]:
    for msg in dq["issues"]:
        if msg.startswith("🔴"):
            st.error(msg)
        elif msg.startswith("⚠️"):
            st.warning(msg)
        else:
            st.info(msg)
else:
    st.success("✅ No data quality issues found — file is compatible.")

with st.expander("🔍 Data quality detail", expanded=False):
    col_dq1, col_dq2, col_dq3 = st.columns(3)
    col_dq1.metric("Bad capacity date headers", len(dq.get("bad_cap_date_headers", [])))
    col_dq2.metric("Rows with no Std Pack",     dq.get("null_std_pack_rows", 0),
                   help="Will be scheduled with Std Pack = 1 (exact qty)")
    col_dq3.metric("Unmatched MRP codes",        len(dq.get("unmatched_mrp_codes", [])),
                   help="Will use median fallback capacity")
    if dq.get("repaired_cap_dates"):
        st.caption("Auto-repaired capacity date headers:")
        st.json(dq["repaired_cap_dates"])
    if dq.get("null_std_pack_pns"):
        st.caption(f"Part numbers with no Std Pack ({len(dq['null_std_pack_pns'])}):")
        st.write(", ".join(dq["null_std_pack_pns"][:30]) + ("…" if len(dq["null_std_pack_pns"]) > 30 else ""))
    if dq.get("unmatched_mrp_codes"):
        st.caption("MRP codes using fallback capacity:")
        st.write(", ".join(sorted(dq["unmatched_mrp_codes"])))

# ── Run ───────────────────────────────────────────────────────────────────────
st.subheader("4 · Run")

if st.button("▶ Run Query", type="primary", use_container_width=True):
    with st.spinner("Running scheduling logic…"):
        try:
            result = run_query(
                data_df,
                cap_df,
                capacity_overrides_df=sim_df,
                base_week=int(base_week),
                col_qty=col_qty,
                col_line=col_line,
                col_req_date=col_req_date,
                col_commit_date=col_commit_date,
                col_std_pack=col_std_pack,
                capacity_line_col=capacity_line_col,
                hrs_df=hrs_df,
                cost_df=cost_df,
                sp_df=sp_master_df,
            )
        except Exception as e:
            st.error(f"Error: {e}")
            st.exception(e)
            st.stop()

    # ── Metrics ───────────────────────────────────────────────────────────────
    st.success(f"✅ Done — {len(result):,} rows")
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Output rows",      f"{len(result):,}")
    m2.metric("Split orders",     f"{(result['SplitFlag']=='SPLIT').sum():,}")
    m3.metric("Lines",            result['Line'].nunique())
    m4.metric("Total Sched. Qty", f"{result['ScheduledQty'].sum():,.0f}")
    m5.metric("Total Excess Std Pack", f"{result['Excess Std Pack'].sum():,.0f}")

    # ── Summary pivot ─────────────────────────────────────────────────────────
    with st.expander("📊 Summary — Line × Production Week", expanded=True):
        summ = (
            result.groupby(["Line", "Line Name", "ProductionWeek"], sort=True)
            .agg(
                Orders          = ("ScheduledQty", "count"),
                ScheduledQty    = ("ScheduledQty", "sum"),
                Splits          = ("SplitFlag", lambda s: (s == "SPLIT").sum()),
            )
            .reset_index()
        )
        # Pivot for compact view
        pivot = summ.pivot_table(
            index=["Line", "Line Name"], columns="ProductionWeek",
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

    # ── MA3 Summary ───────────────────────────────────────────────────────────
    st.subheader("5 · MA3 Summary")
    st.caption("Pieces / Hours / Value by Line × Quarter × Month, based on ProductionWeekDate")

    if "Quarter" not in result.columns or "Month Name" not in result.columns:
        st.info("Date dimensions not available — ProductionWeekDate column required.")
    elif "Line Name" not in result.columns:
        st.info("Line Name column not found in result.")
    else:
        # Build quarterly + monthly totals
        def _ma3_section(label, metric_col, fmt_fn):
            if metric_col not in result.columns:
                st.warning(f"Column '{metric_col}' not in results.")
                return
            st.markdown(f"**{label}**")
            try:
                sub = result[result[metric_col].notna()].copy()
                sub[metric_col] = pd.to_numeric(sub[metric_col], errors="coerce").fillna(0)

                # Line × Month pivot
                monthly_pvt = sub.pivot_table(
                    index=["Line", "Line Name"],
                    columns="Month Name",
                    values=metric_col,
                    aggfunc="sum",
                    fill_value=0,
                )
                # Reorder months in calendar order
                month_order = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
                ordered_months = [m for m in month_order if m in monthly_pvt.columns]
                monthly_pvt = monthly_pvt[ordered_months] if ordered_months else monthly_pvt

                # Quarterly totals
                qtr_pvt = sub.pivot_table(
                    index=["Line", "Line Name"],
                    columns="Quarter",
                    values=metric_col,
                    aggfunc="sum",
                    fill_value=0,
                )

                # Combine: Quarter columns first, then Monthly
                combined = pd.concat([qtr_pvt, monthly_pvt], axis=1)
                combined["Total"] = combined.sum(axis=1)

                # Daily and weekly averages from result
                if "ProductionWeekDate" in result.columns:
                    week_dates = pd.to_datetime(result["ProductionWeekDate"], errors="coerce")
                    n_weeks = week_dates.dt.to_period("W").nunique()
                    n_days  = week_dates.dt.date.nunique()
                    if n_weeks > 0:
                        weekly_avg = sub.groupby(["Line", "Line Name"])[metric_col].sum() / n_weeks
                        daily_avg  = sub.groupby(["Line", "Line Name"])[metric_col].sum() / max(n_days, 1)
                        combined["Weekly avg"] = weekly_avg
                        combined["Daily avg"]  = daily_avg

                # Format display
                display = combined.copy()
                for col in display.columns:
                    display[col] = display[col].apply(fmt_fn)

                st.dataframe(display, use_container_width=True)
            except Exception as e:
                st.error(f"Error building {label}: {e}")

        _ma3_section("MA3 Pieces (Scheduled Qty)", "ScheduledQty", lambda v: f"{int(v):,}" if isinstance(v, (int, float)) else v)
        st.divider()
        _ma3_section("MA3 Hours (Total Hours)",    "Total Hours",  lambda v: f"{v:,.1f}" if isinstance(v, (int, float)) else v)
        st.divider()
        _ma3_section("MA3 Sales (Total Value $)",  "Total Value",  lambda v: f"${v:,.2f}" if isinstance(v, (int, float)) else v)

        # MA3 Summary download
        st.markdown("**Download MA3 Summary**")
        buf_ma3 = io.BytesIO()
        with pd.ExcelWriter(buf_ma3, engine="openpyxl") as w:
            for sheet_key, metric_col in [
                ("MA3_Pieces", "ScheduledQty"),
                ("MA3_Hours",  "Total Hours"),
                ("MA3_Sales",  "Total Value"),
            ]:
                if metric_col in result.columns:
                    pvt = result.pivot_table(
                        index=["Line", "Line Name"],
                        columns="Month Name",
                        values=metric_col,
                        aggfunc="sum",
                        fill_value=0,
                    )
                    pvt["Total"] = pvt.sum(axis=1)
                    pvt.reset_index().to_excel(w, sheet_name=sheet_key, index=False)
        st.download_button(
            "⬇️ MA3 Summary Excel",
            data=buf_ma3.getvalue(),
            file_name="ma3_summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    # ── Section 6: Build Analysis ─────────────────────────────────────────────
    st.divider()
    st.subheader("6 · Build Analysis (BOM-based)")
    st.caption(
        "Explodes the MPS schedule through the BOM to show component demand, "
        "raw material coverage, and build capability (FG output). "
        "Enter available stock to activate shortage and buildable-qty calculations."
    )

    if bom_df is None:
        st.info("Select a BOM sheet above to enable Build Analysis.")
    else:
        with st.spinner("Parsing BOM and exploding demand…"):
            cost_dated_lookup = build_cost_lookup_dated(cost_df)
            bom_clean = parse_bom(bom_df, alt_bom=1, bom_level=1)
            exploded_base = explode_bom_demand(
                result, bom_clean,
                fg_col="Product code",
                qty_col="ScheduledQty",
            )

        if exploded_base.empty:
            matched_fgs = set(bom_clean["Material"].astype(str)) & set(
                result["Product code"].astype(str) if "Product code" in result.columns else set()
            )
            st.warning(
                f"BOM parsed ({len(bom_clean):,} components from "
                f"{bom_clean['Material'].nunique():,} assemblies) but no orders "
                f"matched BOM materials. Matched FGs: {len(matched_fgs)}"
            )
        else:
            # ── BOM stats ─────────────────────────────────────────────────────
            n_fgs  = exploded_base["FG"].nunique()
            n_comp = exploded_base["Component"].nunique()
            total_demand = exploded_base["Comp_Demand"].sum()

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Matched FGs",        f"{n_fgs:,}")
            m2.metric("Unique Components",  f"{n_comp:,}")
            m3.metric("BOM Rows",           f"{len(bom_clean):,}")
            m4.metric("Total Comp. Demand", f"{total_demand:,.0f}")

            has_supply_tables = (
                (stock_df is not None and not stock_df.empty)
                or (rm_po_df is not None and not rm_po_df.empty)
            )
            if has_supply_tables:
                exploded_supply = compute_build_capability(
                    exploded_base,
                    stock_df=stock_df,
                    rm_po_df=rm_po_df,
                    week_col="ProductionWeek",
                    week_date_col="ProductionWeekDate",
                )
            else:
                exploded_supply = exploded_base.copy()

            # ── Tabs ──────────────────────────────────────────────────────────
            tab_demand, tab_stock, tab_capability, tab_cost, tab_coverage, tab_output = st.tabs([
                "📦 Component Demand",
                "🏷️ Stock Input & Shortage",
                "🏗️ Build Capability (FG Output)",
                "💰 Cost Traceability",
                "📋 RM Coverage Report",
                "🏭 Line Output Plan",
            ])

            # ── Tab 1: Component demand pivot ─────────────────────────────────
            with tab_demand:
                st.markdown("**Raw material demand by production week** — exploded through BOM")
                demand_pvt = component_demand_pivot(exploded_base, week_col="ProductionWeek")
                if not demand_pvt.empty:
                    week_cols = [c for c in demand_pvt.columns
                                 if c not in ("Component", "Component_Desc", "Total")]
                    fmt = {c: "{:,.0f}" for c in week_cols + ["Total"]}
                    st.dataframe(
                        demand_pvt.style.format(fmt).background_gradient(
                            subset=["Total"], cmap="YlOrRd"
                        ),
                        use_container_width=True, height=420,
                    )

                    with st.expander("📄 Full exploded demand (row-level)", expanded=False):
                        disp_cols = [c for c in [
                            "FG", "FG_Desc", "Component", "Component_Desc",
                            "BOM_Usage", "FG_Qty", "Comp_Demand",
                            "ProductionWeek", "ProductionWeekDate", "Line", "Line Name",
                        ] if c in exploded_base.columns]
                        st.dataframe(exploded_base[disp_cols], use_container_width=True)

                    buf_demand = io.BytesIO()
                    with pd.ExcelWriter(buf_demand, engine="openpyxl") as w:
                        demand_pvt.to_excel(w, sheet_name="Component Demand", index=False)
                        exploded_base.to_excel(w, sheet_name="Exploded Detail", index=False)
                    st.download_button(
                        "⬇️ Component Demand Excel",
                        data=buf_demand.getvalue(),
                        file_name="component_demand.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

            # ── Tab 2: Stock input & shortage ──────────────────────────────────
            with tab_stock:
                st.markdown(
                    "**CTB supply logic (weekly):** `Available Supply = Stock + cumulative RM_PO up to week`"
                )
                if not has_supply_tables:
                    st.warning(
                        "Load `Stock` and/or `RM_PO` sheets in Section 2 to enable time-based CTB. "
                        "Without these sheets, CTB defaults to unconstrained demand."
                    )
                else:
                    if "CTB_Flag" in exploded_supply.columns:
                        hard_zero_rows = int((exploded_supply["CTB_Flag"] == False).sum())
                        st.metric("Rows with CTB = 0 (no Stock and no PO by week)", f"{hard_zero_rows:,}")

                    short_df = shortage_report(exploded_supply, week_col="ProductionWeek")
                    if not short_df.empty:
                        st.markdown("#### ⚠️ Shortage Summary — Components × Week")
                        week_cols_s = [c for c in short_df.columns
                                       if c not in ("Component", "Component_Desc", "Total Shortage")]
                        fmt_s = {c: "{:,.0f}" for c in week_cols_s + ["Total Shortage"]}
                        st.dataframe(
                            short_df.style.format(fmt_s).background_gradient(
                                subset=["Total Shortage"], cmap="Reds"
                            ),
                            use_container_width=True,
                        )
                    else:
                        st.success("✅ No shortages — component demand is covered by Stock + RM_PO timeline.")

                    with st.expander("📄 Component-week CTB detail", expanded=False):
                        cols_ctb = [c for c in [
                            "Component", "Component_Desc", "ProductionWeek", "ProductionWeekDate",
                            "Comp_Demand", "Stock_Qty", "Cum_PO_Qty", "Available_Supply",
                            "Coverage_Pct", "CTB_Flag", "Shortage",
                        ] if c in exploded_supply.columns]
                        ctb_view = (
                            exploded_supply[cols_ctb]
                            .groupby([c for c in [
                                "Component", "Component_Desc", "ProductionWeek", "ProductionWeekDate",
                                "Stock_Qty", "Cum_PO_Qty", "Available_Supply", "Coverage_Pct", "CTB_Flag",
                            ] if c in cols_ctb], as_index=False)
                            .agg(Comp_Demand=("Comp_Demand", "sum"), Shortage=("Shortage", "sum"))
                            .sort_values(["Component", "ProductionWeek"]) if "ProductionWeek" in cols_ctb else exploded_supply[cols_ctb]
                        )
                        st.dataframe(ctb_view, use_container_width=True, height=380)

            # ── Tab 3: Build capability ────────────────────────────────────────
            with tab_capability:
                st.markdown("**Buildable FG quantity per part per week** — limited by most constrained component")

                enriched_cap = exploded_supply

                capability_df = build_capability_summary(
                    enriched_cap, cost_dated_lookup,
                    has_stock=has_supply_tables,
                )

                if not capability_df.empty:
                    # Summary metrics
                    total_buildable = capability_df["Buildable_Qty"].sum()
                    total_scheduled = capability_df["FG_Scheduled_Qty"].sum()
                    total_cost      = capability_df["Total_Cost"].sum() if "Total_Cost" in capability_df.columns else 0
                    avg_coverage    = capability_df["Material_Coverage_Pct"].mean()

                    ca, cb, cc, cd = st.columns(4)
                    ca.metric("Total Scheduled Qty", f"{total_scheduled:,.0f}")
                    cb.metric("Total Buildable Qty",  f"{total_buildable:,.0f}",
                              delta=f"{total_buildable-total_scheduled:+,.0f}" if has_supply_tables else None)
                    cc.metric("Avg Coverage %",       f"{avg_coverage:.1f}%")
                    cd.metric("Total Cost",           f"${total_cost:,.0f}")

                    # Pivot: FG × Week buildable qty
                    pn_col_cap = "Part_Number" if "Part_Number" in capability_df.columns else capability_df.columns[0]
                    week_col_cap = "ProductionWeek" if "ProductionWeek" in capability_df.columns else None

                    if week_col_cap and pn_col_cap in capability_df.columns:
                        buildable_pvt = capability_df.pivot_table(
                            index=[pn_col_cap],
                            columns=week_col_cap,
                            values="Buildable_Qty",
                            aggfunc="sum",
                            fill_value=0,
                        )
                        buildable_pvt["Total"] = buildable_pvt.sum(axis=1)
                        st.markdown("##### Buildable Qty — FG × Production Week")
                        bld_fmt = {c: "{:,.0f}" for c in buildable_pvt.columns}
                        st.dataframe(
                            buildable_pvt.style.format(bld_fmt).background_gradient(
                                cmap="Greens" if not has_supply_tables else "RdYlGn"
                            ),
                            use_container_width=True, height=400,
                        )

                    # Full detail table
                    with st.expander("📄 Full capability dataset (dashboard-ready)", expanded=False):
                        display_cols = [c for c in [
                            "Part_Number", "FG_Description",
                            "Line", "Line Name", "ProductionWeek", "ProductionWeekDate",
                            "Quarter", "Month Name",
                            "FG_Scheduled_Qty", "Buildable_Qty", "Material_Coverage_Pct",
                            "Constraint_Component", "Constraint_Desc",
                            "Num_BOM_Components", "Total_Shortage",
                            "Std_Cost", "Cost_Effective_Date", "Total_Cost",
                        ] if c in capability_df.columns]
                        st.dataframe(capability_df[display_cols], use_container_width=True)

                    # Coverage heatmap pivot
                    if has_supply_tables and week_col_cap:
                        with st.expander("🗺️ Material Coverage % heatmap — FG × Week", expanded=True):
                            cov_pvt = capability_df.pivot_table(
                                index=[pn_col_cap],
                                columns=week_col_cap,
                                values="Material_Coverage_Pct",
                                aggfunc="min",
                                fill_value=0,
                            )
                            st.dataframe(
                                cov_pvt.style.format("{:.1f}%").background_gradient(
                                    cmap="RdYlGn", vmin=0, vmax=100
                                ),
                                use_container_width=True, height=400,
                            )

                        # Bottleneck components
                        if "Constraint_Component" in capability_df.columns:
                            bottlenecks = (
                                capability_df[capability_df["Constraint_Component"].notna()]
                                .groupby("Constraint_Component")
                                .agg(
                                    Description=("Constraint_Desc", "first"),
                                    Affected_FGs=("Part_Number", "nunique"),
                                    Min_Coverage=("Material_Coverage_Pct", "min"),
                                    Total_Shortage=("Total_Shortage", "sum"),
                                )
                                .sort_values("Min_Coverage")
                                .reset_index()
                            )
                            if not bottlenecks.empty:
                                st.markdown("#### 🔴 Bottleneck Components")
                                st.dataframe(
                                    bottlenecks.style.format({
                                        "Min_Coverage": "{:.1f}%",
                                        "Total_Shortage": "{:,.0f}",
                                    }).background_gradient(subset=["Min_Coverage"], cmap="RdYlGn", vmin=0, vmax=100),
                                    use_container_width=True,
                                )

                    # Download
                    buf_cap = io.BytesIO()
                    with pd.ExcelWriter(buf_cap, engine="openpyxl") as w:
                        capability_df.to_excel(w, sheet_name="Build Capability", index=False)
                        if has_supply_tables and not shortage_report(enriched_cap).empty:
                            shortage_report(enriched_cap).to_excel(w, sheet_name="Shortage Report", index=False)
                        demand_pvt = component_demand_pivot(exploded_base)
                        if not demand_pvt.empty:
                            demand_pvt.to_excel(w, sheet_name="Component Demand", index=False)
                    st.download_button(
                        "⬇️ Build Analysis Excel",
                        data=buf_cap.getvalue(),
                        file_name="build_analysis.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )

            # ── Tab 4: Cost traceability ───────────────────────────────────────
            with tab_cost:
                st.markdown(
                    "**Standard Cost — most recent record per part** with traceability to the "
                    "source update date. Parts not in the latest update automatically fall back "
                    "to the previous available date (progressive fallback)."
                )
                if cost_dated_lookup:
                    n_dated    = len(cost_dated_lookup)
                    dates_used = {}
                    for info in cost_dated_lookup.values():
                        d = info.get("date", "unknown")
                        dates_used[d] = dates_used.get(d, 0) + 1

                    t1, t2 = st.columns(2)
                    t1.metric("Parts with cost data", f"{n_dated:,}")
                    with t2:
                        st.caption("Records by update date used:")
                        for d, cnt in sorted(dates_used.items(), reverse=True):
                            st.write(f"  `{d}` → {cnt:,} parts")

                    cost_trace = cost_date_summary(cost_df, cost_dated_lookup)
                    if not cost_trace.empty:
                        # Filter to parts in the current scheduled result
                        if "Product code" in result.columns:
                            scheduled_pns = set(result["Product code"].astype(str).str.strip())
                            cost_trace_filtered = cost_trace[cost_trace["Part Number"].isin(scheduled_pns)]
                            other_count = len(cost_trace) - len(cost_trace_filtered)
                            if not cost_trace_filtered.empty:
                                st.markdown("**Cost traceability for scheduled parts:**")
                                st.dataframe(
                                    cost_trace_filtered.style.format({"Std Cost Used": "${:.4f}"}),
                                    use_container_width=True, height=360,
                                )
                                if other_count > 0:
                                    with st.expander(f"Show all {len(cost_trace):,} parts with cost data"):
                                        st.dataframe(
                                            cost_trace.style.format({"Std Cost Used": "${:.4f}"}),
                                            use_container_width=True,
                                        )
                        else:
                            st.dataframe(
                                cost_trace.style.format({"Std Cost Used": "${:.4f}"}),
                                use_container_width=True,
                            )

                        buf_cost = io.BytesIO()
                        with pd.ExcelWriter(buf_cost, engine="openpyxl") as w:
                            cost_trace.to_excel(w, sheet_name="Cost Traceability", index=False)
                        st.download_button(
                            "⬇️ Cost Traceability Excel",
                            data=buf_cost.getvalue(),
                            file_name="cost_traceability.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
                else:
                    st.info("No Cost sheet loaded — select the Cost sheet in Section 2.")

            # ── Tab 5: RM Coverage Report ──────────────────────────────────────
            with tab_coverage:
                st.markdown(
                    "**Raw Material Coverage** — weekly MRP simulation per component: "
                    "`Ending Balance = Prior Balance + Receipts − Demand`"
                )

                cov_df, cov_kpi = build_rm_coverage_table(
                    exploded_base,
                    stock_df=stock_df,
                    rm_po_df=rm_po_df,
                )

                if cov_df.empty:
                    st.info(
                        "No coverage data. Ensure BOM explosion matched components "
                        "and Stock/RM_PO sheets are loaded."
                    )
                else:
                    info_cols_cov = ["Component", "Component_Desc", "UOM",
                                     "Initial_Inventory", "Metric"]
                    week_keys_cov = [c for c in cov_df.columns if c not in info_cols_cov]

                    # ── Top KPIs ──────────────────────────────────────────────
                    at_risk    = int((cov_kpi["Pct_Weeks_Covered"] < 100).sum())
                    peak_short  = cov_kpi["Peak_Shortage"].sum() if "Peak_Shortage" in cov_kpi.columns else 0
                    shortage_rows = cov_kpi[cov_kpi["First_Shortage_Week"] != "—"]
                    first_short = (
                        str(shortage_rows["First_Shortage_Week"].min())
                        if not shortage_rows.empty else "—"
                    )

                    ka, kb, kc = st.columns(3)
                    ka.metric("Materials at Risk",
                              f"{at_risk} / {len(cov_kpi)}",
                              delta=None if at_risk == 0 else f"{at_risk} shortage",
                              delta_color="inverse")
                    kb.metric("Total Peak Shortage",  f"{peak_short:,.0f}",
                              help="Sum of the worst-week deficit per component — minimum qty to procure to resolve all shortages.")
                    kc.metric("First Shortage Week",  first_short)

                    # ── Demand Audit ───────────────────────────────────────────
                    with st.expander("🔍 Demand Derivation Audit — ScheduledQty × BOM Usage", expanded=False):
                        sched_total   = int(result["ScheduledQty"].sum())
                        raw_total     = int(result["QtyNum"].sum()) if "QtyNum" in result.columns else 0
                        planned_total = int(result["QtyPlannedInput"].sum()) if "QtyPlannedInput" in result.columns else 0
                        fgs_sched     = result["Product code"].nunique() if "Product code" in result.columns else 0
                        fgs_bom       = bom_clean["Material"].nunique() if not bom_clean.empty else 0
                        fgs_matched   = len(
                            set(result["Product code"].dropna().astype(str))
                            & set(bom_clean["Material"].dropna().astype(str))
                        ) if not bom_clean.empty else 0

                        st.markdown(
                            "RM demand is derived **strictly** from the capacity-scheduled, "
                            "std-pack-adjusted production plan. "
                            "Formula: `RM Demand = ScheduledQty × BOM_Usage`"
                        )
                        col_a, col_b, col_c, col_d = st.columns(4)
                        col_a.metric("Original Order Qty",    f"{raw_total:,}",
                                     help="Sum of raw order quantities before any processing.")
                        col_b.metric("After Std Pack Round",  f"{planned_total:,}",
                                     help="Rounded up to nearest std pack multiple.")
                        col_c.metric("ScheduledQty (CTB input)", f"{sched_total:,}",
                                     help="Post-capacity leveling and week-split. Used as demand driver for RM.")
                        col_d.metric("BOM Match Rate",
                                     f"{100*fgs_matched//max(1,fgs_sched)}%",
                                     help=f"{fgs_matched} of {fgs_sched} scheduled FGs found in BOM.")

                        if not exploded_base.empty:
                            audit_sample = (
                                exploded_base
                                .groupby(["FG","Component","BOM_Usage","ProductionWeek"], as_index=False)
                                ["Comp_Demand"].sum()
                                .head(8)
                            )
                            audit_sample["Formula"] = (
                                audit_sample["FG"] + " × " + audit_sample["BOM_Usage"].map("{:.4g}".format)
                                + " (BOM) = " + audit_sample["Comp_Demand"].map("{:,.0f}".format)
                            )
                            st.caption("Sample demand derivation rows (FG Qty × BOM Usage = Component Demand):")
                            st.dataframe(
                                audit_sample[["FG","Component","BOM_Usage","ProductionWeek","Comp_Demand","Formula"]],
                                use_container_width=True, height=260,
                            )

                    # ── Styled coverage table ─────────────────────────────────
                    display_cov = cov_df.copy()
                    renamed_weeks = [
                        f"Wk {c}" if isinstance(c, int) else str(c)
                        for c in week_keys_cov
                    ]
                    display_cov.columns = (
                        ["Part Number", "Description", "UOM",
                         "Initial Inventory", "Metric"]
                        + renamed_weeks
                    )
                    num_cols_cov = renamed_weeks + ["Initial Inventory"]

                    def _style_cov(df: pd.DataFrame) -> pd.DataFrame:
                        styles = pd.DataFrame("", index=df.index, columns=df.columns)
                        for idx in df.index:
                            metric = df.at[idx, "Metric"]
                            for wk in renamed_weeks:
                                if wk not in df.columns:
                                    continue
                                raw = df.at[idx, wk]
                                if metric == "Ending Balance":
                                    try:
                                        v = float(raw)
                                    except (ValueError, TypeError):
                                        continue
                                    if v < 0:
                                        styles.at[idx, wk] = (
                                            "background-color:#FFCDD2;"
                                            "color:#B71C1C;font-weight:bold"
                                        )
                                    elif v > 0:
                                        styles.at[idx, wk] = (
                                            "background-color:#C8E6C9;color:#1B5E20"
                                        )
                                    else:
                                        styles.at[idx, wk] = (
                                            "background-color:#F5F5F5;color:#757575"
                                        )
                                elif metric == "Receipts":
                                    styles.at[idx, wk] = "background-color:#E8F5E9"
                        return styles

                    fmt_cov = {c: "{:,.0f}" for c in num_cols_cov}
                    styled_cov = (
                        display_cov.style
                        .apply(_style_cov, axis=None)
                        .set_table_styles([{
                            "selector": "thead th",
                            "props": [
                                ("background-color", "#1B5E20"),
                                ("color", "white"),
                                ("font-weight", "bold"),
                            ],
                        }])
                        .format(fmt_cov, na_rep="")
                    )
                    st.dataframe(
                        styled_cov,
                        use_container_width=True,
                        height=min(700, len(cov_df) * 35 + 55),
                    )

                    # ── KPI summary table ─────────────────────────────────────
                    with st.expander("📊 Per-Material KPI Summary", expanded=True):
                        kpi_fmt = {
                            "Initial_Inventory": "{:,.0f}",
                            "Pct_Weeks_Covered": "{:.1f}%",
                            "Peak_Shortage":     "{:,.0f}",
                            "Net_Deficit_Weeks": "{:,.0f}",
                        }
                        kpi_cols_show = [c for c in [
                            "Component", "Description", "Initial_Inventory",
                            "First_Shortage_Week", "Pct_Weeks_Covered",
                            "Peak_Shortage", "Net_Deficit_Weeks",
                        ] if c in cov_kpi.columns]
                        grad_cols = [c for c in ["Pct_Weeks_Covered","Peak_Shortage","Net_Deficit_Weeks"]
                                     if c in cov_kpi.columns]
                        st.caption(
                            "**Peak Shortage** = maximum deficit at any single week (minimum qty to procure).  "
                            "**Net Deficit Weeks** = sum of running negative balances across all weeks (urgency index)."
                        )
                        sty = cov_kpi[kpi_cols_show].style.format(kpi_fmt)
                        if "Pct_Weeks_Covered" in grad_cols:
                            sty = sty.background_gradient(
                                subset=["Pct_Weeks_Covered"], cmap="RdYlGn", vmin=0, vmax=100
                            )
                        if "Peak_Shortage" in grad_cols:
                            sty = sty.background_gradient(subset=["Peak_Shortage"], cmap="Reds")
                        st.dataframe(sty, use_container_width=True)

                    # ── Excel download ────────────────────────────────────────
                    cov_excel = rm_coverage_to_excel(cov_df, cov_kpi)
                    st.download_button(
                        "⬇️ RM Coverage Excel (Formatted)",
                        data=cov_excel,
                        file_name="rm_coverage_report.xlsx",
                        mime=(
                            "application/vnd.openxmlformats-"
                            "officedocument.spreadsheetml.sheet"
                        ),
                        use_container_width=True,
                    )

            # ── Tab 6: Line × Week Output Plan ──────────────────────────────
            with tab_output:
                st.markdown(
                    "**Production Output Plan — Line × Week** "
                    "(`FG Output` = what CAN be built given CTB; `Shortage` = risk pcs)"
                )

                if "Line" not in result.columns:
                    st.warning(
                        "Column `Line` not found in schedule output. "
                        "Check that the MPS sheet contains an MRP/Line column."
                    )
                else:
                    out_plan_df, out_kpi_df = build_line_output_plan(
                        result,
                        capability_df=capability_df if not capability_df.empty else None,
                        week_col="ProductionWeek",
                        qty_col="ScheduledQty",
                    )

                    if out_plan_df.empty:
                        st.info("No line output data could be generated.")
                    else:
                        # ── Top KPIs ─────────────────────────────────────────
                        total_pl   = out_kpi_df["Total_Planned"].sum()
                        total_out  = out_kpi_df["Total_FG_Output"].sum()
                        total_sh   = out_kpi_df["Total_Shortage_Pcs"].sum()
                        lines_risk = int((out_kpi_df["Total_Shortage_Pcs"] > 0).sum())
                        pct_global = round(total_out / total_pl * 100, 1) if total_pl > 0 else 100.0

                        oa, ob, oc, od = st.columns(4)
                        oa.metric("Total Planned",      f"{total_pl:,.0f}")
                        ob.metric("Total FG Output",    f"{total_out:,.0f}",
                                  delta=f"{total_out - total_pl:+,.0f}" if has_supply_tables else None,
                                  delta_color="inverse")
                        oc.metric("Total Shortage Pcs", f"{total_sh:,.0f}",
                                  delta_color="inverse")
                        od.metric("Lines at Risk",
                                  f"{lines_risk} / {len(out_kpi_df)}",
                                  delta=f"{pct_global:.1f}% achievable",
                                  delta_color="off")

                        # ── 3-layer pivot: styled table ───────────────────────
                        st.markdown("##### Production Output — Line × Production Week")
                        info_cols_op = ["Line", "Line_Name", "Metric"]
                        wk_cols_op   = [c for c in out_plan_df.columns if c not in info_cols_op]

                        def _style_output_plan(df: pd.DataFrame) -> pd.DataFrame:
                            """Return DataFrame of CSS strings aligned with df shape."""
                            styles = pd.DataFrame("", index=df.index, columns=df.columns)
                            for idx in df.index:
                                metric = df.loc[idx, "Metric"] if "Metric" in df.columns else ""
                                if metric == "FG Output":
                                    bg, fg = "#C8E6C9", "#1B5E20"
                                elif metric == "Shortage":
                                    bg, fg = "#FFCDD2", "#B71C1C"
                                else:
                                    bg, fg = "#ECEFF1", "#263238"
                                for col in df.columns:
                                    styles.loc[idx, col] = (
                                        f"background-color:{bg};color:{fg};font-weight:bold;"
                                        if metric in ("FG Output", "Shortage")
                                        else f"background-color:{bg};color:{fg};"
                                    )
                            return styles

                        display_op = out_plan_df.copy()
                        display_op = display_op.rename(columns={"Line_Name": "Line Name"})
                        fmt_op = {c: "{:,.0f}" for c in wk_cols_op
                                  if pd.api.types.is_numeric_dtype(out_plan_df[c])}

                        st.dataframe(
                            display_op.style
                            .apply(_style_output_plan, axis=None)
                            .format(fmt_op, na_rep="—"),
                            use_container_width=True,
                            height=min(60 + len(out_plan_df) * 36, 620),
                        )

                        # ── KPI breakdown by line ─────────────────────────────
                        with st.expander("📊 Line KPI Breakdown", expanded=True):
                            kpi_display = out_kpi_df.copy().rename(
                                columns={"Line_Name": "Line Name",
                                         "Total_Planned": "Planned",
                                         "Total_FG_Output": "FG Output",
                                         "Total_Shortage_Pcs": "Shortage Pcs",
                                         "Pct_Achievable": "% Achievable",
                                         "First_Shortage_Week": "First Shortage Wk",
                                         "Recovery_Week": "Recovery Wk"}
                            )
                            def _style_kpi(df: pd.DataFrame) -> pd.DataFrame:
                                styles = pd.DataFrame("", index=df.index, columns=df.columns)
                                pct_col = "% Achievable" if "% Achievable" in df.columns else None
                                sh_col  = "Shortage Pcs" if "Shortage Pcs" in df.columns else None
                                for idx in df.index:
                                    if pct_col:
                                        try:
                                            pct = float(df.loc[idx, pct_col])
                                        except (TypeError, ValueError):
                                            pct = 100.0
                                        color = (
                                            "#C8E6C9" if pct >= 95
                                            else "#FFF9C4" if pct >= 80
                                            else "#FFCDD2"
                                        )
                                        styles.loc[idx, pct_col] = f"background-color:{color};font-weight:bold;"
                                    if sh_col:
                                        try:
                                            sh = float(df.loc[idx, sh_col])
                                        except (TypeError, ValueError):
                                            sh = 0.0
                                        if sh > 0:
                                            styles.loc[idx, sh_col] = "background-color:#FFCDD2;color:#B71C1C;font-weight:bold;"
                                return styles

                            fmt_kpi = {
                                "Planned":    "{:,.0f}",
                                "FG Output":  "{:,.0f}",
                                "Shortage Pcs": "{:,.0f}",
                                "% Achievable": "{:.1f}%",
                            }
                            st.dataframe(
                                kpi_display.style
                                .apply(_style_kpi, axis=None)
                                .format(fmt_kpi, na_rep="—"),
                                use_container_width=True,
                            )

                        # ── Separate pivots: Planned / FG Output / Shortage ───
                        with st.expander("📋 Pivot: Planned vs FG Output vs Shortage (split view)", expanded=False):
                            for metric_label, fill_hex in [
                                ("Planned",   None),
                                ("FG Output", "#C8E6C9"),
                                ("Shortage",  "#FFCDD2"),
                            ]:
                                subset = out_plan_df[out_plan_df["Metric"] == metric_label].copy()
                                if subset.empty:
                                    continue
                                pvt = subset.drop(columns=["Metric"]).set_index(["Line", "Line_Name"])
                                pvt.index.names = ["Line", "Line Name"]
                                st.markdown(f"**{metric_label}**")
                                fmt_pvt = {c: "{:,.0f}" for c in pvt.columns
                                           if pd.api.types.is_numeric_dtype(pvt[c])}
                                if fill_hex:
                                    st.dataframe(
                                        pvt.style.format(fmt_pvt, na_rep="—")
                                        .map(lambda _: f"background-color:{fill_hex};"),
                                        use_container_width=True,
                                    )
                                else:
                                    st.dataframe(pvt.style.format(fmt_pvt, na_rep="—"),
                                                 use_container_width=True)

                        # ── Excel download ────────────────────────────────────
                        out_xlsx = line_output_to_excel(out_plan_df, out_kpi_df)
                        st.download_button(
                            "⬇️ Line Output Plan Excel (Formatted)",
                            data=out_xlsx,
                            file_name="line_output_plan.xlsx",
                            mime=(
                                "application/vnd.openxmlformats-"
                                "officedocument.spreadsheetml.sheet"
                            ),
                            use_container_width=True,
                        )

# ── M code viewer ─────────────────────────────────────────────────────────────
with st.expander("📝 Current M code (Excel.xlsx)", expanded=False):
    try:
        st.code(open("Excel.xlsx").read(), language="plaintext")
    except Exception:
        st.warning("Could not read Excel.xlsx.")
