"""
bom_analysis.py
Build Analysis engine: BOM explosion, build capability, and cost with
progressive date fallback.
"""
from __future__ import annotations

import math
import pandas as pd


# ── helpers ──────────────────────────────────────────────────────────────────

def _col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """Return the first column name that matches any candidate (case-insensitive)."""
    lower_map = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        found = lower_map.get(cand.strip().lower())
        if found is not None:
            return found
    return None


# ── 1. Std Cost with progressive date fallback ────────────────────────────────

def build_cost_lookup_dated(
    cost_df: pd.DataFrame | None,
) -> dict[str, dict]:
    """
    Returns {PN: {"cost": float, "date": str}} using the most recent
    non-zero cost per part number.

    Progressive fallback logic:
      Sort all records by update date DESC.
      For each PN, take the first row whose Std Cost > 0.
      This naturally falls back to earlier dates when the latest record
      has a zero or missing cost.

    Fields returned per PN:
      "cost"  – float, the selected Std Cost
      "date"  – str ISO-8601 date of the record used (traceability)
    """
    if cost_df is None or cost_df.empty:
        return {}

    df = cost_df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    pn_col   = _col(df, ["PN", "Material", "Part Number", "Part"])
    cost_col = _col(df, ["Std Cost", "StdCost", "Cost", "Price"])
    date_col = _col(df, ["Last update", "LastUpdate", "Update Date", "Date"])

    if pn_col is None or cost_col is None:
        return {}

    df["_pn"]   = df[pn_col].astype(str).str.strip()
    df["_cost"] = pd.to_numeric(df[cost_col], errors="coerce")
    df["_date"] = (
        pd.to_datetime(df[date_col], errors="coerce")
        if date_col
        else pd.Timestamp("1900-01-01")
    )

    # Keep only rows with a valid positive cost
    valid = df[df["_cost"].notna() & (df["_cost"] > 0)].copy()
    # Sort descending so groupby.first() picks the most recent
    valid = valid.sort_values("_date", ascending=False)

    best = valid.groupby("_pn", sort=False).first().reset_index()

    result: dict[str, dict] = {}
    for _, row in best.iterrows():
        d = row["_date"]
        result[row["_pn"]] = {
            "cost": float(row["_cost"]),
            "date": d.strftime("%Y-%m-%d") if pd.notna(d) else "",
        }
    return result


def cost_date_summary(
    cost_df: pd.DataFrame | None,
    cost_dated_lookup: dict,
) -> pd.DataFrame:
    """
    Returns a traceability table: PN, Std Cost Used, Cost Effective Date.
    Includes only the parts that appear in the dated lookup (i.e., have a cost).
    """
    if not cost_dated_lookup:
        return pd.DataFrame(columns=["Part Number", "Std Cost Used", "Cost Effective Date"])

    rows = [
        {"Part Number": pn, "Std Cost Used": info["cost"], "Cost Effective Date": info["date"]}
        for pn, info in cost_dated_lookup.items()
    ]
    return pd.DataFrame(rows).sort_values("Part Number").reset_index(drop=True)


# ── 2. BOM parsing ────────────────────────────────────────────────────────────

def parse_bom(
    bom_df: pd.DataFrame | None,
    alt_bom: int = 1,
    bom_level: int = 1,
) -> pd.DataFrame:
    """
    Returns a clean BOM DataFrame with columns:
      Material, Component, Component_Desc, BOM_Usage, BOM_Level, Alt_BOM, MRP

    BOM_Usage = qty of Component per 1 unit of Material.
    Filters to primary BOM (Alt_BOM == alt_bom) and the requested BOM level.
    """
    if bom_df is None or bom_df.empty:
        return pd.DataFrame()

    df = bom_df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    mat_col    = _col(df, ["Material"])
    comp_col   = _col(df, ["Component"])
    desc_col   = _col(df, ["Component Description", "Component Desc"])
    qty_col    = _col(df, ["Quantity", "Qty"])
    base_col   = _col(df, ["Base quantity", "Base Quantity", "BaseQty"])
    usage_col  = _col(df, ["Usage"])
    level_col  = _col(df, ["BOM Level", "Level"])
    altbom_col = _col(df, ["Alt BOM", "AltBOM", "Alt_BOM"])
    mrp_col    = _col(df, ["MRP"])
    mat_desc_col = _col(df, ["Material Description"])

    if mat_col is None or comp_col is None:
        return pd.DataFrame()

    out = pd.DataFrame()
    out["Material"]       = df[mat_col].astype(str).str.strip()
    out["Component"]      = df[comp_col].astype(str).str.strip()
    out["Component_Desc"] = (
        df[desc_col].astype(str).str.strip() if desc_col else out["Component"]
    )
    out["Material_Desc"]  = (
        df[mat_desc_col].astype(str).str.strip() if mat_desc_col else out["Material"]
    )
    out["BOM_Level"] = (
        pd.to_numeric(df[level_col], errors="coerce").fillna(1).astype(int)
        if level_col else 1
    )
    out["Alt_BOM"] = (
        pd.to_numeric(df[altbom_col], errors="coerce").fillna(1).astype(int)
        if altbom_col else 1
    )
    out["MRP"] = df[mrp_col].astype(str).str.strip() if mrp_col else ""

    # Compute usage per unit of parent
    if usage_col:
        out["BOM_Usage"] = pd.to_numeric(df[usage_col], errors="coerce").fillna(0.0)
    elif qty_col and base_col:
        qty  = pd.to_numeric(df[qty_col],  errors="coerce").fillna(0.0)
        base = pd.to_numeric(df[base_col], errors="coerce").replace(0, 1000.0).fillna(1000.0)
        out["BOM_Usage"] = qty / base
    elif qty_col:
        out["BOM_Usage"] = pd.to_numeric(df[qty_col], errors="coerce").fillna(0.0)
    else:
        out["BOM_Usage"] = 1.0

    # Filters
    out = out[
        out["Material"].notna() & (out["Material"] != "nan") & (out["Material"] != "")
        & out["Component"].notna() & (out["Component"] != "nan") & (out["Component"] != "")
        & (out["BOM_Usage"] > 0)
    ]
    out = out[out["Alt_BOM"] == alt_bom]
    out = out[out["BOM_Level"] == bom_level]

    return out.reset_index(drop=True)


# ── 3. BOM explosion ──────────────────────────────────────────────────────────

def explode_bom_demand(
    scheduled_df:  pd.DataFrame,
    bom_clean:     pd.DataFrame,
    fg_col:        str = "Product code",
    qty_col:       str = "ScheduledQty",
    week_col:      str = "ProductionWeek",
    week_date_col: str = "ProductionWeekDate",
) -> pd.DataFrame:
    """
    Explodes the MPS schedule through the BOM to produce component demand
    per FG per production week.

    Returns columns:
      FG, FG_Desc, Component, Component_Desc, BOM_Usage,
      ProductionWeek, ProductionWeekDate, [Line, Line Name, Quarter, Month Name],
      FG_Qty, Comp_Demand
    """
    if scheduled_df is None or scheduled_df.empty:
        return pd.DataFrame()
    if bom_clean is None or bom_clean.empty:
        return pd.DataFrame()
    if fg_col not in scheduled_df.columns:
        return pd.DataFrame()

    sch = scheduled_df.copy()
    bom = bom_clean.copy()

    sch["_fg"]  = sch[fg_col].astype(str).str.strip()
    bom["_mat"] = bom["Material"].astype(str).str.strip()

    bom_cols = ["_mat", "Component", "Component_Desc", "Material_Desc", "BOM_Usage"]
    bom_cols = [c for c in bom_cols if c in bom.columns]

    merged = sch.merge(bom[bom_cols], left_on="_fg", right_on="_mat", how="inner")
    if merged.empty:
        return pd.DataFrame()

    fg_qty_num = pd.to_numeric(merged[qty_col], errors="coerce").fillna(0)
    merged["Comp_Demand"] = fg_qty_num * merged["BOM_Usage"]

    keep: dict[str, pd.Series] = {
        "FG":             merged["_fg"],
        "FG_Desc":        merged.get("Material_Desc", merged["_fg"]),
        "Component":      merged["Component"],
        "Component_Desc": merged["Component_Desc"],
        "BOM_Usage":      merged["BOM_Usage"],
        "FG_Qty":         fg_qty_num,
        "Comp_Demand":    merged["Comp_Demand"],
    }

    # Pass through dimension columns if present
    for c in [week_col, week_date_col, "Line", "Line Name", "Quarter", "Month Name"]:
        if c in merged.columns:
            keep[c] = merged[c]

    return pd.DataFrame(keep).reset_index(drop=True)


# ── 4. Component demand pivot (raw material coverage view) ───────────────────

def component_demand_pivot(
    exploded_df: pd.DataFrame,
    week_col: str = "ProductionWeek",
) -> pd.DataFrame:
    """
    Returns a pivot: Component × Week → total demand quantity.
    Includes Component_Desc and an overall Total column.
    """
    if exploded_df is None or exploded_df.empty:
        return pd.DataFrame()
    if week_col not in exploded_df.columns:
        return pd.DataFrame()

    pvt = exploded_df.pivot_table(
        index=["Component", "Component_Desc"],
        columns=week_col,
        values="Comp_Demand",
        aggfunc="sum",
        fill_value=0,
    )
    pvt["Total"] = pvt.sum(axis=1)
    pvt = pvt.sort_values("Total", ascending=False)
    return pvt.reset_index()


# ── 5. Build capability computation ──────────────────────────────────────────

def compute_build_capability(
    exploded_df: pd.DataFrame,
    stock_lookup: dict[str, float],
) -> pd.DataFrame:
    """
    Enriches exploded demand with stock-based metrics:
      Available_Qty   – from stock_lookup (0 if unknown)
      Coverage_Pct    – min(100, available / demand * 100)
      Shortage        – max(0, demand - available)
      Buildable_From_Comp – available / BOM_Usage  (FG units this component supports)
    """
    df = exploded_df.copy()
    df["Available_Qty"] = df["Component"].map(stock_lookup).fillna(0.0)

    def _coverage(row):
        if row["Comp_Demand"] <= 0:
            return 100.0
        return min(100.0, row["Available_Qty"] / row["Comp_Demand"] * 100.0)

    def _buildable(row):
        if row["BOM_Usage"] <= 0:
            return row["FG_Qty"]
        return row["Available_Qty"] / row["BOM_Usage"]

    df["Coverage_Pct"]        = df.apply(_coverage, axis=1)
    df["Buildable_From_Comp"] = df.apply(_buildable, axis=1)
    df["Shortage"]            = (df["Comp_Demand"] - df["Available_Qty"]).clip(lower=0)
    return df


def build_capability_summary(
    enriched_df:       pd.DataFrame,
    cost_dated_lookup: dict,
    has_stock:         bool = False,
    week_col:          str = "ProductionWeek",
    week_date_col:     str = "ProductionWeekDate",
) -> pd.DataFrame:
    """
    Summarises build capability per FG × Week.

    Returns the dashboard-ready dataset:
      Part Number, FG_Desc, Line, Line Name, Week, WeekDate, Quarter, Month Name,
      FG_Scheduled_Qty, Buildable_Qty, Material_Coverage_Pct,
      Constraint_Component, Constraint_Desc,
      Num_BOM_Components, Total_Shortage,
      Std_Cost, Cost_Effective_Date, Total_Cost
    """
    if enriched_df is None or enriched_df.empty:
        return pd.DataFrame()

    grp_keys = [c for c in
                ["FG", "FG_Desc", week_col, week_date_col,
                 "Line", "Line Name", "Quarter", "Month Name"]
                if c in enriched_df.columns]

    def _agg(g: pd.DataFrame) -> pd.Series:
        fg_qty = float(g["FG_Qty"].iloc[0])

        if has_stock and "Coverage_Pct" in g.columns:
            idx_bot         = g["Coverage_Pct"].idxmin()
            constraint_comp = g.loc[idx_bot, "Component"]
            constraint_desc = g.loc[idx_bot, "Component_Desc"] if "Component_Desc" in g.columns else ""
            coverage        = float(g["Coverage_Pct"].min())
            buildable       = float(max(0.0, g["Buildable_From_Comp"].min()))
            shortage        = float(g["Shortage"].sum())
        else:
            constraint_comp = None
            constraint_desc = None
            coverage        = 100.0
            buildable       = fg_qty
            shortage        = 0.0

        return pd.Series({
            "FG_Scheduled_Qty":      fg_qty,
            "Buildable_Qty":         buildable,
            "Material_Coverage_Pct": round(coverage, 1),
            "Constraint_Component":  constraint_comp,
            "Constraint_Desc":       constraint_desc,
            "Num_BOM_Components":    int(len(g)),
            "Total_Shortage":        shortage,
        })

    summary = enriched_df.groupby(grp_keys, sort=True).apply(
        _agg, include_groups=False
    ).reset_index()

    summary = summary.rename(columns={"FG": "Part_Number", "FG_Desc": "FG_Description"})

    # Add cost columns
    if cost_dated_lookup and "Part_Number" in summary.columns:
        pn = summary["Part_Number"].astype(str).str.strip()
        summary["Std_Cost"]           = pn.map(lambda p: cost_dated_lookup.get(p, {}).get("cost", 0.0))
        summary["Cost_Effective_Date"] = pn.map(lambda p: cost_dated_lookup.get(p, {}).get("date", ""))
        summary["Total_Cost"]         = summary["Buildable_Qty"] * summary["Std_Cost"]

    return summary


# ── 6. Shortage analysis ──────────────────────────────────────────────────────

def shortage_report(
    enriched_df: pd.DataFrame,
    week_col: str = "ProductionWeek",
) -> pd.DataFrame:
    """
    Returns a shortage summary: Component × Week → Shortage qty,
    sorted by total shortage descending.
    """
    if "Shortage" not in enriched_df.columns or enriched_df.empty:
        return pd.DataFrame()

    pvt = enriched_df.pivot_table(
        index=["Component", "Component_Desc"],
        columns=week_col,
        values="Shortage",
        aggfunc="sum",
        fill_value=0,
    )
    pvt["Total Shortage"] = pvt.sum(axis=1)
    pvt = pvt[pvt["Total Shortage"] > 0].sort_values("Total Shortage", ascending=False)
    return pvt.reset_index()


# ── 7. Full pipeline ──────────────────────────────────────────────────────────

def run_build_analysis(
    scheduled_df:  pd.DataFrame,
    bom_df:        pd.DataFrame | None,
    cost_df:       pd.DataFrame | None,
    stock_lookup:  dict[str, float] | None = None,
    fg_col:        str = "Product code",
    qty_col:       str = "ScheduledQty",
    week_col:      str = "ProductionWeek",
    week_date_col: str = "ProductionWeekDate",
    alt_bom:       int = 1,
    bom_level:     int = 1,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, dict]:
    """
    Full build analysis pipeline.

    Returns:
      bom_clean        – parsed and filtered BOM
      exploded         – component demand per FG per week (enriched with stock if provided)
      capability_df    – dashboard-ready capability summary per FG × week
      cost_dated_lookup – {PN: {"cost", "date"}} for traceability
    """
    cost_dated_lookup = build_cost_lookup_dated(cost_df)
    bom_clean         = parse_bom(bom_df, alt_bom=alt_bom, bom_level=bom_level)

    exploded = explode_bom_demand(
        scheduled_df, bom_clean,
        fg_col=fg_col, qty_col=qty_col,
        week_col=week_col, week_date_col=week_date_col,
    )

    if exploded.empty:
        return bom_clean, exploded, pd.DataFrame(), cost_dated_lookup

    has_stock = bool(stock_lookup)
    if has_stock:
        exploded = compute_build_capability(exploded, stock_lookup)

    capability = build_capability_summary(
        exploded, cost_dated_lookup,
        has_stock=has_stock,
        week_col=week_col, week_date_col=week_date_col,
    )

    return bom_clean, exploded, capability, cost_dated_lookup
