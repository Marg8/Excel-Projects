import math
from datetime import timedelta

import pandas as pd

# ── module-level constants (imported by app.py) ─────────────────────────────

DEFAULT_CAPS: dict[str, int] = {
    "HC8_H3H": 35000,
    "H4Z": 8250,
    "HB5": 32000,
    "H5C_HC9": 6100,
    "HC7": 19000,
    "H2J": 9500,
    "HC4_A9R": 35000,
    "HC0": 16000,
}

DEFAULT_COL_QTY = "Qty"
DEFAULT_COL_LINE = "MRP"
DEFAULT_COL_REQ_DATE = "Requested date"
DEFAULT_COL_COMMIT_DATE = "Plan Request Date"
DEFAULT_COL_STD_PACK = "Std Pack"
DEFAULT_COL_PART = "Product code"    # join key
DEFAULT_COL_HRS = "Std x Hr"        # in Hrs sheet
DEFAULT_COL_COST = "Std Cost"       # in Cost sheet
DEFAULT_COL_MATERIAL = "Material"   # key in master tables


# ── helpers ──────────────────────────────────────────────────────────────────

def to_number(x) -> float:
    if x is None:
        return 0.0
    if isinstance(x, (int, float)):
        return 0.0 if math.isnan(x) else float(x)
    s = str(x).strip().replace(",", "")
    if s in ("", "-", "None", "nan", "NaN"):
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def to_date(x):
    if x is None:
        return None
    if isinstance(x, pd.Timestamp):
        return x.normalize()
    try:
        return pd.to_datetime(x).normalize()
    except Exception:
        return None


def normalize_line(x) -> str | None:
    if x is None:
        return None
    t = str(x).strip().upper()
    if not t:
        return None

    # Keep planning keyed by standard codes used in capacity.
    if t in {"HC8", "H3H"}:
        return "HC8_H3H"
    if t in {"H5C", "HC9"}:
        return "H5C_HC9"
    if t in {"HC4", "A9R"}:
        return "HC4_A9R"
    return t


def _find_first_existing_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    by_lower = {str(c).strip().lower(): c for c in df.columns}
    for name in candidates:
        col = by_lower.get(name.strip().lower())
        if col is not None:
            return col
    return None


def _round_up_to_std_pack(order_qty: float, std_pack: float) -> tuple[float, float]:
    q = to_number(order_qty)
    sp = to_number(std_pack)

    if q <= 0:
        return 0.0, 0.0
    if sp <= 0:
        return q, 0.0

    planned = float(math.ceil(q / sp) * sp)
    excess = max(0.0, planned - q)
    return planned, excess


# ── Master table enrichment ──────────────────────────────────────────────────

def build_hrs_lookup(
    hrs_df: pd.DataFrame | None,
    material_col: str = DEFAULT_COL_MATERIAL,
    hrs_col: str = DEFAULT_COL_HRS,
) -> dict[str, float]:
    """
    Returns {Material -> HrsPerUnit} by summing all routing operations per part.
    Base quantity in SAP routing is per 1000 units; Std x Hr is already hr/unit.
    """
    if hrs_df is None or hrs_df.empty:
        return {}
    df = hrs_df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    mat_col = _find_first_existing_col(df, [material_col, "Material"])
    h_col   = _find_first_existing_col(df, [hrs_col, "Std x Hr", "Std_x_Hr", "StdxHr"])
    if mat_col is None or h_col is None:
        return {}
    df["_mat"] = df[mat_col].astype(str).str.strip()
    df["_hrs"] = df[h_col].apply(to_number)
    df = df[df["_hrs"] > 0]
    return df.groupby("_mat")["_hrs"].sum().to_dict()


def build_cost_lookup(
    cost_df: pd.DataFrame | None,
    pn_col: str = "PN",
    cost_col: str = DEFAULT_COL_COST,
    date_col: str = "Last update",
) -> dict[str, float]:
    """
    Returns {PartNumber -> StdCost} using the most recent non-zero cost.
    Progressive fallback: if the latest update date has a zero/null cost for a
    part, the previous available record is used automatically.
    """
    if cost_df is None or cost_df.empty:
        return {}
    df = cost_df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    p_col = _find_first_existing_col(df, [pn_col, DEFAULT_COL_MATERIAL, "Material", "Part"])
    c_col = _find_first_existing_col(df, [cost_col, "Std Cost", "StdCost", "Cost"])
    d_col = _find_first_existing_col(df, [date_col, "Last update", "Update Date", "Date"])
    if p_col is None or c_col is None:
        return {}
    df["_pn"]   = df[p_col].astype(str).str.strip()
    df["_cost"] = pd.to_numeric(df[c_col], errors="coerce")
    df["_date"] = (
        pd.to_datetime(df[d_col], errors="coerce")
        if d_col
        else pd.Timestamp("1900-01-01")
    )
    # Keep valid positive costs; sort descending so first per group = most recent
    valid = df[df["_cost"].notna() & (df["_cost"] > 0)].copy()
    valid = valid.sort_values("_date", ascending=False)
    best  = valid.groupby("_pn", sort=False).first().reset_index()
    return best.set_index("_pn")["_cost"].to_dict()


def build_std_pack_lookup(
    sp_df: pd.DataFrame | None,
    material_col: str = DEFAULT_COL_MATERIAL,
    pack_col: str = "Delivery unit",
) -> dict[str, float]:
    """Returns {Material -> DeliveryUnit (Std Pack)}."""
    if sp_df is None or sp_df.empty:
        return {}
    df = sp_df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    m_col = _find_first_existing_col(df, [material_col, "Material"])
    p_col = _find_first_existing_col(df, [pack_col, "Delivery unit", "Min. MtO quantity", "Minimum delivery qty"])
    if m_col is None or p_col is None:
        return {}
    df["_mat"]  = df[m_col].astype(str).str.strip()
    df["_pack"] = df[p_col].apply(to_number)
    return df.drop_duplicates(subset=["_mat"]).set_index("_mat")["_pack"].to_dict()


def enrich_with_masters(
    df: pd.DataFrame,
    hrs_lookup: dict,
    cost_lookup: dict,
    sp_lookup: dict,
    part_col: str = DEFAULT_COL_PART,
    fallback_std_pack_col: str | None = DEFAULT_COL_STD_PACK,
) -> pd.DataFrame:
    """
    Left-joins Hrs, Cost, and Std Pack master data onto the orders DataFrame.
    Falls back to the existing Std Pack column in orders when the master is missing.
    """
    df = df.copy()
    part_key = df[part_col].astype(str).str.strip() if part_col in df.columns else pd.Series([""]*len(df))

    df["HrsPerUnit"] = part_key.map(hrs_lookup).fillna(0.0)
    df["Std Cost"]   = part_key.map(cost_lookup).fillna(0.0)

    # Std Pack: prefer master table; fall back to column in orders
    sp_from_master = part_key.map(sp_lookup)
    if fallback_std_pack_col and fallback_std_pack_col in df.columns:
        sp_fallback = df[fallback_std_pack_col].apply(to_number)
    else:
        sp_fallback = pd.Series([0.0] * len(df), index=df.index)
    df["Std Pack (Master)"] = sp_from_master.combine_first(sp_fallback).fillna(0.0)

    df["Total Hours"] = df["QtyNum"] * df["HrsPerUnit"]
    df["Total Value"] = df["QtyNum"] * df["Std Cost"]

    return df


def add_date_dimensions(df: pd.DataFrame, date_col: str = "PlanningDate") -> pd.DataFrame:
    """Adds Year, Quarter, Month, MonthName, ISOWeek columns from a date column."""
    df = df.copy()
    dates = pd.to_datetime(df[date_col], errors="coerce")
    df["Year"]      = dates.dt.year
    df["Quarter"]   = "Q" + dates.dt.quarter.astype(str)
    df["Month"]     = dates.dt.to_period("M").astype(str)
    df["Month Name"] = dates.dt.strftime("%b")
    df["ISOWeek"]   = dates.apply(lambda d: int(d.isocalendar()[1]) if pd.notna(d) else None)
    return df


def build_ma3_summary(
    result: pd.DataFrame,
    group_by: list[str] | None = None,
) -> dict[str, pd.DataFrame]:
    """
    Returns three DataFrames (MA3 Pieces, MA3 Hours, MA3 Sales/Value)
    pivoted by the requested dimension (default: Line Name × Month).
    """
    if result.empty:
        empty = pd.DataFrame()
        return {"Pieces": empty, "Hours": empty, "Value": empty}

    group_by = group_by or ["Line Name", "Quarter", "Month Name"]
    available = [c for c in group_by if c in result.columns]

    def _pivot(metric_col: str) -> pd.DataFrame:
        if metric_col not in result.columns:
            return pd.DataFrame()
        cols_needed = available + [metric_col]
        sub = result[cols_needed].copy()
        sub[metric_col] = pd.to_numeric(sub[metric_col], errors="coerce").fillna(0)
        if len(available) < 2:
            return sub.groupby(available)[metric_col].sum().reset_index()
        row_key = available[0]
        col_key = available[1:]
        # For pivot we use the last grouping level as columns
        pivot_col = col_key[-1]
        extra_rows = [row_key] + col_key[:-1]
        try:
            pvt = sub.pivot_table(
                index=extra_rows,
                columns=pivot_col,
                values=metric_col,
                aggfunc="sum",
                fill_value=0,
            )
            pvt["Total"] = pvt.sum(axis=1)
            return pvt.reset_index()
        except Exception:
            return sub.groupby(extra_rows + [pivot_col])[metric_col].sum().reset_index()

    return {
        "Pieces": _pivot("ScheduledQty"),
        "Hours":  _pivot("Total Hours"),
        "Value":  _pivot("Total Value"),
    }


# ── Capacity parsing / simulation ────────────────────────────────────────────

def _choose_capacity_line_column(
    df: pd.DataFrame,
    preferred_line_col: str | None = None,
    line_col_candidates: list[str] | None = None,
    valid_lines: set[str] | None = None,
) -> str:
    cols = list(df.columns)

    base_candidates = line_col_candidates or ["Line", "Line Grp", "Line Group", "LineGroup", "MRP"]
    candidates = []
    if preferred_line_col:
        candidates.append(preferred_line_col)
    candidates.extend(base_candidates)

    existing = []
    by_lower = {str(c).strip().lower(): c for c in cols}
    for c in candidates:
        col = by_lower.get(str(c).strip().lower())
        if col is not None and col not in existing:
            existing.append(col)

    if not existing:
        return cols[0]

    if not valid_lines:
        return existing[0]

    best_col = existing[0]
    best_score = -1
    for c in existing:
        vals = set(normalize_line(v) for v in df[c].tolist())
        vals.discard(None)
        score = len(vals.intersection(valid_lines))
        if score > best_score:
            best_col = c
            best_score = score

    return best_col


def build_cap_flat(
    capacity_df: pd.DataFrame | None,
    line_col_candidates: list[str] | None = None,
    preferred_line_col: str | None = None,
    valid_lines: set[str] | None = None,
) -> list[dict]:
    """
    Build long-form capacity rows from wide table:
      Line | 6/8/2026 | 6/15/2026 | ...
        Returns sorted rows with keys:
            _Line, _LineName, _WeekDate, _WkNum, _Cap
    """
    if capacity_df is None or capacity_df.empty:
        return []

    df = capacity_df.copy()

    first_row_vals = [str(v).strip() for v in df.iloc[0].tolist()]
    if any("line" in v.lower() for v in first_row_vals):
        df.columns = first_row_vals
        df = df.iloc[1:].reset_index(drop=True)

    cols = list(df.columns)
    line_col = _choose_capacity_line_column(
        df,
        preferred_line_col=preferred_line_col,
        line_col_candidates=(line_col_candidates or ["Line", "Line Grp", "Line Group", "LineGroup", "Production Line"]),
        valid_lines=valid_lines,
    )
    line_name_col = _find_first_existing_col(df, ["Line", "Line Name", "Line Description", "Description"])
    if line_name_col is None:
        line_name_col = line_col

    date_cols = [c for c in cols if c != line_col and to_date(c) is not None]

    rows = []
    for _, row in df.iterrows():
        line_key = normalize_line(row.get(line_col))
        if not line_key:
            continue
        line_name_raw = row.get(line_name_col)
        line_name = str(line_name_raw).strip() if line_name_raw is not None else ""
        if not line_name:
            line_name = line_key

        for dc in date_cols:
            cap = to_number(row.get(dc, 0))
            if cap <= 0:
                continue

            week_date = to_date(dc)
            if week_date is None:
                continue

            rows.append(
                {
                    "_Line": line_key,
                    "_LineName": line_name,
                    "_WeekDate": week_date,
                    "_WkNum": int(week_date.isocalendar()[1]),
                    "_Cap": cap,
                }
            )

    return sorted(rows, key=lambda r: (r["_Line"], r["_WeekDate"]))


def apply_capacity_overrides(
    cap_flat: list[dict],
    overrides_df: pd.DataFrame | None,
) -> list[dict]:
    """
    Overrides schema:
      Line | Week | New Capacity
    Only matching Line+Week rows are overridden; other weeks stay original.
    If a Line+Week does not exist and New Capacity > 0, it is inserted.
    """
    if not cap_flat:
        base = []
    else:
        base = [dict(r) for r in cap_flat]

    if overrides_df is None or overrides_df.empty:
        return sorted(base, key=lambda r: (r["_Line"], r["_WeekDate"]))

    ov = overrides_df.copy()
    line_col = _find_first_existing_col(ov, ["Line"])
    week_col = _find_first_existing_col(ov, ["Week", "Date"])
    cap_col = _find_first_existing_col(ov, ["New Capacity", "Capacity", "Override Capacity"])
    if line_col is None or week_col is None or cap_col is None:
        return sorted(base, key=lambda r: (r["_Line"], r["_WeekDate"]))

    index = {(r["_Line"], r["_WeekDate"]): i for i, r in enumerate(base)}

    for _, row in ov.iterrows():
        line_key = normalize_line(row.get(line_col))
        week_date = to_date(row.get(week_col))
        new_cap = to_number(row.get(cap_col))

        if not line_key or week_date is None:
            continue

        key = (line_key, week_date)
        if key in index:
            base[index[key]]["_Cap"] = max(0.0, new_cap)
        else:
            if new_cap > 0:
                base.append(
                    {
                        "_Line": line_key,
                        "_LineName": line_key,
                        "_WeekDate": week_date,
                        "_WkNum": int(week_date.isocalendar()[1]),
                        "_Cap": new_cap,
                    }
                )

    base = [r for r in base if r["_Cap"] > 0]
    return sorted(base, key=lambda r: (r["_Line"], r["_WeekDate"]))


def capacity_comparison(original_cap_flat: list[dict], adjusted_cap_flat: list[dict]) -> pd.DataFrame:
    o = pd.DataFrame(original_cap_flat)
    a = pd.DataFrame(adjusted_cap_flat)

    if o.empty:
        o = pd.DataFrame(columns=["_Line", "_LineName", "_WeekDate", "_Cap"])
    if a.empty:
        a = pd.DataFrame(columns=["_Line", "_LineName", "_WeekDate", "_Cap"])

    if "_LineName" not in o.columns:
        o["_LineName"] = o.get("_Line")
    if "_LineName" not in a.columns:
        a["_LineName"] = a.get("_Line")

    o = o.rename(columns={"_Cap": "Original capacity", "_LineName": "Original Line Name"})[["_Line", "Original Line Name", "_WeekDate", "Original capacity"]]
    a = a.rename(columns={"_Cap": "Adjusted capacity", "_LineName": "Adjusted Line Name"})[["_Line", "Adjusted Line Name", "_WeekDate", "Adjusted capacity"]]

    cmp_df = o.merge(a, on=["_Line", "_WeekDate"], how="outer")
    cmp_df["Original capacity"] = cmp_df["Original capacity"].fillna(0)
    cmp_df["Adjusted capacity"] = cmp_df["Adjusted capacity"].fillna(0)
    cmp_df["Line Name"] = cmp_df["Adjusted Line Name"].fillna(cmp_df["Original Line Name"]).fillna(cmp_df["_Line"])
    cmp_df["Delta"] = cmp_df["Adjusted capacity"] - cmp_df["Original capacity"]
    cmp_df = cmp_df.rename(columns={"_Line": "Line", "_WeekDate": "Week"})
    return cmp_df[["Line", "Line Name", "Week", "Original capacity", "Adjusted capacity", "Delta"]].sort_values(["Line", "Week"]).reset_index(drop=True)


# ── Scheduling core ──────────────────────────────────────────────────────────

def get_buckets(line_key: str, cap_flat: list[dict], base_week: int):
    rows = [r for r in cap_flat if r["_Line"] == line_key]
    if rows:
        rows = sorted(rows, key=lambda r: r["_WeekDate"])
        return [
            {
                "Week": r["_WkNum"],
                "WeekDate": r["_WeekDate"],
                "Cap": r["_Cap"],
                "LineName": r.get("_LineName", r["_Line"]),
            }
            for r in rows
        ]

    if line_key in DEFAULT_CAPS:
        fallback = float(DEFAULT_CAPS[line_key])
    elif cap_flat:
        # Prevent pathological very slow runs when line keys do not match and fallback=1.
        fallback = float(pd.Series([r["_Cap"] for r in cap_flat]).median())
    else:
        fallback = 1.0
    fallback = max(1.0, fallback)
    return [{"Week": base_week, "WeekDate": None, "Cap": fallback, "LineName": line_key}]


def find_idx(cum_pos, cum_cap_ends):
    for i, end in enumerate(cum_cap_ends):
        if end > cum_pos:
            return i
    return len(cum_cap_ends) - 1


def _process_group(grp, line_key, cap_flat, base_week):
    grp = grp.copy().reset_index(drop=True)
    qty_list = [to_number(q) for q in grp["QtyPlannedInput"]]

    raw_buckets = get_buckets(line_key, cap_flat, base_week)

    total_qty = sum(qty_list)
    last_cap = raw_buckets[-1]["Cap"]
    last_week = raw_buckets[-1]["Week"]
    last_week_date = raw_buckets[-1].get("WeekDate")

    raw_sum = sum(b["Cap"] for b in raw_buckets)
    deficit = max(0, total_qty - raw_sum)
    extra_n = math.ceil(deficit / last_cap) if last_cap > 0 else 0

    ext_buckets = raw_buckets + [
        {
            "Week": last_week + n,
            "WeekDate": (last_week_date + timedelta(days=7 * n)) if last_week_date is not None else None,
            "Cap": last_cap,
            "LineName": raw_buckets[-1].get("LineName", line_key),
        }
        for n in range(1, extra_n + 1)
    ]

    cum_cap_ends = []
    running = 0
    for b in ext_buckets:
        running += b["Cap"]
        cum_cap_ends.append(running)

    week_nums = [b["Week"] for b in ext_buckets]
    week_dates = [b["WeekDate"] for b in ext_buckets]

    output_rows = []

    for idx in range(len(grp)):
        row = grp.iloc[idx].to_dict()
        qty = qty_list[idx]

        cum_s = sum(qty_list[:idx])
        cum_e = cum_s + qty

        i_s = find_idx(cum_s, cum_cap_ends)
        i_e = find_idx(cum_e - 1, cum_cap_ends)

        for i in range(i_s, i_e + 1):
            bucket_start = 0 if i == 0 else cum_cap_ends[i - 1]
            bucket_end = cum_cap_ends[i]

            scheduled_qty = max(0, min(bucket_end, cum_e) - max(bucket_start, cum_s))

            output_rows.append(
                {
                    **row,
                    "Line": line_key,
                    "Line Name": ext_buckets[i].get("LineName", line_key),
                    "ProductionWeek": week_nums[i],
                    "ProductionWeekDate": week_dates[i],
                    "ScheduledQtyBase": scheduled_qty,
                    "SplitFlag": "SPLIT" if i_s != i_e else "",
                    # Keep excess only on first slice to avoid duplicate excess totals.
                    "Excess Std Pack": row.get("Excess Std Pack", 0) if i == i_s else 0,
                }
            )

    return pd.DataFrame(output_rows)


# ── Main ─────────────────────────────────────────────────────────────────────

def run_query(
    data_df,
    capacity_df=None,
    capacity_overrides_df=None,
    base_week=24,
    col_qty=DEFAULT_COL_QTY,
    col_line=DEFAULT_COL_LINE,
    col_req_date=DEFAULT_COL_REQ_DATE,
    col_commit_date=DEFAULT_COL_COMMIT_DATE,
    col_std_pack=DEFAULT_COL_STD_PACK,
    capacity_line_col="Line",
    hrs_df=None,
    cost_df=None,
    sp_df=None,
    col_part=DEFAULT_COL_PART,
):
    df = data_df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    df["QtyNum"] = df[col_qty].apply(to_number)
    df["Line"] = df[col_line].apply(normalize_line)

    # ── Enrich with master tables (Hrs, Cost, Std Pack) ──────────────────────
    hrs_lookup = build_hrs_lookup(hrs_df)
    cost_lookup = build_cost_lookup(cost_df)
    sp_lookup = build_std_pack_lookup(sp_df)
    df = enrich_with_masters(
        df,
        hrs_lookup=hrs_lookup,
        cost_lookup=cost_lookup,
        sp_lookup=sp_lookup,
        part_col=col_part if col_part in df.columns else DEFAULT_COL_PART,
        fallback_std_pack_col=col_std_pack,
    )

    # Std Pack source: prefer master lookup result, fall back to orders column
    std_pack_col = _find_first_existing_col(
        df, ["Std Pack (Master)", col_std_pack, "Std Pack", "StdPack", "STD PACK"]
    )
    if std_pack_col is None:
        df["StdPackNum"] = 0.0
    else:
        df["StdPackNum"] = df[std_pack_col].apply(to_number)

    std_pack_calc = df.apply(lambda r: _round_up_to_std_pack(r["QtyNum"], r["StdPackNum"]), axis=1)
    df["QtyPlannedInput"] = std_pack_calc.apply(lambda t: t[0])
    df["Excess Std Pack"] = std_pack_calc.apply(lambda t: t[1])

    def _plan_date(row):
        c = row.get(col_commit_date)
        r = row.get(col_req_date)
        val = c if (c is not None and not (isinstance(c, float) and math.isnan(c))) else r
        return to_date(val)

    df["PlanningDate"] = df.apply(_plan_date, axis=1)
    df = df.sort_values(["Line", "PlanningDate"], na_position="last").reset_index(drop=True)

    valid_lines = set(df["Line"].dropna().tolist())
    base_cap_flat = build_cap_flat(
        capacity_df,
        preferred_line_col=capacity_line_col,
        valid_lines=valid_lines,
    )
    cap_flat = apply_capacity_overrides(base_cap_flat, capacity_overrides_df)

    parts = []
    for line_key, grp in df.groupby("Line", sort=False):
        parts.append(_process_group(grp, line_key, cap_flat, base_week))

    result = pd.concat(parts, ignore_index=True) if parts else pd.DataFrame()

    if result.empty:
        return result

    result["ScheduledQty"] = result["ScheduledQtyBase"].apply(lambda q: int(math.floor(q)) if q else 0)

    # ── Compute per-slice metrics (proportional to scheduled qty) ────────────
    hrs_per_unit = result.get("HrsPerUnit", pd.Series([0.0] * len(result), index=result.index)).fillna(0.0)
    std_cost     = result.get("Std Cost",   pd.Series([0.0] * len(result), index=result.index)).fillna(0.0)
    result["Total Hours"] = result["ScheduledQty"] * hrs_per_unit
    result["Total Value"] = result["ScheduledQty"] * std_cost

    # ── Date dimensions ───────────────────────────────────────────────────────
    if "ProductionWeekDate" in result.columns:
        result = add_date_dimensions(result, date_col="ProductionWeekDate")

    # ── Column ordering ───────────────────────────────────────────────────────
    cols = list(result.columns)

    def _move_after(lst, item, after):
        if item in lst and after in lst:
            lst.remove(item)
            lst.insert(lst.index(after) + 1, item)

    _move_after(cols, "Excess Std Pack", "ScheduledQty")
    _move_after(cols, "Line Name",       "Line")
    _move_after(cols, "Total Hours",     "ScheduledQty")
    _move_after(cols, "Total Value",     "Total Hours")

    result = result[cols]
    return result
