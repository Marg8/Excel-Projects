import math
import pandas as pd

# ── module-level constants (imported by app.py) ──────────────────────

DEFAULT_CAPS: dict[str, int] = {
    "HC8_H3H": 35000,
    "H4Z":      8250,
    "HB5":     32000,
    "H5C_HC9":  6100,
    "HC7":     19000,
    "H2J":      9500,
    "HC4_A9R": 35000,
    "HC0":     16000,
}

DEFAULT_COL_QTY         = "Qty"
DEFAULT_COL_MRP         = "MRP"
DEFAULT_COL_REQ_DATE    = "Requested date"
DEFAULT_COL_COMMIT_DATE = "Plan Request Date"

# ── helpers ─────────────────────────────────────────────────────────

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
        return x
    try:
        return pd.to_datetime(x)
    except Exception:
        return None

def normalize_group(mrp) -> str | None:
    if mrp is None:
        return None
    t = str(mrp).strip().upper()
    if not t:
        return None
    if t in {"HC8", "H3H"}:
        return "HC8_H3H"
    if t in {"H5C", "HC9"}:
        return "H5C_HC9"
    if t in {"HC4", "A9R"}:
        return "HC4_A9R"
    return t

# ── Capacity parsing ────────────────────────────────────────────────

def build_cap_flat(capacity_df: pd.DataFrame | None) -> list[dict]:
    if capacity_df is None or capacity_df.empty:
        return []

    df = capacity_df.copy()

    first_row_vals = [str(v).strip() for v in df.iloc[0].tolist()]
    if any("group" in v.lower() for v in first_row_vals):
        df.columns = first_row_vals
        df = df.iloc[1:].reset_index(drop=True)

    cols = list(df.columns)

    group_col = next(
        (c for c in cols if str(c).strip().lower() == "line group"),
        next((c for c in cols if "group" in str(c).lower()), cols[0]),
    )

    date_cols = [c for c in cols if c != group_col and to_date(c) is not None]

    rows = []
    for _, row in df.iterrows():
        lg = normalize_group(row.get(group_col))
        if not lg:
            continue

        line_val = row.get("Line")  # ✅ columna adicional

        for dc in date_cols:
            cap = to_number(row.get(dc, 0))
            if cap <= 0:
                continue

            d = to_date(dc)
            if d is None:
                continue

            wk = int(d.isocalendar()[1])

            rows.append({
                "_LG": lg,
                "_WkNum": wk,
                "_Cap": cap,
                "_Line": line_val  # ✅ guardar Line
            })

    return sorted(rows, key=lambda r: (r["_LG"], r["_WkNum"]))

# ── Buckets ─────────────────────────────────────────────────────────

def get_buckets(lg, cap_flat, base_week):
    buckets = [r for r in cap_flat if r["_LG"] == lg]
    if buckets:
        return sorted(
            [{"Week": b["_WkNum"], "Cap": b["_Cap"], "_Line": b.get("_Line")} for b in buckets],
            key=lambda b: b["Week"]
        )
    return [{"Week": base_week, "Cap": 1, "_Line": None}]

def find_idx(cum_pos, cum_cap_ends):
    for i, end in enumerate(cum_cap_ends):
        if end > cum_pos:
            return i
    return len(cum_cap_ends) - 1

def _find_first_existing_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    by_lower = {str(c).strip().lower(): c for c in df.columns}
    for name in candidates:
        col = by_lower.get(name.strip().lower())
        if col is not None:
            return col
    return None

def _apply_std_pack_multiple_by_part(
    result: pd.DataFrame,
    part_candidates: list[str] | None = None,
    std_pack_candidates: list[str] | None = None,
) -> pd.DataFrame:
    if result.empty:
        result["ExcessCPSAdded"] = []
        return result

    part_candidates = part_candidates or ["Product code", "Part Number", "Part"]
    std_pack_candidates = std_pack_candidates or ["Std Pack", "StdPack", "STD PACK"]

    part_col = _find_first_existing_col(result, part_candidates)
    std_pack_col = _find_first_existing_col(result, std_pack_candidates)

    result["ExcessCPSAdded"] = 0

    # If either column is missing, keep the current schedule untouched.
    if part_col is None or std_pack_col is None:
        return result

    # Adjust each part total to the next std-pack multiple and place excess on last row.
    for part_val, idxs in result.groupby(part_col, sort=False).groups.items():
        part_rows = result.loc[idxs]
        std_pack_values = [int(to_number(v)) for v in part_rows[std_pack_col].tolist() if to_number(v) > 0]
        if not std_pack_values:
            continue

        std_pack = std_pack_values[0]
        if std_pack <= 0:
            continue

        current_total = int(part_rows["ScheduledQty"].sum())
        adjusted_total = int(math.ceil(current_total / std_pack) * std_pack)
        excess = max(0, adjusted_total - current_total)
        if excess == 0:
            continue

        last_idx = part_rows.index[-1]
        result.at[last_idx, "ScheduledQty"] = int(result.at[last_idx, "ScheduledQty"]) + excess
        result.at[last_idx, "ExcessCPSAdded"] = excess

    return result

# ── Core logic ──────────────────────────────────────────────────────

def _process_group(grp, line_group, cap_flat, base_week):
    grp = grp.copy().reset_index(drop=True)
    qty_list = [to_number(q) for q in grp["QtyNum"]]

    raw_buckets = get_buckets(line_group, cap_flat, base_week)

    total_qty = sum(qty_list)
    last_cap = raw_buckets[-1]["Cap"]
    last_week = raw_buckets[-1]["Week"]

    raw_sum = sum(b["Cap"] for b in raw_buckets)
    deficit = max(0, total_qty - raw_sum)
    extra_n = math.ceil(deficit / last_cap) if last_cap > 0 else 0

    ext_buckets = raw_buckets + [
        {"Week": last_week + n, "Cap": last_cap, "_Line": raw_buckets[-1].get("_Line")}
        for n in range(1, extra_n + 1)
    ]

    cum_cap_ends = []
    running = 0
    for b in ext_buckets:
        running += b["Cap"]
        cum_cap_ends.append(running)

    week_nums = [b["Week"] for b in ext_buckets]

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

            scheduled_qty = max(
                0,
                min(bucket_end, cum_e) - max(bucket_start, cum_s)
            )

            line_value = ext_buckets[i].get("_Line")  # ✅ rescatar Line

            output_rows.append({
                **row,
                "LineGroup": line_group,
                "Line": line_value,  # ✅ NUEVA COLUMNA
                "ProductionWeek": week_nums[i],
                "ScheduledQtyBase": scheduled_qty,
                "SplitFlag": "SPLIT" if i_s != i_e else ""
            })

    return pd.DataFrame(output_rows)

# ── Main ────────────────────────────────────────────────────────────

def run_query(
    data_df,
    capacity_df=None,
    base_week=24,
    col_qty=DEFAULT_COL_QTY,
    col_mrp=DEFAULT_COL_MRP,
    col_req_date=DEFAULT_COL_REQ_DATE,
    col_commit_date=DEFAULT_COL_COMMIT_DATE,
):

    df = data_df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    df["QtyNum"]    = df[col_qty].apply(to_number)
    df["LineGroup"] = df[col_mrp].apply(normalize_group)

    # PlanningDate: CommitDate first, fall back to Req. Date
    def _plan_date(row):
        c = row.get(col_commit_date)
        r = row.get(col_req_date)
        val = c if (c is not None and not (isinstance(c, float) and math.isnan(c))) else r
        return to_date(val)

    df["PlanningDate"] = df.apply(_plan_date, axis=1)

    df = df.sort_values(["LineGroup", "PlanningDate"], na_position="last").reset_index(drop=True)

    cap_flat = build_cap_flat(capacity_df)

    parts = []
    for lg, grp in df.groupby("LineGroup", sort=False):
        parts.append(_process_group(grp, lg, cap_flat, base_week))

    result = pd.concat(parts, ignore_index=True)

    result["ScheduledQty"] = result["ScheduledQtyBase"].apply(
        lambda q: int(math.floor(q)) if q else 0
    )

    result = _apply_std_pack_multiple_by_part(result)

    # ✅ ORDENAR columnas (Line después de LineGroup)
    cols = list(result.columns)
    if "ExcessCPSAdded" in cols and "ScheduledQty" in cols:
        cols.remove("ExcessCPSAdded")
        idx_sched = cols.index("ScheduledQty") + 1
        cols.insert(idx_sched, "ExcessCPSAdded")

    if "Line" in cols and "LineGroup" in cols:
        cols.remove("Line")
        idx = cols.index("LineGroup") + 1
        cols.insert(idx, "Line")

    result = result[cols]

    return result