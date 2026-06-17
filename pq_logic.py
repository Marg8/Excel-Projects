"""
pq_logic.py
Python port of the Power Query M code in Excel.xlsx.

Column mapping (matches M code):
  Data sheet   : "Order quantity", "MRP", "Req. Date", "Commit Date"
  Capacity sheet: "Line Group" + weekly date column headers

Algorithm (identical to M code):
  - NormalizeGroup  → HC8/H3H→HC8_H3H, H5C/HC9→H5C_HC9, HC4/A9R→HC4_A9R
  - CapFlat         → reads wide Capacity table, unpivots to (_LG, _WkNum, _Cap)
  - GetBuckets      → per-group sorted {Week, Cap} list with auto-extension
  - PlanningDate    → CommitDate if set, else Req. Date
  - CumStart/CumEnd → cumulative qty per row inside each group
  - FindIdx (iS/iE) → first bucket where CumCaps[i] > cumS / >= cumE
  - Slices          → min(bucketEnd,cumE) - max(bucketStart,cumS)
  - Final           → Number.RoundDown (integer floor, no StdPack logic)
"""

import math
import pandas as pd


# ── helpers ───────────────────────────────────────────────────────────────────

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
    """Try to parse x as a date; return pd.Timestamp or None."""
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

# Default column names — match the M code
DEFAULT_COL_QTY         = "Qty"
DEFAULT_COL_MRP         = "MRP"
DEFAULT_COL_REQ_DATE    = "Requested date"
DEFAULT_COL_COMMIT_DATE = "Plan Request Date"


def default_cap(lg: str | None) -> int | None:
    return DEFAULT_CAPS.get(lg) if lg else None


# ── CapFlat: parse Capacity table → list of {_LG, _WkNum, _Cap} ──────────────

def build_cap_flat(capacity_df: pd.DataFrame | None) -> list[dict]:
    """
    Reads a wide-format Capacity table:
        Line Group | 6/8/2026 | 6/15/2026 | ...
    Returns sorted list of dicts: {_LG, _WkNum, _Cap}
    """
    if capacity_df is None or capacity_df.empty:
        return []

    # If the first row looks like a real header (all-NaN row was skipped externally,
    # or the table has Unnamed cols), detect and skip it automatically.
    df = capacity_df.copy()
    # If columns are all numeric/Unnamed but row 0 contains "Line Group" text,
    # the sheet was read without the correct header offset — promote row 0.
    first_row_vals = [str(v).strip() for v in df.iloc[0].tolist()]
    if any("group" in v.lower() or v.lower() == "line" for v in first_row_vals):
        df.columns = first_row_vals
        df = df.iloc[1:].reset_index(drop=True)

    cols = list(df.columns)

    # Find the Line Group column — exact "line group" match preferred
    # (mirrors M: Table.UnpivotOtherColumns with {"Line Group"})
    group_col = next(
        (c for c in cols if str(c).strip().lower() == "line group"),
        next((c for c in cols if "group" in str(c).lower()), cols[0]),
    )

    # Date columns: every column except group_col whose name parses as a date
    date_cols = []
    for c in cols:
        if c == group_col:
            continue
        d = to_date(c)
        if d is not None:
            date_cols.append(c)

    if not date_cols:
        return []

    rows = []
    for _, row in capacity_df.iterrows():
        lg = normalize_group(row.get(group_col))
        if not lg:
            continue
        for dc in date_cols:
            cap = to_number(row.get(dc, 0))
            if cap <= 0:
                continue
            d = to_date(dc)
            if d is None:
                continue
            wk = int(d.isocalendar()[1])   # ISO week number, matches M's Date.WeekOfYear
            rows.append({"_LG": lg, "_WkNum": wk, "_Cap": cap})

    return sorted(rows, key=lambda r: (r["_LG"], r["_WkNum"]))


# ── GetBuckets: {Week, Cap} list for one LineGroup ────────────────────────────

def get_buckets(lg: str | None, cap_flat: list[dict], base_week: int) -> list[dict]:
    """
    Returns sorted [{Week, Cap}, ...] for lg.
    Falls back to a single DefaultCap bucket at base_week when table has no data.
    """
    buckets = [
        {"Week": r["_WkNum"], "Cap": r["_Cap"]}
        for r in cap_flat if r["_LG"] == lg
    ]
    if buckets:
        return sorted(buckets, key=lambda b: b["Week"])

    dc = default_cap(lg)
    return [{"Week": base_week, "Cap": dc if dc else 1}]


# ── FindIdx: mirrors BaseWeek + IntegerDivide(cumPos, cap) ───────────────────

def find_idx(cum_pos: float, cum_cap_ends: list[float]) -> int:
    """First bucket index i where cum_cap_ends[i] > cum_pos."""
    for i, end in enumerate(cum_cap_ends):
        if end > cum_pos:
            return i
    return len(cum_cap_ends) - 1   # fallback: last bucket


# ── Process one LineGroup (mirrors M's Grouped inner let block) ───────────────

def _process_group(
    grp: pd.DataFrame,
    line_group: str | None,
    cap_flat: list[dict],
    base_week: int,
) -> pd.DataFrame:

    grp = grp.copy().reset_index(drop=True)
    qty_list = [to_number(q) for q in grp["QtyNum"]]
    total_qty = sum(qty_list)

    # Build and extend bucket list until cumulative cap ≥ TotalQty
    raw_buckets = get_buckets(line_group, cap_flat, base_week)
    last_cap  = raw_buckets[-1]["Cap"]
    last_week = raw_buckets[-1]["Week"]
    raw_sum   = sum(b["Cap"] for b in raw_buckets)
    deficit   = max(0.0, total_qty - raw_sum)
    extra_n   = math.ceil(deficit / last_cap) if last_cap > 0 else 0
    ext_buckets = raw_buckets + [
        {"Week": last_week + n, "Cap": last_cap} for n in range(1, extra_n + 1)
    ]

    # Cumulative capacity bucket endings
    cum_cap_ends: list[float] = []
    running = 0.0
    for b in ext_buckets:
        running += b["Cap"]
        cum_cap_ends.append(running)
    week_nums = [b["Week"] for b in ext_buckets]

    # Process each row
    output_rows = []
    for idx in range(len(grp)):
        row = grp.iloc[idx].to_dict()
        qty   = qty_list[idx]
        cum_s = sum(qty_list[:idx])
        cum_e = cum_s + qty

        if qty == 0:
            output_rows.append({
                **row,
                "ProductionWeek":    week_nums[0],
                "ScheduledQtyBase": 0,
                "SplitFlag":         "",
            })
        else:
            i_s = find_idx(cum_s, cum_cap_ends)
            i_e = find_idx(cum_e - 1, cum_cap_ends)
            w_s = week_nums[i_s]
            w_e = week_nums[i_e]
            is_split = w_s != w_e

            for i in range(i_s, i_e + 1):
                bucket_start  = 0.0 if i == 0 else cum_cap_ends[i - 1]
                bucket_end    = cum_cap_ends[i]
                scheduled_qty = max(
                    0.0,
                    min(bucket_end, cum_e) - max(bucket_start, cum_s)
                )
                output_rows.append({
                    **row,
                    "ProductionWeek":    week_nums[i],
                    "ScheduledQtyBase": scheduled_qty,
                    "SplitFlag":         "SPLIT" if is_split else "",
                })

    return pd.DataFrame(output_rows)


# ── Main query function (mirrors the full M let block) ────────────────────────

def run_query(
    data_df: pd.DataFrame,
    capacity_df: pd.DataFrame | None = None,
    base_week: int = 24,
    col_qty: str = DEFAULT_COL_QTY,
    col_mrp: str = DEFAULT_COL_MRP,
    col_req_date: str = DEFAULT_COL_REQ_DATE,
    col_commit_date: str = DEFAULT_COL_COMMIT_DATE,
) -> pd.DataFrame:
    """
    Runs the full scheduling query.

    Parameters
    ----------
    data_df         : the Data sheet as a DataFrame
    capacity_df     : the Capacity sheet as a DataFrame (None → use default caps)
    base_week       : starting ISO week number (default 24)
    col_qty         : order quantity column name
    col_mrp         : MRP / line code column name
    col_req_date    : Requested date column name
    col_commit_date : Commit date column name (used first; falls back to col_req_date)

    Returns
    -------
    DataFrame with original columns plus:
        LineGroup, PlanningDate, ProductionWeek, ScheduledQtyBase,
        SplitFlag, ScheduledQty
    """
    df = data_df.copy()
    # Trim column names (mirrors M's Table.TransformColumnNames)
    df.columns = [str(c).strip() for c in df.columns]

    # QtyNum
    df["QtyNum"]   = df[col_qty].apply(to_number)

    # LineGroup (with Text.Upper + Text.Trim normalization)
    df["LineGroup"] = df[col_mrp].apply(normalize_group)

    # PlanningDate: CommitDate first, fall back to Req. Date
    def planning_date(row):
        c = row.get(col_commit_date)
        r = row.get(col_req_date)
        val = c if (c is not None and not (isinstance(c, float) and math.isnan(c))) else r
        d = to_date(val)
        return d

    df["PlanningDate"] = df.apply(planning_date, axis=1)

    # Sort (mirrors M's Table.Sort)
    df = df.sort_values(
        ["LineGroup", "PlanningDate"], na_position="last"
    ).reset_index(drop=True)

    # Build capacity flat table
    cap_flat = build_cap_flat(capacity_df)

    # Group and process
    parts = []
    for lg, grp in df.groupby("LineGroup", sort=False, dropna=False):
        part = _process_group(grp, lg, cap_flat, base_week)
        parts.append(part)

    if not parts:
        return pd.DataFrame()

    result = pd.concat(parts, ignore_index=True)

    # Final: Number.RoundDown (integer floor — no StdPack logic)
    result["ScheduledQty"] = result["ScheduledQtyBase"].apply(
        lambda q: int(math.floor(q)) if q else 0
    )

    # Remove helper column
    result = result.drop(columns=["QtyNum"], errors="ignore")

    return result
