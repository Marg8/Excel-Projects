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
    return t if t else None


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


# ── Capacity parsing / simulation ────────────────────────────────────────────

def build_cap_flat(
    capacity_df: pd.DataFrame | None,
    line_col_candidates: list[str] | None = None,
) -> list[dict]:
    """
    Build long-form capacity rows from wide table:
      Line | 6/8/2026 | 6/15/2026 | ...
    Returns sorted rows with keys:
      _Line, _WeekDate, _WkNum, _Cap
    """
    if capacity_df is None or capacity_df.empty:
        return []

    df = capacity_df.copy()

    first_row_vals = [str(v).strip() for v in df.iloc[0].tolist()]
    if any("line" in v.lower() for v in first_row_vals):
        df.columns = first_row_vals
        df = df.iloc[1:].reset_index(drop=True)

    cols = list(df.columns)
    line_col_candidates = line_col_candidates or ["Line", "Production Line"]
    line_col = _find_first_existing_col(df, line_col_candidates)
    if line_col is None:
        line_col = cols[0]

    date_cols = [c for c in cols if c != line_col and to_date(c) is not None]

    rows = []
    for _, row in df.iterrows():
        line_key = normalize_line(row.get(line_col))
        if not line_key:
            continue

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
        o = pd.DataFrame(columns=["_Line", "_WeekDate", "_Cap"])
    if a.empty:
        a = pd.DataFrame(columns=["_Line", "_WeekDate", "_Cap"])

    o = o.rename(columns={"_Cap": "Original capacity"})[["_Line", "_WeekDate", "Original capacity"]]
    a = a.rename(columns={"_Cap": "Adjusted capacity"})[["_Line", "_WeekDate", "Adjusted capacity"]]

    cmp_df = o.merge(a, on=["_Line", "_WeekDate"], how="outer")
    cmp_df["Original capacity"] = cmp_df["Original capacity"].fillna(0)
    cmp_df["Adjusted capacity"] = cmp_df["Adjusted capacity"].fillna(0)
    cmp_df["Delta"] = cmp_df["Adjusted capacity"] - cmp_df["Original capacity"]
    cmp_df = cmp_df.rename(columns={"_Line": "Line", "_WeekDate": "Week"})
    return cmp_df.sort_values(["Line", "Week"]).reset_index(drop=True)


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
            }
            for r in rows
        ]

    fallback = float(DEFAULT_CAPS.get(line_key, 1))
    return [{"Week": base_week, "WeekDate": None, "Cap": fallback}]


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
):
    df = data_df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    df["QtyNum"] = df[col_qty].apply(to_number)
    df["Line"] = df[col_line].apply(normalize_line)

    std_pack_col = _find_first_existing_col(df, [col_std_pack, "Std Pack", "StdPack", "STD PACK"])
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

    base_cap_flat = build_cap_flat(capacity_df)
    cap_flat = apply_capacity_overrides(base_cap_flat, capacity_overrides_df)

    parts = []
    for line_key, grp in df.groupby("Line", sort=False):
        parts.append(_process_group(grp, line_key, cap_flat, base_week))

    result = pd.concat(parts, ignore_index=True) if parts else pd.DataFrame()

    if result.empty:
        return result

    result["ScheduledQty"] = result["ScheduledQtyBase"].apply(lambda q: int(math.floor(q)) if q else 0)

    cols = list(result.columns)
    if "Excess Std Pack" in cols and "ScheduledQty" in cols:
        cols.remove("Excess Std Pack")
        idx_sched = cols.index("ScheduledQty") + 1
        cols.insert(idx_sched, "Excess Std Pack")

    result = result[cols]
    return result
