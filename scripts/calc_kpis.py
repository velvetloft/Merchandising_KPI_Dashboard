import sqlite3
import pandas as pd
import json
import os
from datetime import datetime, date, timedelta

DB_PATH = r"C:\kpi_dashboard_local\kpi_history.db"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "..", "docs")
OUTPUT_PATH = os.path.join(OUTPUT_DIR, "dashboard_data.json")

TODAY = date.today()
LOOKBACK_START = (TODAY - timedelta(days=225)).isoformat()
LOOKBACK_END = TODAY.isoformat()
DAY_LIST_7 = [TODAY - timedelta(days=i + 1) for i in range(7)]

def iso_week_label(d):
    y, w, _ = d.isocalendar()
    return y * 100 + w

WEEK_LABELS = [iso_week_label(TODAY - timedelta(days=7 * i)) for i in range(1, 17)]
MONTH_KEYS = [
    f"{(TODAY.replace(day=1) - pd.DateOffset(months=i)).year:04d}-"
    f"{(TODAY.replace(day=1) - pd.DateOffset(months=i)).month:02d}"
    for i in range(1, 4)
]

def classify_a(p):
    if p is None or pd.isna(p): return "na"
    v = p * 100 if p <= 1 else p
    return "green" if v > 94.99 else ("yellow" if v >= 87.99 else "orange")

def classify_b(p):
    if p is None or pd.isna(p): return "na"
    v = p * 100 if p <= 1 else p
    return "green" if v > 87.99 else ("yellow" if v >= 74.99 else "orange")

def classify_moh(actual, target):
    if actual is None or target is None or pd.isna(actual) or pd.isna(target) or target == 0:
        return "na"
    var = (actual - target) / target
    if -0.10 <= var <= 0.05: return "green"
    if (-0.15 < var < -0.10) or (0.05 < var < 0.10): return "yellow"
    return "red"


# ============================================================
# LOAD DATA
# ============================================================
print("Loading data...")
conn = sqlite3.connect(DB_PATH)

# SPOC mapping - the source of truth for store universe
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
spoc_raw = pd.read_excel(os.path.join(PROJECT_DIR, "spoc_name.xlsx"))
spoc_raw.columns = ["spoc_code", "store_name", "spoc_name"]
SPOC_UNIVERSE = set(spoc_raw["spoc_code"].tolist())
SPOC_MAP = dict(zip(spoc_raw["spoc_code"], spoc_raw["spoc_name"]))
STORE_NAME_MAP = dict(zip(spoc_raw["spoc_code"], spoc_raw["store_name"]))

# Fill blank SPOC names
SPOC_MAP = {k: (v if pd.notna(v) and str(v).strip() else "Unassigned") for k, v in SPOC_MAP.items()}

# MOH targets
try:
    moh_t = pd.read_excel(os.path.join(PROJECT_DIR, "Moh_targets.xlsx"))
    moh_t.columns = ["store_code", "store_name", "moh_target"]
    MOH_TARGETS = dict(zip(moh_t["store_code"], moh_t["moh_target"]))
except Exception as e:
    print(f"  MOH targets not loaded: {e}")
    MOH_TARGETS = {}

STORE_NAMES = {k: STORE_NAME_MAP.get(k, k) for k in SPOC_UNIVERSE}
ALL_STORES = sorted(SPOC_UNIVERSE)
print(f"  Store universe: {len(ALL_STORES)} SPOC-mapped stores")

ALL_STOCK_DATES = pd.read_sql(
    f"SELECT DISTINCT snap_date FROM stock_snapshot WHERE snap_date >= '{LOOKBACK_START}' ORDER BY snap_date", conn
)["snap_date"].tolist()
print(f"  Stock dates: {len(ALL_STOCK_DATES)}")


# ============================================================
# A/B CLASS RANKING (single SQL)
# ============================================================
print("Computing rankings...")
all_ranks = pd.read_sql("""
WITH daily_agg AS (
    SELECT snap_date, product_code,
        SUM(sales_qty) AS qty, SUM(sales_value) AS amt
    FROM sales_daily
    WHERE snap_date >= date(?, '-60 days') AND snap_date <= ?
      AND category NOT IN ('Decoration Materials','Consumables','Consumable')
    GROUP BY snap_date, product_code
),
rolling AS (
    SELECT a1.snap_date AS report_date, a1.product_code,
        SUM(a2.qty) AS qty_60d, SUM(a2.amt) AS amt_60d
    FROM daily_agg a1
    JOIN daily_agg a2 ON a1.product_code = a2.product_code
        AND a2.snap_date > date(a1.snap_date, '-59 days')
        AND a2.snap_date <= a1.snap_date
    GROUP BY a1.snap_date, a1.product_code
),
ranked AS (
    SELECT report_date, product_code, qty_60d, amt_60d,
        ROW_NUMBER() OVER (PARTITION BY report_date ORDER BY qty_60d DESC, product_code) AS qr,
        ROW_NUMBER() OVER (PARTITION BY report_date ORDER BY amt_60d DESC, product_code) AS ar
    FROM rolling
)
SELECT report_date, product_code,
    ROW_NUMBER() OVER (PARTITION BY report_date ORDER BY (qr+ar)/2.0, product_code) AS final_rank
FROM ranked
""", conn, params=(LOOKBACK_START, LOOKBACK_END))
print(f"  {len(all_ranks)} ranking rows")


# ============================================================
# STOCK SNAPSHOTS (SPOC stores only)
# ============================================================
print("Loading stock...")
stock_df = pd.read_sql(f"""
    SELECT snap_date, store_code, product_code, available_stock
    FROM stock_snapshot
    WHERE snap_date >= '{LOOKBACK_START}' AND snap_date <= '{LOOKBACK_END}'
      AND store_code IN ({','.join('?' for _ in ALL_STORES)})
""", conn, params=tuple(ALL_STORES))

stock_idx = {}
for row in stock_df.itertuples(index=False):
    key = (row.snap_date, row.store_code)
    if key not in stock_idx:
        stock_idx[key] = {}
    stock_idx[key][row.product_code] = row.available_stock
del stock_df


# ============================================================
# MANIFEST (MOS) with fallback
# ============================================================
print("Loading manifest...")
manifest_df = pd.read_sql("SELECT DISTINCT manifest_date, product_code FROM ordering_manifest", conn)
manifest_sku = {}
for row in manifest_df.itertuples(index=False):
    manifest_sku.setdefault(row.manifest_date, set()).add(row.product_code)
del manifest_df
manifest_dates_sorted = sorted(manifest_sku.keys())

def get_mos(target_date):
    for i in range(len(manifest_dates_sorted) - 1, -1, -1):
        if manifest_dates_sorted[i] <= target_date:
            return manifest_sku[manifest_dates_sorted[i]]
    return set()


# ============================================================
# NEW SKU LAUNCH DATES
# ============================================================
launch_df = pd.read_sql(
    "SELECT product_code, b2b_launch_date FROM product_master WHERE b2b_launch_date IS NOT NULL", conn
)
launch_df["ld"] = pd.to_datetime(launch_df["b2b_launch_date"], errors="coerce").dt.date
launch_df = launch_df.dropna(subset=["ld"])


# ============================================================
# AVAILABILITY (SPOC stores only)
# ============================================================
def compute_availability(eligible_fn):
    rows = []
    for d_str in ALL_STOCK_DATES:
        eligible = eligible_fn(d_str)
        if not eligible:
            continue
        mos = get_mos(d_str)
        in_mos_set = eligible & mos
        for store in ALL_STORES:
            stock = stock_idx.get((d_str, store), {})
            numer = sum(1 for p in eligible if stock.get(p, 0) >= 3)
            oos_not_mos = sum(1 for p in eligible if stock.get(p, 0) < 3 and p not in in_mos_set)
            denom = len(eligible) - oos_not_mos
            rows.append({"store_code": store, "snap_date": d_str, "value": numer / denom if denom > 0 else 0.0})
    return pd.DataFrame(rows) if rows else pd.DataFrame(columns=["store_code", "snap_date", "value"])

print("Computing A-Class...")
a_sets = {d: set(g.loc[g["final_rank"].between(1, 300), "product_code"]) for d, g in all_ranks.groupby("report_date")}
a_daily = compute_availability(lambda d: a_sets.get(d, set()))

print("Computing B-Class...")
b_sets = {d: set(g.loc[g["final_rank"].between(301, 600), "product_code"]) for d, g in all_ranks.groupby("report_date")}
b_daily = compute_availability(lambda d: b_sets.get(d, set()))

print("Computing New SKU...")
new_daily = compute_availability(lambda d: set(launch_df.loc[
    (launch_df["ld"] >= datetime.strptime(d, "%Y-%m-%d").date() - timedelta(days=30)) &
    (launch_df["ld"] <= datetime.strptime(d, "%Y-%m-%d").date())
]["product_code"]))


# ============================================================
# MOH (SPOC stores only)
# ============================================================
print("Computing MOH...")
stock_moh = pd.read_sql(f"""
    SELECT snap_date, store_code, SUM(MAX(total_stock, 0)) AS stock_qty
    FROM stock_snapshot
    WHERE snap_date >= '{LOOKBACK_START}' AND snap_date <= '{LOOKBACK_END}'
      AND category NOT IN ('Decoration Materials','Consumables','Consumable')
      AND store_code IN ({','.join('?' for _ in ALL_STORES)})
    GROUP BY snap_date, store_code
""", conn, params=tuple(ALL_STORES))

sales_moh = pd.read_sql(f"""
    SELECT snap_date, store_code, SUM(sales_qty) AS sales_qty
    FROM sales_daily
    WHERE snap_date >= '{LOOKBACK_START}' AND snap_date <= '{LOOKBACK_END}'
      AND category NOT IN ('Decoration Materials','Consumables','Consumable')
      AND store_code IN ({','.join('?' for _ in ALL_STORES)})
    GROUP BY snap_date, store_code
""", conn, params=tuple(ALL_STORES))

moh_rows = []
for d_str in ALL_STOCK_DATES:
    d_date = datetime.strptime(d_str, "%Y-%m-%d").date()
    stk = stock_moh[stock_moh["snap_date"] == d_str][["store_code", "stock_qty"]]
    cutoff = (d_date - timedelta(days=29)).isoformat()
    sls = sales_moh[(sales_moh["snap_date"] <= d_str) & (sales_moh["snap_date"] >= cutoff)]
    sls_agg = sls.groupby("store_code", as_index=False)["sales_qty"].sum()
    m = pd.merge(stk, sls_agg, on="store_code", how="left")
    m["sales_qty"] = m["sales_qty"].fillna(0)
    m["value"] = m.apply(lambda r: r["stock_qty"] / (r["sales_qty"] if r["sales_qty"] > 0 else 1), axis=1)
    m["snap_date"] = d_str
    moh_rows.append(m[["store_code", "snap_date", "value"]])
moh_daily = pd.concat(moh_rows, ignore_index=True)

conn.close()


# ============================================================
# WIDE FORMAT + AGGREGATION
# ============================================================
def build_wide(df, rnd=4):
    if df.empty: return {}
    df = df.copy()
    df["snap_date"] = pd.to_datetime(df["snap_date"]).dt.date
    out = {}
    for store, g in df.groupby("store_code"):
        s = g.set_index("snap_date")["value"]
        rec = {}
        for i, d in enumerate(DAY_LIST_7, 1):
            v = s.get(d)
            rec[f"d{i}"] = None if v is None or pd.isna(v) else round(float(v), rnd)
        for i, wl in enumerate(WEEK_LABELS[:4], 1):
            vals = [v for d, v in s.items() if iso_week_label(d) == wl and pd.notna(v)]
            rec[f"w{i}"] = round(sum(vals)/len(vals), rnd) if vals else None
        for i, mk in enumerate(MONTH_KEYS[:3], 1):
            vals = [v for d, v in s.items() if f"{d.year:04d}-{d.month:02d}" == mk and pd.notna(v)]
            rec[f"m{i}"] = round(sum(vals)/len(vals), rnd) if vals else None
        out[store] = rec
    return out


def agg_spoc(wide, exclude_extreme=False):
    groups = {}
    for store, rec in wide.items():
        groups.setdefault(SPOC_MAP.get(store, "Unassigned"), {})[store] = rec
    out = {}
    for spoc, stores in groups.items():
        out[spoc] = {}
        keys = set().union(*(r.keys() for r in stores.values()))
        for k in keys:
            vals = [r[k] for r in stores.values() if k in r and r[k] is not None]
            if exclude_extreme:
                vals = [v for v in vals if v < 50]
            out[spoc][k] = round(sum(vals)/len(vals), 4) if vals else None
    return out


def agg_miniso(wide, exclude_extreme=False):
    if not wide: return {}
    keys = set().union(*(r.keys() for r in wide.values()))
    result = {}
    for k in keys:
        vals = [r[k] for r in wide.values() if k in r and r[k] is not None]
        if exclude_extreme:
            vals = [v for v in vals if v < 50]
        result[k] = round(sum(vals)/max(1, len(vals)), 4)
    return result


# ============================================================
# BUILD OUTPUT
# ============================================================
print("Building output...")
a_w = build_wide(a_daily)
b_w = build_wide(b_daily)
n_w = build_wide(new_daily)
m_w = build_wide(moh_daily, 2)

a_s = agg_spoc(a_w)
b_s = agg_spoc(b_w)
n_s = agg_spoc(n_w)
m_s = agg_spoc(m_w, exclude_extreme=True)

a_mi = agg_miniso(a_w)
b_mi = agg_miniso(b_w)
n_mi = agg_miniso(n_w)
m_mi = agg_miniso(m_w, exclude_extreme=True)

moh_tgt_avg = round(sum(MOH_TARGETS.values())/max(1, len(MOH_TARGETS)), 2)


def build_store_row(store, wide, clf, is_moh=False):
    rec = wide.get(store, {})
    row = {
        "store_code": store,
        "store_name": STORE_NAMES.get(store, store),
        "spoc_name": SPOC_MAP.get(store, ""),
        "metrics": rec,
        "status": "na",
    }
    if is_moh:
        d1 = rec.get("d1")
        row["target"] = MOH_TARGETS.get(store)
        row["variance"] = None
        row["non_trading"] = d1 is not None and d1 > 50
        if d1 is not None and MOH_TARGETS.get(store):
            row["variance"] = round((d1 - MOH_TARGETS[store]) / MOH_TARGETS[store], 4)
        row["status"] = classify_moh(d1, MOH_TARGETS.get(store))
    else:
        row["status"] = clf(rec.get("d1")) if clf else "na"
    return row


output = {
    "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    "data_freshness": {"latest_stock_date": ALL_STOCK_DATES[-1] if ALL_STOCK_DATES else None},
    "store_count": len(ALL_STORES),
    "kpis": {
        "availability_a": {
            "label": "A-Class Availability",
            "thresholds": {"green": "> 94.99%", "yellow": "87.99% - 94.99%", "orange": "< 87.99%"},
            "miniso": {"metrics": a_mi, "status": classify_a(a_mi.get("d1"))},
            "spocs": [
                {"spoc_name": sp, "store_count": len([s for s, v in SPOC_MAP.items() if v == sp]),
                 "metrics": rec, "status": classify_a(rec.get("d1"))}
                for sp, rec in sorted(a_s.items())
            ],
            "stores": [build_store_row(s, a_w, classify_a) for s in ALL_STORES],
        },
        "availability_b": {
            "label": "B-Class Availability",
            "thresholds": {"green": "> 87.99%", "yellow": "74.99% - 87.99%", "orange": "< 74.99%"},
            "miniso": {"metrics": b_mi, "status": classify_b(b_mi.get("d1"))},
            "spocs": [
                {"spoc_name": sp, "store_count": len([s for s, v in SPOC_MAP.items() if v == sp]),
                 "metrics": rec, "status": classify_b(rec.get("d1"))}
                for sp, rec in sorted(b_s.items())
            ],
            "stores": [build_store_row(s, b_w, classify_b) for s in ALL_STORES],
        },
        "new_sku_availability": {
            "label": "New SKU Availability",
            "thresholds": {"green": "> 94.99%", "yellow": "87.99% - 94.99%", "orange": "< 87.99%"},
            "miniso": {"metrics": n_mi, "status": classify_a(n_mi.get("d1"))},
            "spocs": [
                {"spoc_name": sp, "store_count": len([s for s, v in SPOC_MAP.items() if v == sp]),
                 "metrics": rec, "status": classify_a(rec.get("d1"))}
                for sp, rec in sorted(n_s.items())
            ],
            "stores": [build_store_row(s, n_w, classify_a) for s in ALL_STORES],
        },
        "moh": {
            "label": "MOH (Months of Holding)",
            "thresholds": {"green": "var -10% to +5%", "yellow": "var -15% to -10% / +5% to +10%", "red": "var <= -15% or >= +10%"},
            "miniso": {
                "metrics": m_mi,
                "target": moh_tgt_avg,
                "variance": round((m_mi.get("d1", 0) - moh_tgt_avg) / moh_tgt_avg, 4) if moh_tgt_avg else None,
                "status": classify_moh(m_mi.get("d1"), moh_tgt_avg),
            },
            "spocs": [
                {
                    "spoc_name": sp,
                    "store_count": len([s for s, v in SPOC_MAP.items() if v == sp]),
                    "metrics": rec,
                    "target": round(sum(MOH_TARGETS.get(s, 0) for s in SPOC_MAP if SPOC_MAP.get(s) == sp and MOH_TARGETS.get(s, 0) > 0) /
                                    max(1, len([s for s, v in SPOC_MAP.items() if v == sp and MOH_TARGETS.get(s, 0) > 0])), 2),
                    "status": classify_moh(
                        rec.get("d1"),
                        sum(MOH_TARGETS.get(s, 0) for s in SPOC_MAP if SPOC_MAP.get(s) == sp and MOH_TARGETS.get(s, 0) > 0) /
                        max(1, len([s for s, v in SPOC_MAP.items() if v == sp and MOH_TARGETS.get(s, 0) > 0]))
                    ),
                }
                for sp, rec in sorted(m_s.items())
            ],
            "stores": [build_store_row(s, m_w, classify_moh, is_moh=True) for s in ALL_STORES],
        },
    },
}


os.makedirs(OUTPUT_DIR, exist_ok=True)
with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
    json.dump(output, f, indent=2, default=str)

print(f"\nSaved: {OUTPUT_PATH}")
print(f"A:{len(a_w)} B:{len(b_w)} N:{len(n_w)} MOH:{len(m_w)} stores")
print("DONE")
