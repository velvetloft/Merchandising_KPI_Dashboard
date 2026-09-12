import pandas as pd
import sqlite3
import glob
import os
from datetime import datetime

DB_PATH = r"C:\kpi_dashboard_local\kpi_history.db"


try:
    _conn = sqlite3.connect(DB_PATH)
    _conn.execute("CREATE INDEX IF NOT EXISTS idx_stock_date ON stock_snapshot(snap_date)")
    _conn.execute("CREATE INDEX IF NOT EXISTS idx_sales_date ON sales_daily(snap_date)")
    _conn.commit()
    _conn.close()
except Exception:
    pass

# Anchored to the script's own folder, not the working directory it happens
# to be launched from - a relative path here silently drifts between runs
# (same class of bug that broke calc_kpis.py's output path).
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROCESSED_LOG = os.path.join(SCRIPT_DIR, "..", "processed_files.txt")
PROCESSED_LOG = os.path.normpath(PROCESSED_LOG)

# Every run appends one line here - check this file first if the dashboard
# ever looks stale again. A run that found 0 new files in every folder is
# a sign the source drive/shortcut wasn't mounted when this ran, not that
# there's nothing new to load.
RUN_LOG = os.path.join(SCRIPT_DIR, "..", "ingest_run_log.txt")
RUN_LOG = os.path.normpath(RUN_LOG)
_run_summary = {"started_at": datetime.now().isoformat(timespec="seconds"), "folders": {}}

def check_folder(label, folder):
    """Warn loudly (not just silently return []) when a source folder is
    missing or unreachable - this is the most likely cause of a dashboard
    that quietly stops updating: the J:\\ Drive shortcut not being mounted
    when the scheduled task runs."""
    exists = os.path.isdir(folder)
    if not exists:
        print(f"  *** WARNING: {label} folder is not reachable: {folder}")
        print(f"  *** This usually means the network/Drive shortcut isn't mounted. "
              f"No files will be loaded from this source this run.")
    _run_summary["folders"][label] = {"path": folder, "reachable": exists, "files_found": 0, "files_loaded": 0}
    return exists

def load_processed_log():
    if not os.path.exists(PROCESSED_LOG):
        return set()
    with open(PROCESSED_LOG, "r", encoding="utf-8") as f:
        return set(line.strip() for line in f if line.strip())

def mark_processed(filepath):
    with open(PROCESSED_LOG, "a", encoding="utf-8") as f:
        f.write(filepath + "\n")

PROCESSED = load_processed_log()

def upsert_by_date(df, table, date_col, conn):
    """Delete rows for the dates present in df, then insert df fresh.
    Keyed on actual date values inside the data, never on filename -
    safe against duplicate uploads under different filenames, and safe
    to re-run any number of times."""
    if df.empty:
        return
    dates_in_file = df[date_col].dropna().unique().tolist()
    existing_tables = pd.read_sql("SELECT name FROM sqlite_master WHERE type='table'", conn)["name"].tolist()
    if table in existing_tables and dates_in_file:
        placeholders = ",".join(["?"] * len(dates_in_file))
        conn.execute(f"DELETE FROM {table} WHERE {date_col} IN ({placeholders})", dates_in_file)
        conn.commit()
        df.to_sql(table, conn, if_exists="append", index=False)
    else:
        df.to_sql(table, conn, if_exists="replace", index=False)
    print(f"  {table}: upserted {len(df)} rows for dates {sorted(set(dates_in_file))}")


def list_source_files(folder, label):
    if not check_folder(label, folder):
        return []
    files = [f for f in glob.glob(os.path.join(folder, "*.xlsx")) if not os.path.basename(f).startswith("~$")]
    _run_summary["folders"][label]["files_found"] = len(files)
    if not files:
        print(f"  No .xlsx files found in {label} folder (folder is reachable, just empty or all already processed).")
    return files


# ============================================================
# STOCK REPORT - scan whole folder
# ============================================================
STOCK_FOLDER = r"J:\.shortcut-targets-by-id\1-OzYqibxMbDT38omFZxDLknsTKiBDbpt\Business Analytics\Dashboard\Supporting_Folder_For_Complete_dashboard\Last 90 days stock"

STOCK_COLMAP = {
    "日期": "snap_date",
    "Store Name门店代码": "store_code",
    "Store Name门店名称": "store_name",
    "store_status_name门店状态": "store_status",
    "Product Code商品代码": "product_code",
    "Category Name大类名称": "category",
    "Sub-Category Name中类名称": "sub_category",
    "Class Name小类名称": "class_name",
    "Sub-Class Name细类名称": "sub_class_name",
    "Price(零售价)": "retail_price",
    "Daily Sales日均销售": "daily_sales_avg",
    "Total Stock总库存": "total_stock",
    "Available Stock 门店可用库存": "available_stock",
    "Stock AMT(Local)库存金额": "stock_value",
    "QTY-Negative Stock负库存数量": "negative_stock_qty",
}

print("=" * 60)
print("STOCK REPORTS")
print("=" * 60)
conn = sqlite3.connect(DB_PATH)
for f in list_source_files(STOCK_FOLDER, "stock"):
    if f in PROCESSED:
        print(f"Skipping (already processed): {os.path.basename(f)}")
        continue
    print(f"Processing: {os.path.basename(f)}")
    try:
        df = pd.read_excel(f)
        df = df.rename(columns=STOCK_COLMAP)
        missing = [c for c in STOCK_COLMAP.values() if c not in df.columns]
        if missing:
            print(f"  SKIPPED - missing columns: {missing}")
            continue
        df = df[list(STOCK_COLMAP.values())]
        before = len(df)
        df = df[~df["category"].isin(["Decoration Materials", "Consumables"])]
        df["snap_date"] = pd.to_datetime(df["snap_date"]).dt.strftime("%Y-%m-%d")
        print(f"  Excluded {before - len(df)} rows (Decoration/Consumables). {len(df)} rows remain.")
        upsert_by_date(df, "stock_snapshot", "snap_date", conn)
        mark_processed(f)
    except Exception as e:
        print(f"  ERROR processing {f}: {e}")
conn.close()


# ============================================================
# SALES REPORT - scan whole folder
# ============================================================
SALES_FOLDER = r"J:\.shortcut-targets-by-id\1-OzYqibxMbDT38omFZxDLknsTKiBDbpt\Business Analytics\Dashboard\Supporting_Folder_For_Complete_dashboard\last_3_month_sale"

SALES_COLMAP = {
    "日期": "snap_date",
    "Store Name门店代码": "store_code",
    "Store Name门店名称": "store_name",
    "Product Code商品代码": "product_code",
    "Category Name大类名称": "category",
    "Sub-Category Name中类名称": "sub_category",
    "QTY销售数量": "sales_qty",
    "Stock库存数量": "stock_qty",
    "Sales销售金额": "sales_value",
    "Net Sales  QTY净销售数量": "net_sales_qty",
    "Net Sales净销售金额": "net_sales_value",
}

print("\n" + "=" * 60)
print("SALES REPORTS")
print("=" * 60)
conn = sqlite3.connect(DB_PATH)
for f in list_source_files(SALES_FOLDER, "sales"):
    if f in PROCESSED:
        print(f"Skipping (already processed): {os.path.basename(f)}")
        continue
    print(f"Processing: {os.path.basename(f)}")
    try:
        df = pd.read_excel(f)
        df = df.rename(columns=SALES_COLMAP)
        missing = [c for c in SALES_COLMAP.values() if c not in df.columns]
        if missing:
            print(f"  SKIPPED - missing columns: {missing}")
            continue
        df = df[list(SALES_COLMAP.values())]
        before = len(df)
        df = df[~df["category"].isin(["Decoration Materials", "Consumables"])]
        df["snap_date"] = pd.to_datetime(df["snap_date"]).dt.strftime("%Y-%m-%d")
        print(f"  Excluded {before - len(df)} rows (Decoration/Consumables). {len(df)} rows remain.")
        upsert_by_date(df, "sales_daily", "snap_date", conn)
        mark_processed(f)
    except Exception as e:
        print(f"  ERROR processing {f}: {e}")
conn.close()


# ============================================================
# B2B REPORT - scan whole folder
# ============================================================
B2B_FOLDER = r"J:\.shortcut-targets-by-id\1-OzYqibxMbDT38omFZxDLknsTKiBDbpt\Business Analytics\Dashboard\Supporting_Folder_For_Complete_dashboard\b2b_purchase"

B2B_COLMAP = {
    "document_date": "doc_date",
    "sold-to party": "store_code",
    "product_description_item": "product_code",
    "English Product Name": "product_name",
    "quantity_item": "qty",
    "taxable_amount_item": "taxable_amount",
}

print("\n" + "=" * 60)
print("B2B REPORTS")
print("=" * 60)
conn = sqlite3.connect(DB_PATH)
for f in list_source_files(B2B_FOLDER, "b2b"):
    if f in PROCESSED:
        print(f"Skipping (already processed): {os.path.basename(f)}")
        continue
    print(f"Processing: {os.path.basename(f)}")
    try:
        df = pd.read_excel(f)
        df = df.rename(columns=B2B_COLMAP)
        missing = [c for c in B2B_COLMAP.values() if c not in df.columns]
        if missing:
            print(f"  SKIPPED - missing columns: {missing}")
            continue
        df = df[list(B2B_COLMAP.values())]
        df["doc_date"] = pd.to_datetime(df["doc_date"], format="%Y%m%d", errors="coerce").dt.strftime("%Y-%m-%d")
        upsert_by_date(df, "b2b_daily", "doc_date", conn)
        mark_processed(f)
    except Exception as e:
        print(f"  ERROR processing {f}: {e}")
conn.close()


# ============================================================
# B2C REPORT - scan whole folder
# ============================================================
B2C_FOLDER = r"J:\.shortcut-targets-by-id\1-OzYqibxMbDT38omFZxDLknsTKiBDbpt\Business Analytics\Dashboard\Supporting_Folder_For_Complete_dashboard\b2c_sale"

B2C_COLMAP = {
    "门店代码(Store Code)": "store_code",
    "门店名称(Store Name)": "store_name",
    "开票日期(Date)": "invoice_date",
    "商品代码(Product Code)": "product_code",
    "商品名称(英文)(English Product Name)": "product_name",
    "销售数量(Qty)": "qty",
    "实收金额(Amt with GST)": "amt_with_gst",
    "Basic Price": "basic_price",
}

print("\n" + "=" * 60)
print("B2C REPORTS")
print("=" * 60)
conn = sqlite3.connect(DB_PATH)
for f in list_source_files(B2C_FOLDER, "b2c"):
    if f in PROCESSED:
        print(f"Skipping (already processed): {os.path.basename(f)}")
        continue
    print(f"Processing: {os.path.basename(f)}")
    try:
        df = pd.read_excel(f)
        df = df.rename(columns=B2C_COLMAP)
        missing = [c for c in B2C_COLMAP.values() if c not in df.columns]
        if missing:
            print(f"  SKIPPED - missing columns: {missing}")
            continue
        df = df[list(B2C_COLMAP.values())]
        df["invoice_date"] = pd.to_datetime(df["invoice_date"]).dt.strftime("%Y-%m-%d")
        upsert_by_date(df, "b2c_daily", "invoice_date", conn)
        mark_processed(f)
    except Exception as e:
        print(f"  ERROR processing {f}: {e}")
conn.close()


# ============================================================
# ORDERING MANIFEST - scan whole folder
# This is a BRAND-WIDE catalog (not per store) - a SKU only has a row here
# on a given date if it was actually orderable from the brand that day (you
# download/keep this file per your own confirmation, so presence = in MOS,
# absence = not in MOS - no status column to interpret).
#
# calc_kpis.py's Availability KPI uses this as the "not available in MOS
# for 5 consecutive days" signal: a SKU missing from the manifest for 5+
# days is dropped from every store's denominator, while a store's actual
# on-shelf stock is still checked separately, per store.
#
# Only date + skuCode are needed. barcode, skuName, categories, price,
# stockQty, activated, etc. from the source file are dropped.
# ============================================================
ORDERING_MANIFEST_FOLDER = r"J:\.shortcut-targets-by-id\1-OzYqibxMbDT38omFZxDLknsTKiBDbpt\Business Analytics\Dashboard\Supporting_Folder_For_Complete_dashboard\ordering_manifest"

ORDERING_MANIFEST_COLMAP = {
    "date": "manifest_date",
    "skuCode": "product_code",
}

print("\n" + "=" * 60)
print("ORDERING MANIFEST (MOS status)")
print("=" * 60)
conn = sqlite3.connect(DB_PATH)
for f in list_source_files(ORDERING_MANIFEST_FOLDER, "ordering_manifest"):
    if f in PROCESSED:
        print(f"Skipping (already processed): {os.path.basename(f)}")
        continue
    print(f"Processing: {os.path.basename(f)}")
    try:
        df = pd.read_excel(f)
        df = df.rename(columns=ORDERING_MANIFEST_COLMAP)
        missing = [c for c in ORDERING_MANIFEST_COLMAP.values() if c not in df.columns]
        if missing:
            print(f"  SKIPPED - missing columns: {missing}")
            continue
        df = df[list(ORDERING_MANIFEST_COLMAP.values())]
        df = df.drop_duplicates(subset=["manifest_date", "product_code"])
        df["manifest_date"] = pd.to_datetime(df["manifest_date"], dayfirst=True).dt.strftime("%Y-%m-%d")
        upsert_by_date(df, "ordering_manifest", "manifest_date", conn)
        mark_processed(f)
    except Exception as e:
        print(f"  ERROR processing {f}: {e}")
conn.close()


# ============================================================
# PRODUCT MASTER - always truncate + reload fresh (current snapshot,
# not a date-based history log). Warns if more than one file is
# sitting in the folder, since that creates ambiguity.
# ============================================================
PRODUCT_MASTER_FOLDER = r"J:\.shortcut-targets-by-id\1-OzYqibxMbDT38omFZxDLknsTKiBDbpt\Business Analytics\Dashboard\Supporting_Folder_For_Complete_dashboard\product_master"

# REQUIRED for the New-SKU Availability / New-SKU Sales Contribution KPIs -
# calc_kpis.py needs a launch-date column to know which SKUs are "new"
# (launched in the last 30 days). If your product master export's launch
# date column has a different header, change the key on the left below to
# match it exactly - the value on the right (b2b_launch_date) must stay as
# is since calc_kpis.py reads that exact column name.
PRODUCT_MASTER_COLMAP = {
    "product_code": "product_code",
    "product_name": "product_name",
    "category": "category",
    "sub_category": "sub_category",
    "class": "class_name",
    "sub_class": "sub_class_name",
    "avg_price": "avg_price",
    "avg_price_sale": "avg_price_sale",
    "b2b_avg_price": "b2b_avg_price",
    "b2b_launch_date": "b2b_launch_date",
}
# Columns that are nice-to-have but shouldn't block loading the rest of the
# product master if they're missing from the source file - New-SKU KPIs
# will just come back empty (and calc_kpis.py will say so clearly) instead
# of the whole ingest step failing.
PRODUCT_MASTER_OPTIONAL = {"b2b_launch_date"}

print("\n" + "=" * 60)
print("PRODUCT MASTER")
print("=" * 60)
pm_files = list_source_files(PRODUCT_MASTER_FOLDER, "product_master")

if len(pm_files) == 0:
    print("No product_master files found - table left unchanged.")
elif len(pm_files) > 1:
    print(f"WARNING: {len(pm_files)} files found in product_master folder - expected only 1.")
    print("Files found:")
    for f in pm_files:
        print(f"  - {os.path.basename(f)}")
    print("Please keep only the current file in this folder to avoid ambiguity.")
    print("Using the most recently modified one for now:")
    chosen_file = max(pm_files, key=os.path.getmtime)
    print(f"  -> {os.path.basename(chosen_file)}")
    pm_files_to_use = [chosen_file]
else:
    pm_files_to_use = pm_files

if pm_files:
    latest_pm_file = pm_files_to_use[0] if len(pm_files) > 1 else pm_files[0]
    print(f"Loading: {os.path.basename(latest_pm_file)}")
    pm_df = pd.read_excel(latest_pm_file)
    pm_df = pm_df.rename(columns=PRODUCT_MASTER_COLMAP)
    missing = [c for c in PRODUCT_MASTER_COLMAP.values() if c not in pm_df.columns]
    required_missing = [c for c in missing if c not in PRODUCT_MASTER_OPTIONAL]
    optional_missing = [c for c in missing if c in PRODUCT_MASTER_OPTIONAL]
    if required_missing:
        print(f"  SKIPPED - missing required columns: {required_missing}")
    else:
        for c in optional_missing:
            pm_df[c] = None
        if optional_missing:
            print(f"  NOTE: source file has no '{', '.join(optional_missing)}' column(s) - "
                  f"New-SKU KPIs (New-SKU Availability, New-SKU Sales Contribution) will come "
                  f"back empty until this is added to the product master export.")
        pm_df = pm_df[list(PRODUCT_MASTER_COLMAP.values())]
        pm_df = pm_df.drop_duplicates(subset="product_code", keep="last")
        conn = sqlite3.connect(DB_PATH)
        pm_df.to_sql("product_master", conn, if_exists="replace", index=False)
        conn.close()
        print("Saved to kpi_history.db, table: product_master (fully replaced)")
        print("Rows:", len(pm_df))
        print("Category filled?", pm_df["category"].notna().sum(), "out of", len(pm_df))
        if "b2b_launch_date" not in optional_missing:
            print("Launch date filled?", pm_df["b2b_launch_date"].notna().sum(), "out of", len(pm_df))


# ============================================================
# SPOC NAME - lookup table (store → SPOC person)
# Always truncate + reload fresh (current snapshot, like product_master).
# calc_kpis.py reads from store_spoc_mapping to display SPOC names on
# A-Class, B-Class, and MOH dashboard rows. If this table is missing,
# spoc_name silently shows as null on every row.
# ============================================================
SPOC_FOLDER = r"J:\.shortcut-targets-by-id\1-OzYqibxMbDT38omFZxDLknsTKiBDbpt\Business Analytics\Dashboard\Supporting_Folder_For_Complete_dashboard\Merchandising_KPI_Dashboard"

SPOC_COLMAP = {
    "Store Code": "store_code",
    "Store Name": "store_name",
    "SPOC Name": "spoc_name",
}

print("\n" + "=" * 60)
print("SPOC NAME (store → SPOC mapping)")
print("=" * 60)
spoc_files = list_source_files(SPOC_FOLDER, "spoc_name")

# Filter to only spoc_name.xlsx (folder may contain other files like Moh_targets.xlsx)
spoc_files = [f for f in spoc_files if os.path.basename(f).lower() == "spoc_name.xlsx"]

if not spoc_files:
    print("No spoc_name.xlsx found in folder - table left unchanged.")
else:
    spoc_file = spoc_files[0]
    print(f"Loading: {os.path.basename(spoc_file)}")
    try:
        spoc_df = pd.read_excel(spoc_file)
        spoc_df = spoc_df.rename(columns=SPOC_COLMAP)
        missing = [c for c in SPOC_COLMAP.values() if c not in spoc_df.columns]
        if missing:
            print(f"  SKIPPED - missing columns: {missing}")
        else:
            spoc_df = spoc_df[list(SPOC_COLMAP.values())]
            spoc_df = spoc_df.drop_duplicates(subset="store_code", keep="last")
            spoc_df["store_code"] = spoc_df["store_code"].astype(str).str.strip()
            spoc_df["spoc_name"] = spoc_df["spoc_name"].fillna("").astype(str).str.strip()
            conn = sqlite3.connect(DB_PATH)
            spoc_df.to_sql("store_spoc_mapping", conn, if_exists="replace", index=False)
            conn.close()
            mark_processed(spoc_file)
            print(f"Saved to kpi_history.db, table: store_spoc_mapping (fully replaced)")
            print(f"Rows: {len(spoc_df)}")
            print(f"SPOC name filled? {spoc_df['spoc_name'].replace('', pd.NA).notna().sum()} out of {len(spoc_df)}")
    except Exception as e:
        print(f"  ERROR processing {os.path.basename(spoc_file)}: {e}")


# ============================================================
# MOH TARGETS - lookup table (store → monthly MOH target)
# Always truncate + reload fresh (current snapshot, like product_master).
# The source file has a month-specific target column (e.g. "Sep MOH Tgt").
# This ingest detects whichever month column is present and stores it as
# a generic "moh_target" column so calc_kpis.py can read it regardless
# of which month's file is uploaded.
#
# calc_kpis.py currently computes MOH targets dynamically (target_moh()
# function). Once this table is ingested, that function can be replaced
# with a direct lookup against moh_targets.moh_target.
# ============================================================
MOH_TARGET_FOLDER = r"J:\.shortcut-targets-by-id\1-OzYqibxMbDT38omFZxDLknsTKiBDbpt\Business Analytics\Dashboard\Supporting_Folder_For_Complete_dashboard\Merchandising_KPI_Dashboard"

MOH_TARGET_ID_COLS = {
    "New store Code": "store_code",
    "Store Name": "store_name",
}

print("\n" + "=" * 60)
print("MOH TARGETS (store → monthly MOH target)")
print("=" * 60)
moh_files = list_source_files(MOH_TARGET_FOLDER, "moh_targets")

# Filter to only Moh_targets.xlsx (folder may contain other files like spoc_name.xlsx)
moh_files = [f for f in moh_files if os.path.basename(f).lower() == "moh_targets.xlsx"]

if not moh_files:
    print("No Moh_targets.xlsx found in folder - table left unchanged.")
else:
    moh_file = moh_files[0]
    print(f"Loading: {os.path.basename(moh_file)}")
    try:
        moh_df = pd.read_excel(moh_file)
        moh_df = moh_df.rename(columns=MOH_TARGET_ID_COLS)

        # Detect the month-specific target column dynamically.
        # Expected pattern: a column containing both "moh" and "tgt" (case-insensitive).
        target_col = None
        for col in moh_df.columns:
            col_lower = col.lower()
            if "moh" in col_lower and "tgt" in col_lower:
                target_col = col
                break

        id_missing = [c for c in MOH_TARGET_ID_COLS.values() if c not in moh_df.columns]
        if id_missing:
            print(f"  SKIPPED - missing ID columns: {id_missing}")
        elif target_col is None:
            print(f"  SKIPPED - no column containing 'moh' and 'tgt' found. "
                  f"Columns present: {moh_df.columns.tolist()}")
        else:
            print(f"  Detected target column: '{target_col}' → moh_target")
            moh_df = moh_df.rename(columns={target_col: "moh_target"})
            keep_cols = list(MOH_TARGET_ID_COLS.values()) + ["moh_target"]
            moh_df = moh_df[keep_cols]
            moh_df = moh_df.drop_duplicates(subset="store_code", keep="last")
            moh_df["store_code"] = moh_df["store_code"].astype(str).str.strip()
            moh_df["moh_target"] = pd.to_numeric(moh_df["moh_target"], errors="coerce")
            conn = sqlite3.connect(DB_PATH)
            moh_df.to_sql("moh_targets", conn, if_exists="replace", index=False)
            conn.close()
            mark_processed(moh_file)
            print(f"Saved to kpi_history.db, table: moh_targets (fully replaced)")
            print(f"Rows: {len(moh_df)}")
            print(f"MOH target filled? {moh_df['moh_target'].notna().sum()} out of {len(moh_df)}")
    except Exception as e:
        print(f"  ERROR processing {os.path.basename(moh_file)}: {e}")


# ============================================================
# RUN SUMMARY — quick, unmissable proof of what actually happened this run.
# ============================================================
print("\n" + "=" * 60)
print("RUN SUMMARY")
print("=" * 60)
for label, info in _run_summary["folders"].items():
    status = "OK" if info["reachable"] else "UNREACHABLE"
    print(f"  {label:16s} [{status:11s}] files found: {info['files_found']}")
with open(RUN_LOG, "a", encoding="utf-8") as f:
    f.write(f"{datetime.now().isoformat(timespec='seconds')} | " +
            " | ".join(f"{k}: {'OK' if v['reachable'] else 'UNREACHABLE'} ({v['files_found']} files)"
                        for k, v in _run_summary["folders"].items()) + "\n")

print("\n" + "=" * 60)
print("INGEST COMPLETE")
print("=" * 60)
print(f"Full run log: {RUN_LOG}")
