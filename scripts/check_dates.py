
import sqlite3, pandas as pd

conn = sqlite3.connect(r"C:\kpi_dashboard_local\kpi_history.db")

for store in ["INDR", "INEL"]:
    df = pd.read_sql("""
        SELECT snap_date, product_code, available_stock, total_stock
        FROM stock_snapshot
        WHERE store_code = ? AND snap_date IN ('2026-09-10','2026-09-09')
        ORDER BY product_code, snap_date
    """, conn, params=(store,))
    pivot = df.pivot(index="product_code", columns="snap_date", values="available_stock")
    identical = pivot.iloc[:,0].equals(pivot.iloc[:,1]) if pivot.shape[1] == 2 else "only one date has data"
    print(f"{store}: identical across the two dates? {identical}")