import pandas as pd
import numpy as np


def normalize_item_code(x) -> str:
    s = str(x).strip()
    if s == "" or s.lower() == "nan":
        return ""
    if "." in s:
        left, right = s.split(".", 1)
        if right.strip("0") == "":
            s = left
    return s.strip()


def _z_from_service_level(service_level: float) -> float:
    if service_level is None:
        return 0.0
    sl = float(service_level)
    if sl >= 0.999:
        return 3.09
    elif sl >= 0.99:
        return 2.33
    elif sl >= 0.98:
        return 2.05
    elif sl >= 0.95:
        return 1.64
    elif sl >= 0.90:
        return 1.28
    elif sl >= 0.85:
        return 1.04
    else:
        return 0.0


def plan_inventory_dc(
    df_product_forecast: pd.DataFrame,
    df_inventory_dc: pd.DataFrame,
    days_per_month: int = 30,
    service_level: float = 0.95,
    safety_stock_mode: str = "stat",      # "stat" أو "days"
    review_days: int = 7,                 # Review buffer فوق ROP
) -> pd.DataFrame:

    df_f = df_product_forecast.copy()
    df_i = df_inventory_dc.copy()

    df_f.columns = [str(c).strip() for c in df_f.columns]
    df_i.columns = [str(c).strip() for c in df_i.columns]

    # دعم item name بدل iteam name (ملفات المخزون)
    if "iteam name" not in df_i.columns and "item name" in df_i.columns:
        df_i = df_i.rename(columns={"item name": "iteam name"})

    required_forecast_cols = [
        "item code",
        "months_count",
        "months_list",
        "demand_total",
        "demand_avg_all",
        "demand_avg_nonzero",
        "demand_std",
        "demand_cv",
        "last_month_demand",
        "forecast_next_month",
        "forecast_next_month_rounded",
        "forecast_method",
    ]

    required_inv_cols = [
        "item code",
        "iteam name",
        "on_hand_dc",
        "on_order_from_china",
        "lead_time_import_days",
        "safety_stock_days",
    ]

    missing_f = [c for c in required_forecast_cols if c not in df_f.columns]
    missing_i = [c for c in required_inv_cols if c not in df_i.columns]

    if missing_f:
        raise ValueError(f"Missing columns in forecast df: {missing_f}")
    if missing_i:
        raise ValueError(f"Missing columns in inventory df: {missing_i}")

    # توحيد كود الصنف
    df_f["item_code_key"] = df_f["item code"].apply(normalize_item_code)
    df_i["item_code_key"] = df_i["item code"].apply(normalize_item_code)

    # منع تكرار item_code_key في forecast (احتياط إضافي)
    df_f["demand_total"] = pd.to_numeric(df_f["demand_total"], errors="coerce").fillna(0.0)
    df_f = df_f.sort_values("demand_total", ascending=False).drop_duplicates("item_code_key", keep="first")

    forecast_cols_for_merge = ["item_code_key"] + [c for c in required_forecast_cols if c != "item code"]
    if "abc_class" in df_f.columns:
        forecast_cols_for_merge.append("abc_class")

    df_merged = pd.merge(
        df_i,
        df_f[forecast_cols_for_merge],
        on="item_code_key",
        how="left",
    )

    # forecast الشهري -> يومي
    df_merged["forecast_next_month"] = pd.to_numeric(df_merged["forecast_next_month"], errors="coerce").fillna(0.0)
    df_merged["forecast_daily"] = df_merged["forecast_next_month"] / float(days_per_month)

    # مخزون متاح = on_hand + on_order (Inventory Position مبسط)
    df_merged["on_hand_dc"] = pd.to_numeric(df_merged["on_hand_dc"], errors="coerce").fillna(0.0)
    df_merged["on_order_from_china"] = pd.to_numeric(df_merged["on_order_from_china"], errors="coerce").fillna(0.0)
    df_merged["available_qty_total"] = df_merged["on_hand_dc"] + df_merged["on_order_from_china"]

    # LT & SS days
    df_merged["lead_time_import_days"] = pd.to_numeric(df_merged["lead_time_import_days"], errors="coerce").fillna(0.0)
    df_merged["safety_stock_days"] = pd.to_numeric(df_merged["safety_stock_days"], errors="coerce").fillna(0.0)

    # std -> daily std
    df_merged["demand_std"] = pd.to_numeric(df_merged["demand_std"], errors="coerce").fillna(0.0)
    df_merged["demand_daily_std"] = df_merged["demand_std"] / np.sqrt(float(days_per_month))

    # Service level حسب ABC
    def _map_service_level(row):
        abc = str(row.get("abc_class", "")).strip().upper()
        base = float(service_level)
        if abc == "A":
            return max(base, 0.98)
        elif abc == "B":
            return base
        elif abc == "C":
            return min(base, 0.90)
        else:
            return base

    if "abc_class" in df_merged.columns:
        df_merged["service_level_effective"] = df_merged.apply(_map_service_level, axis=1)
    else:
        df_merged["service_level_effective"] = float(service_level)

    # Safety stock
    def _calc_safety_stock(row):
        fd = row["forecast_daily"]
        ss_days = row["safety_stock_days"]
        lt_days = row["lead_time_import_days"]
        sigma_d = row["demand_daily_std"]
        sl = row["service_level_effective"]
        z = _z_from_service_level(sl)

        if fd <= 0:
            return 0.0

        if safety_stock_mode == "days":
            return fd * ss_days

        if z > 0 and sigma_d > 0 and lt_days > 0:
            return z * sigma_d * np.sqrt(lt_days)

        return fd * ss_days

    df_merged["safety_stock_qty"] = df_merged.apply(_calc_safety_stock, axis=1)

    # Days of supply
    def _calc_dos(row):
        fd = row["forecast_daily"]
        if fd <= 0:
            return np.nan
        return row["available_qty_total"] / fd

    df_merged["days_of_supply"] = df_merged.apply(_calc_dos, axis=1)

    # Reorder point
    df_merged["reorder_point"] = (
        df_merged["forecast_daily"] * df_merged["lead_time_import_days"]
        + df_merged["safety_stock_qty"]
    )

    # Min order qty to cover lead time (ضمن المنطق: لا تقل عن LT)
    df_merged["demand_during_lead_time"] = df_merged["forecast_daily"] * df_merged["lead_time_import_days"]
    df_merged["min_order_qty_lt"] = (df_merged["demand_during_lead_time"] - df_merged["available_qty_total"]).clip(lower=0.0)

    # Review buffer quantity (ROP + review_days)
    df_merged["review_buffer_qty"] = df_merged["forecast_daily"] * float(review_days)

    # =========================
    # قرار المخزون (تصحيح مهم):
    # - إذا available <= ROP  => Order
    # - إذا ROP < available <= ROP + buffer => Review
    # - غير ذلك => Hold
    # =========================
    def _decide_action(row):
        fd = row["forecast_daily"]
        avail = row["available_qty_total"]
        rop = row["reorder_point"]
        buffer_qty = row["review_buffer_qty"]

        if fd <= 0:
            return "Hold"

        if avail <= rop:
            return "Order"

        if avail <= rop + buffer_qty:
            return "Review"

        return "Hold"

    df_merged["inventory_decision"] = df_merged.apply(_decide_action, axis=1)

    # Target stock level (LT + SS days)
    df_merged["target_coverage_days"] = df_merged["lead_time_import_days"] + df_merged["safety_stock_days"]
    df_merged["target_stock_level"] = df_merged["forecast_daily"] * df_merged["target_coverage_days"]

    def _calc_order_qty(row):
        if row["inventory_decision"] != "Order":
            return 0
        fd = row["forecast_daily"]
        if fd <= 0:
            return 0

        avail = row["available_qty_total"]
        target = row["target_stock_level"]
        min_lt = row["min_order_qty_lt"]

        raw_needed = max(target - avail, 0.0)
        order_qty = max(raw_needed, min_lt)

        return int(round(order_qty)) if order_qty > 0 else 0

    df_merged["suggested_order_qty"] = df_merged.apply(_calc_order_qty, axis=1)

    # معلومات إضافية
    df_merged["service_level"] = float(service_level)
    df_merged["safety_stock_mode"] = safety_stock_mode
    df_merged["review_days"] = int(review_days)

    cols_order = [
        "item code",
        "iteam name",
        "months_count",
        "months_list",
        "demand_total",
        "demand_avg_all",
        "demand_avg_nonzero",
        "demand_std",
        "demand_daily_std",
        "demand_cv",
        "last_month_demand",
        "forecast_method",
        "forecast_next_month",
        "forecast_next_month_rounded",
        "forecast_daily",
        "on_hand_dc",
        "on_order_from_china",
        "available_qty_total",
        "lead_time_import_days",
        "demand_during_lead_time",
        "min_order_qty_lt",
        "safety_stock_days",
        "safety_stock_mode",
        "service_level",
        "service_level_effective",
        "safety_stock_qty",
        "days_of_supply",
        "reorder_point",
        "review_days",
        "review_buffer_qty",
        "target_coverage_days",
        "target_stock_level",
        "inventory_decision",
        "suggested_order_qty",
    ]

    if "abc_class" in df_merged.columns:
        cols_order.append("abc_class")

    cols_final = [c for c in cols_order if c in df_merged.columns]
    return df_merged[cols_final].copy()
