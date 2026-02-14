# core_demand.py
import re
import pandas as pd
import numpy as np
from typing import Dict
from decimal import Decimal, InvalidOperation


# =====================================================
# Helpers
# =====================================================
def normalize_item_code(x) -> str:
    """
    يوحّد كود الصنف بأمان (يعالج scientific notation مثل 6.97E+12)
    ويحافظ على الأكواد الصغيرة/الكسور لو ظهرت (مثل 0.5).
    """
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return ""

    s = str(x).strip()
    if s == "" or s.lower() == "nan":
        return ""

    s_clean = s.replace(",", "").strip()

    # حاول تحوله لرقم باستخدام Decimal (يدعم scientific notation)
    try:
        d = Decimal(s_clean)
        # لو رقم صحيح
        if d == d.to_integral_value():
            return str(d.to_integral_value())
        # لو فيه كسور (نحافظ عليه كما هو)
        return s_clean
    except (InvalidOperation, ValueError):
        # إذا مش رقم، رجّع النص (بس غالباً رح ينحذف لاحقاً)
        return s_clean


def standardize_month_label(label: str) -> str:
    """
    يقبل:
      - 2025-01
      - 2025/01
      - 01-2025
      - 01/2025
    ويرجع دائمًا: YYYY-MM
    """
    s = str(label).strip()

    m1 = re.match(r"^(\d{4})[-/](\d{1,2})$", s)
    if m1:
        y, m = int(m1.group(1)), int(m1.group(2))
        return f"{y:04d}-{m:02d}"

    m2 = re.match(r"^(\d{1,2})[-/](\d{4})$", s)
    if m2:
        m, y = int(m2.group(1)), int(m2.group(2))
        return f"{y:04d}-{m:02d}"

    return s


def most_frequent_name(s: pd.Series) -> str:
    s = s.dropna().astype(str).str.strip()
    s = s[s != ""]
    if len(s) == 0:
        return ""
    return s.value_counts().idxmax()


# =====================================================
# 1) تحويل شيت شهر واحد من Wide إلى Long
# =====================================================
def build_long_from_wide_month(
    df_wide: pd.DataFrame,
    month_label: str,
    item_col: str = "iteam name",
    code_col: str = "item code",
) -> pd.DataFrame:
    """
    df_wide: داتا الإكسيل الشهرية (كل الأعمدة: item, code, فروع, Totals)
    month_label: مثل "2025-01" أو "01-2025"
    """
    df = df_wide.copy()
    df.columns = [str(c).strip() for c in df.columns]

    month_label_std = standardize_month_label(month_label)

    # =====================================================
    # تصحيح تلقائي لو الأعمدة معكوسة (حسب داتا شركتك)
    # =====================================================
    if "iteam name" in df.columns and "item code" in df.columns:
        sample_name = df["iteam name"].dropna().astype(str).head(10)
        sample_code = df["item code"].dropna().astype(str).head(10)

        name_digits_ratio = (sample_name.str.strip().str.replace(".", "", regex=False).str.isdigit()).mean() if len(sample_name) else 0
        code_digits_ratio = (sample_code.str.strip().str.replace(".", "", regex=False).str.isdigit()).mean() if len(sample_code) else 0

        # لو iteam name أغلبه أرقام (باركود) و item code أغلبه نص (اسم)
        if name_digits_ratio > 0.8 and code_digits_ratio < 0.5:
            tmp = df["item code"].copy()
            df["item code"] = df["iteam name"]
            df["iteam name"] = tmp

    # =====================================================
    # أعمدة الفروع: استثناء أي عمود فيه total
    # =====================================================
    def _is_total_col(c: str) -> bool:
        return "total" in str(c).strip().lower()

    branch_cols = [
        c for c in df.columns
        if c not in [item_col, code_col]
        and (not _is_total_col(c))
    ]

    # melt -> Long
    long_df = df.melt(
        id_vars=[item_col, code_col],
        value_vars=branch_cols,
        var_name="branch",
        value_name="demand",
    )

    long_df["branch"] = long_df["branch"].astype(str).str.strip()

    # توحيد كود الصنف
    long_df[code_col] = long_df[code_col].apply(normalize_item_code)

    # حذف صفوف Totals/العناوين (أي كود صار فاضي أو غير صالح رقميًا بشكل كامل)
    # نعتبر الصالح: قادر يتحول لرقم عبر Decimal -> normalize_item_code رجع شيء غير فارغ
    long_df = long_df[long_df[code_col].astype(str).str.strip() != ""].copy()

    # تنظيف اسم المنتج + حذف صفوف إجمالية لو موجودة
    long_df[item_col] = long_df[item_col].astype(str).str.strip()
    mask_not_total_row = ~long_df[item_col].str.lower().str.contains("total", na=False)
    long_df = long_df[mask_not_total_row].copy()

    # الطلب رقم
    # ========= تعديل الأداء المهم =========
    # الفراغ = منتج متوقف => يبقى NaN => نحذفه مباشرة
    # الصفر الحقيقي يبقى 0
    long_df["demand"] = pd.to_numeric(long_df["demand"], errors="coerce")
    long_df = long_df[long_df["demand"].notna()].copy()

    long_df["month"] = month_label_std

    # إعادة تسمية نهائية ثابتة
    long_df = long_df.rename(columns={code_col: "item code", item_col: "iteam name"})
    long_df = long_df[["month", "branch", "item code", "iteam name", "demand"]]

    return long_df


# =====================================================
# 2) دمج كل الأشهر إلى Long واحد + إزالة التكرارات
# =====================================================
def build_demand_long(month_dfs: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    """
    output: month, branch, item code, iteam name, demand
    """
    all_months = []
    for month_label, df_wide in month_dfs.items():
        long_df = build_long_from_wide_month(df_wide, month_label)
        all_months.append(long_df)

    df_long = pd.concat(all_months, ignore_index=True)

    # تنظيف أساسي
    df_long["month"] = df_long["month"].astype(str).str.strip()
    df_long["branch"] = df_long["branch"].astype(str).str.strip()
    df_long["item code"] = df_long["item code"].apply(normalize_item_code)
    df_long["iteam name"] = df_long["iteam name"].astype(str).str.strip()

    # الطلب لازم يبقى رقمي (بدون fillna(0) عشان ما نرجع نكبر الداتا)
    df_long["demand"] = pd.to_numeric(df_long["demand"], errors="coerce")
    df_long = df_long[df_long["demand"].notna()].copy()

    # اسم ثابت لكل كود (أكثر اسم تكرارًا)
    name_map = (
        df_long.groupby("item code")["iteam name"]
        .apply(most_frequent_name)
        .rename("iteam name")
        .reset_index()
    )

    # ====== منع تكرار نفس الشهر داخل نفس الفرع لنفس الصنف ======
    df_long = (
        df_long.groupby(["month", "branch", "item code"], as_index=False)["demand"]
        .sum()
        .merge(name_map, on="item code", how="left")
    )

    df_long.sort_values(["item code", "branch", "month"], inplace=True)
    return df_long


# =====================================================
# 3) Stats + Forecast
# =====================================================
def _compute_demand_stats_and_forecast(demand: pd.Series) -> dict:
    demand = demand.astype(float).fillna(0.0)
    n = len(demand)

    demand_total = float(demand.sum())
    demand_avg_all = float(demand.mean()) if n > 0 else 0.0

    positive = demand[demand > 0]
    demand_avg_nonzero = float(positive.mean()) if len(positive) > 0 else 0.0

    demand_std = float(demand.std(ddof=0)) if n > 1 else 0.0
    demand_cv = float(demand_std / demand_avg_all) if demand_avg_all > 0 else np.nan

    last_month_demand = float(demand.iloc[-1]) if n > 0 else 0.0
    zeros_count = int((demand == 0).sum())
    zero_ratio = float(zeros_count / n) if n > 0 else 1.0

    # ========= Forecast =========
    if demand_total == 0:
        forecast_next_month = 0.0
        forecast_method = "zero_demand"
    else:
        if zero_ratio >= 0.5 and demand_avg_nonzero > 0:
            forecast_next_month = demand_avg_nonzero
            forecast_method = "avg_nonzero_intermittent"
        else:
            if n >= 4:
                recent = demand.tail(min(9, n)).to_numpy()
                m = len(recent)

                if m <= 3:
                    forecast_next_month = float(recent.mean())
                    forecast_method = f"mean_last_{m}"
                else:
                    num_last_3 = min(3, m)
                    num_old = m - num_last_3

                    weights_old = np.ones(num_old)
                    weights_last = np.full(num_last_3, 2.0)
                    weights = np.concatenate([weights_old, weights_last])

                    forecast_next_month = float((recent * weights).sum() / weights.sum())
                    forecast_method = f"weighted_last_{m}_old1_new2"
            else:
                forecast_next_month = float(demand.mean()) if n > 0 else 0.0
                forecast_method = "mean_all_few_months"

    forecast_next_month_rounded = max(int(round(forecast_next_month)), 0)

    return {
        "months_count": n,
        "demand_total": demand_total,
        "demand_avg_all": demand_avg_all,
        "demand_avg_nonzero": demand_avg_nonzero,
        "demand_std": demand_std,
        "demand_cv": demand_cv,
        "last_month_demand": last_month_demand,
        "zeros_count": zeros_count,
        "zero_ratio": zero_ratio,
        "forecast_next_month": float(forecast_next_month),
        "forecast_next_month_rounded": forecast_next_month_rounded,
        "forecast_method": forecast_method,
    }


# =====================================================
# 4) Branch Level Forecast
# =====================================================
def build_branch_level_forecast(df_long: pd.DataFrame) -> pd.DataFrame:
    df = df_long.copy()
    df["month"] = df["month"].astype(str)
    df["branch"] = df["branch"].astype(str)
    df["item code"] = df["item code"].apply(normalize_item_code)

    # ضمان: صف واحد لكل شهر داخل نفس الفرع
    df = df.groupby(["month", "branch", "item code"], as_index=False).agg(
        demand=("demand", "sum"),
        iteam_name=("iteam name", most_frequent_name),
    )

    def _summarize(g: pd.DataFrame) -> pd.Series:
        g = g.sort_values("month")
        months = g["month"].astype(str).tolist()
        demand = g["demand"]
        stats = _compute_demand_stats_and_forecast(demand)

        return pd.Series({
            "iteam name": most_frequent_name(g["iteam_name"]),
            "months_count": stats["months_count"],
            "months_list": ",".join(months),
            "demand_total": stats["demand_total"],
            "demand_avg_all": stats["demand_avg_all"],
            "demand_avg_nonzero": stats["demand_avg_nonzero"],
            "demand_std": stats["demand_std"],
            "demand_cv": stats["demand_cv"],
            "last_month_demand": stats["last_month_demand"],
            "forecast_next_month": stats["forecast_next_month"],
            "forecast_next_month_rounded": stats["forecast_next_month_rounded"],
            "forecast_method": stats["forecast_method"],
            "zeros_count": stats["zeros_count"],
            "zero_ratio": stats["zero_ratio"],
        })

    summary = df.groupby(["item code", "branch"]).apply(_summarize).reset_index()

    cols = [
        "item code", "iteam name", "branch",
        "months_count", "months_list",
        "demand_total", "demand_avg_all", "demand_avg_nonzero",
        "demand_std", "demand_cv",
        "last_month_demand",
        "forecast_next_month", "forecast_next_month_rounded",
        "forecast_method",
        "zeros_count", "zero_ratio",
    ]
    return summary[cols].copy()


# =====================================================
# 5) Company Level Forecast + ABC (صف واحد لكل item code)
# =====================================================
def build_product_level_forecast(df_long: pd.DataFrame) -> pd.DataFrame:
    df = df_long.copy()
    df["month"] = df["month"].astype(str)
    df["item code"] = df["item code"].apply(normalize_item_code)

    # اسم ثابت لكل كود
    name_map = (
        df.groupby("item code")["iteam name"]
        .apply(most_frequent_name)
        .rename("iteam name")
        .reset_index()
    )

    # طلب شهري على مستوى الشركة
    prod_month = (
        df.groupby(["item code", "month"], as_index=False)["demand"]
        .sum()
        .rename(columns={"demand": "demand_total_branches"})
        .merge(name_map, on="item code", how="left")
    )

    prod_month.sort_values(["item code", "month"], inplace=True)

    def _summarize(g: pd.DataFrame) -> pd.Series:
        g = g.sort_values("month")
        demand = g["demand_total_branches"]
        months = g["month"].astype(str).tolist()
        stats = _compute_demand_stats_and_forecast(demand)

        return pd.Series({
            "months_count": stats["months_count"],
            "months_list": ",".join(months),
            "demand_total": stats["demand_total"],
            "demand_avg_all": stats["demand_avg_all"],
            "demand_avg_nonzero": stats["demand_avg_nonzero"],
            "demand_std": stats["demand_std"],
            "demand_cv": stats["demand_cv"],
            "last_month_demand": stats["last_month_demand"],
            "forecast_next_month": stats["forecast_next_month"],
            "forecast_next_month_rounded": stats["forecast_next_month_rounded"],
            "forecast_method": stats["forecast_method"],
            "zeros_count": stats["zeros_count"],
            "zero_ratio": stats["zero_ratio"],
        })

    summary = prod_month.groupby(["item code", "iteam name"]).apply(_summarize).reset_index()

    # ضمان صف واحد لكل item code
    if summary["item code"].duplicated().any():
        summary = summary.sort_values(["item code", "demand_total"], ascending=[True, False])
        summary = summary.drop_duplicates(subset=["item code"], keep="first").reset_index(drop=True)

    # ABC
    total_all = summary["demand_total"].sum()
    if total_all > 0:
        summary = summary.sort_values("demand_total", ascending=False)
        summary["demand_share"] = summary["demand_total"] / total_all
        summary["cum_demand_share"] = summary["demand_share"].cumsum()

        def _abc(x: float) -> str:
            if x <= 0.8:
                return "A"
            elif x <= 0.95:
                return "B"
            else:
                return "C"

        summary["abc_class"] = summary["cum_demand_share"].apply(_abc)
    else:
        summary["demand_share"] = 0.0
        summary["cum_demand_share"] = 0.0
        summary["abc_class"] = "C"

    cols = [
        "item code", "iteam name",
        "months_count", "months_list",
        "demand_total", "demand_avg_all", "demand_avg_nonzero",
        "demand_std", "demand_cv",
        "last_month_demand",
        "forecast_next_month", "forecast_next_month_rounded",
        "forecast_method",
        "zeros_count", "zero_ratio",
        "demand_share", "cum_demand_share", "abc_class",
    ]
    return summary[cols].copy()
