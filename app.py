import io
import pandas as pd
import streamlit as st

from core_demand import (
    build_demand_long,
    build_branch_level_forecast,
    build_product_level_forecast,
)
from core_inventory import plan_inventory_dc, normalize_item_code


@st.cache_data
def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    # استخدام openpyxl لتفادي xlsxwriter
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
    return output.getvalue()


st.set_page_config(page_title="Demand & Inventory Planning", layout="wide")
st.title("Demand & Inventory Planning System")

# Reset
if "reset_id" not in st.session_state:
    st.session_state["reset_id"] = 0

if st.button("🔄 إعادة تعيين كل البيانات والبدء من جديد"):
    st.session_state["reset_id"] += 1
    for k in list(st.session_state.keys()):
        if k != "reset_id":
            del st.session_state[k]
    st.rerun()


# ==========================
# Step 1: Sales Upload
# ==========================
st.header("الخطوة 1: رفع ملفات المبيعات الشهرية")

uploaded_sales = st.file_uploader(
    "ارفع ملفات المبيعات (شهرية) - يمكن اختيار أكثر من ملف",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
    key=f"sales_uploader_{st.session_state['reset_id']}",
)

month_dfs = {}

if uploaded_sales:
    st.write("حدد اسم الشهر لكل ملف (مثال: 2025-01 أو 01-2025):")
    used_labels = set()

    for f in uploaded_sales:
        default_label = f.name.split(".")[0]

        col1, col2 = st.columns([2, 3])
        with col1:
            st.write(f"الملف: **{f.name}**")
        with col2:
            month_label = st.text_input(
                f"اسم الشهر للملف {f.name}",
                value=default_label,
                key=f"month_{f.name}",
            ).strip()

        if month_label:
            if month_label in used_labels:
                st.error(f"اسم الشهر مكرر: {month_label} — غيّره لاسم فريد حتى لا تتضاعف الداتا.")
                continue

            used_labels.add(month_label)
            df_wide = pd.read_excel(f)
            month_dfs[month_label] = df_wide

build_demand_btn = st.button("بناء داتا الطلب والتوقع", type="primary")

if build_demand_btn:
    if not month_dfs:
        st.error("لم يتم تحديد أي شهر. تأكد من إدخال اسم شهر لكل ملف.")
    else:
        with st.spinner("جاري تجهيز داتا الطلب والتوقع..."):
            df_long = build_demand_long(month_dfs)
            st.session_state["df_long"] = df_long

            df_branch_fc = build_branch_level_forecast(df_long)
            st.session_state["df_branch_fc"] = df_branch_fc

            df_product_fc = build_product_level_forecast(df_long)
            st.session_state["df_product_fc"] = df_product_fc

        st.success("تم بناء داتا الطلب والتوقع.")


# ==========================
# Show + Download
# ==========================
if "df_long" in st.session_state:
    st.subheader("ملخص داتا الطلب (Long Format)")
    st.dataframe(st.session_state["df_long"].head(30))

    st.download_button(
        "تحميل داتا الطلب (Long Format) كملف Excel",
        data=df_to_excel_bytes(st.session_state["df_long"]),
        file_name="demand_long_from_app.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_long",
    )

if "df_branch_fc" in st.session_state:
    st.subheader("تلخيص + توقع على مستوى الفرع (Branch Level)")
    df_branch_fc = st.session_state["df_branch_fc"]

    branches = sorted(df_branch_fc["branch"].unique())
    selected_branch = st.selectbox("اختر الفرع", branches)

    df_branch_selected = df_branch_fc[df_branch_fc["branch"] == selected_branch].copy()
    st.dataframe(df_branch_selected.head(50))

    st.download_button(
        f"تحميل ملف التلخيص + التوقع للفرع ({selected_branch})",
        data=df_to_excel_bytes(df_branch_selected),
        file_name=f"branch_level_forecast_{selected_branch}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_branch_selected",
    )

if "df_product_fc" in st.session_state:
    st.subheader("تلخيص + توقع على مستوى المنتج (Company Level)")
    df_product_fc = st.session_state["df_product_fc"]
    st.dataframe(df_product_fc.head(30))

    st.download_button(
        "تحميل ملف التوقع على مستوى المنتج",
        data=df_to_excel_bytes(df_product_fc),
        file_name="product_baseline_forecast_from_app.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_product",
    )

    st.subheader("Top 100 حسب آخر شهر")
    df_top100 = df_product_fc.sort_values("last_month_demand", ascending=False).head(100)
    st.dataframe(df_top100)

    st.download_button(
        "تحميل Top 100 (آخر شهر)",
        data=df_to_excel_bytes(df_top100),
        file_name="top100_last_month_products.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_top100",
    )


# ==========================
# Step 2: Inventory Upload
# ==========================
st.header("الخطوة 2: رفع ملف المخزون الرئيسي")

inv_file = st.file_uploader(
    "ارفع ملف المخزون (DC Inventory)",
    type=["xlsx", "xls"],
    key=f"inv_uploader_{st.session_state['reset_id']}",
)

days_per_month = st.number_input(
    "عدد أيام الشهر (للتحويل إلى طلب يومي)",
    min_value=1, max_value=31, value=30,
)

review_days = st.number_input(
    "Review Buffer (أيام فوق ROP)",
    min_value=0, max_value=60, value=7,
)

run_plan_btn = st.button("تشغيل تخطيط المخزون (Inventory Planning)")

if run_plan_btn:
    if "df_product_fc" not in st.session_state:
        st.error("قم أولًا ببناء داتا الطلب.")
    elif inv_file is None:
        st.error("لم يتم رفع ملف المخزون.")
    else:
        df_product_fc = st.session_state["df_product_fc"]
        df_inv = pd.read_excel(inv_file)

        # Debug keys
        df_fc_debug = df_product_fc.copy()
        df_inv_debug = df_inv.copy()

        df_fc_debug["item_code_key"] = df_fc_debug["item code"].apply(normalize_item_code)
        df_inv_debug["item_code_key"] = df_inv_debug["item code"].apply(normalize_item_code)

        n_fc = df_fc_debug["item_code_key"].nunique()
        n_inv = df_inv_debug["item_code_key"].nunique()

        inner_keys = pd.merge(
            df_fc_debug[["item_code_key"]].drop_duplicates(),
            df_inv_debug[["item_code_key"]].drop_duplicates(),
            on="item_code_key",
            how="inner",
        )
        n_common = inner_keys["item_code_key"].nunique()

        st.info(f"عدد أكواد التوقع: {n_fc} | عدد أكواد المخزون: {n_inv} | الأكواد المشتركة: {n_common}")

        if n_common == 0:
            st.error("لا يوجد أي كود مشترك بين Files. تأكد من مطابقة item code.")
        else:
            with st.spinner("جارٍ تشغيل تخطيط المخزون..."):
                df_plan = plan_inventory_dc(
                    df_product_fc,
                    df_inv,
                    days_per_month=days_per_month,
                    review_days=int(review_days),
                )
                st.session_state["df_plan"] = df_plan

            st.success("تم حساب خطة المخزون.")


# ==========================
# Results + Reports
# ==========================
if "df_plan" in st.session_state:
    st.subheader("نتائج تخطيط المخزون (Inventory Plan)")
    df_plan = st.session_state["df_plan"]
    st.dataframe(df_plan.head(60))

    st.download_button(
        "تحميل خطة المخزون كاملة",
        data=df_to_excel_bytes(df_plan),
        file_name="inventory_planning_dc_from_app.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_plan",
    )

    st.subheader("تقرير المنتجات التي تحتاج طلبية (Order)")
    df_order = df_plan[df_plan["inventory_decision"] == "Order"].copy()
    st.dataframe(df_order.head(100))
    st.download_button(
        "تحميل تقرير الطلبية (Order)",
        data=df_to_excel_bytes(df_order),
        file_name="PO_proposal_Order_only.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_po_order",
    )

    st.subheader("تقرير المنتجات في حالة Review")
    df_review = df_plan[df_plan["inventory_decision"] == "Review"].copy()
    st.dataframe(df_review.head(100))
    st.download_button(
        "تحميل تقرير Review",
        data=df_to_excel_bytes(df_review),
        file_name="inventory_review_items.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_review",
    )

    st.subheader("تقرير المنتجات في حالة Hold")
    df_hold = df_plan[df_plan["inventory_decision"] == "Hold"].copy()
    st.dataframe(df_hold.head(100))

    # No demand but has stock
    st.subheader("أصناف لا يوجد عليها طلب (Forecast=0) ولكن يوجد عليها مخزون")
    df_no_demand_with_stock = df_plan[
        (df_plan["forecast_next_month"] == 0) &
        (df_plan["available_qty_total"] > 0)
    ].copy()
    st.write(f"عدد الأصناف: {len(df_no_demand_with_stock)}")
    st.dataframe(df_no_demand_with_stock.head(100))
    st.download_button(
        "تحميل تقرير بدون طلب ولكن عليها مخزون",
        data=df_to_excel_bytes(df_no_demand_with_stock),
        file_name="no_demand_with_stock.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_no_demand_with_stock",
    )

    # Excess / Slow
    st.subheader("تقرير Excess / Slow Moving")
    dos_threshold = st.number_input("Days of Supply >=", min_value=30, max_value=3650, value=365)
    mask_excess = (
        (df_plan["days_of_supply"].notna()) &
        (df_plan["days_of_supply"] >= dos_threshold) &
        (df_plan["forecast_daily"] > 0)
    )
    df_excess = df_plan[mask_excess].copy()
    st.dataframe(df_excess.head(100))
    st.download_button(
        "تحميل تقرير Excess",
        data=df_to_excel_bytes(df_excess),
        file_name="excess_slow_moving_items.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_excess",
    )
