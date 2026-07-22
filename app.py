import streamlit as st
import pandas as pd
import io
import difflib
from datetime import datetime

# ============================================================
# 页面配置
# ============================================================
st.set_page_config(
    page_title="眼科业绩奖金核算系统",
    page_icon="🏥",
    layout="wide"
)

st.title("🏥 眼科业务业绩与奖金自动核算系统")
st.markdown("上传 3 张基础数据表，自动识别列名、核算订单业绩与人员奖金，一键导出结果。")

# ============================================================
# 配置区：标准字段 & 别名关键词（可按需扩充）
# ============================================================
# 订单表标准字段 & 常见别名关键词
ORDER_FIELDS = {
    "订单号": ["订单号", "订单编号", "单号", "订单ID"],
    "订单日期": ["订单日期", "下单日期", "签约日期", "日期"],
    "客户名称": ["客户名称", "客户", "客户姓名", "单位名称", "医院名称"],
    "客户类型": ["客户类型", "客户分类", "渠道类型", "客户属性"],
    "产品名称": ["产品名称", "产品名", "品名", "商品名称", "项目名称", "产品"],
    "数量": ["数量", "销售数量", "件数", "台数"],
    "回款率": ["回款率", "回款比例", "到账率", "回款进度"],
    "签约人": ["签约人", "销售", "业务员", "负责人", "销售人员"],
    "所属团队": ["所属团队", "团队", "部门", "所属部门"]
}

# 产品表标准字段 & 常见别名关键词
PRODUCT_FIELDS = {
    "产品名称": ["产品名称", "产品名", "品名", "商品名称", "项目名称", "产品"],
    "基准单价": ["基准单价", "单价", "标准单价", "定价", "价格", "基准价"],
    "业绩系数": ["业绩系数", "系数", "提成系数", "核算系数"],
    "产品品类": ["产品品类", "品类", "产品分类", "类别", "产品类型"]
}

# 人员表标准字段 & 常见别名关键词
STAFF_FIELDS = {
    "姓名": ["姓名", "员工姓名", "名字", "人员姓名"],
    "岗位": ["岗位", "职位", "职级", "岗位名称"],
    "提成比例": ["提成比例", "提成率", "提成系数", "奖金比例", "提成点"],
    "所属团队": ["所属团队", "团队", "部门", "所属部门"]
}

# ============================================================
# 工具函数1：智能列名自动匹配
# ============================================================
def auto_match_columns(actual_columns: list, standard_alias: dict) -> dict:
    """
    根据实际列名，自动匹配到标准字段
    匹配优先级：完全匹配 > 包含关键词 > 字符串相似度
    """
    mapping = {}
    lower_cols = [str(col).strip() for col in actual_columns]
    
    for std_field, aliases in standard_alias.items():
        matched = None
        # 1. 优先完全匹配
        for alias in aliases:
            if alias in lower_cols:
                matched = actual_columns[lower_cols.index(alias)]
                break
        
        # 2. 其次包含关键词匹配
        if not matched:
            for col in actual_columns:
                col_lower = str(col).lower()
                if any(alias.lower() in col_lower for alias in aliases):
                    matched = col
                    break
        
        # 3. 最后相似度匹配
        if not matched:
            best_score = 0
            best_col = None
            for col in actual_columns:
                score = difflib.SequenceMatcher(
                    None, str(col).lower(), std_field.lower()
                ).ratio()
                if score > best_score and score > 0.5:
                    best_score = score
                    best_col = col
            matched = best_col
        
        mapping[std_field] = matched
    return mapping

# ============================================================
# 工具函数2：生成可下载Excel
# ============================================================
def to_excel_download(sheet1_df: pd.DataFrame, sheet2_df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        sheet1_df.to_excel(writer, sheet_name="订单业绩明细", index=False)
        sheet2_df.to_excel(writer, sheet_name="人员奖金汇总", index=False)
    return output.getvalue()

# ============================================================
# 核心计算1：每笔订单业绩（SHEET1）
# ============================================================
def calculate_order_performance(df_orders: pd.DataFrame, df_products: pd.DataFrame) -> pd.DataFrame:
    df = df_orders.merge(
        df_products[["产品名称", "基准单价", "业绩系数", "产品品类"]],
        on="产品名称",
        how="left"
    )

    df["基准业绩"] = df["数量"] * df["基准单价"] * df["业绩系数"]

    # 客户类型系数规则
    customer_coef_map = {"医院": 1.0, "经销商": 0.95, "直客": 1.05}
    df["客户类型系数"] = df["客户类型"].map(customer_coef_map).fillna(1.0)

    df["最终核算业绩"] = df["基准业绩"] * df["客户类型系数"] * df["回款率"]

    df["基准业绩"] = df["基准业绩"].round(2)
    df["最终核算业绩"] = df["最终核算业绩"].round(2)
    df["客户类型系数"] = df["客户类型系数"].round(3)

    output_cols = [
        "订单号", "订单日期", "客户名称", "客户类型", "客户类型系数",
        "产品名称", "产品品类", "数量", "基准单价", "业绩系数",
        "基准业绩", "回款率", "最终核算业绩", "签约人", "所属团队"
    ]
    output_cols = [col for col in output_cols if col in df.columns]
    return df[output_cols]

# ============================================================
# 核心计算2：人员奖金 + 奖金池（SHEET2）
# ============================================================
def calculate_staff_bonus(df_performance: pd.DataFrame, df_staff: pd.DataFrame, total_bonus_pool: float = 0.0) -> pd.DataFrame:
    staff_summary = (
        df_performance
        .groupby("签约人", as_index=False)
        .agg(
            订单笔数=("订单号", "count"),
            个人总业绩=("最终核算业绩", "sum")
        )
    )

    df = staff_summary.merge(
        df_staff[["姓名", "岗位", "提成比例", "所属团队"]],
        left_on="签约人",
        right_on="姓名",
        how="left"
    ).drop(columns=["姓名"])

    df["基础奖金"] = (df["个人总业绩"] * df["提成比例"]).round(2)

    total_base_bonus = df["基础奖金"].sum()
    if total_bonus_pool > 0 and total_base_bonus > 0:
        df["奖金池分配额"] = (df["基础奖金"] / total_base_bonus * total_bonus_pool).round(2)
    else:
        df["奖金池分配额"] = 0.0

    df["实发奖金合计"] = (df["基础奖金"] + df["奖金池分配额"]).round(2)

    summary_row = pd.DataFrame([{
        "签约人": "【奖金池汇总】",
        "订单笔数": df["订单笔数"].sum(),
        "个人总业绩": df["个人总业绩"].sum(),
        "岗位": "-",
        "提成比例": "-",
        "所属团队": "-",
        "基础奖金": df["基础奖金"].sum(),
        "奖金池分配额": df["奖金池分配额"].sum(),
        "实发奖金合计": df["实发奖金合计"].sum(),
    }])

    result = pd.concat([df, summary_row], ignore_index=True)
    return result

# ============================================================
# 主界面：文件上传
# ============================================================
st.subheader("一、上传基础数据表")
col1, col2, col3 = st.columns(3)

with col1:
    file_orders = st.file_uploader("① 订单明细表", type=["xlsx", "xls"])
with col2:
    file_products = st.file_uploader("② 产品基准表", type=["xlsx", "xls"])
with col3:
    file_staff = st.file_uploader("③ 人员信息表", type=["xlsx", "xls"])

# ============================================================
# 字段映射确认区（自动匹配 + 手动修正）
# ============================================================
mapping_ready = False
df_orders_raw = df_products_raw = df_staff_raw = None
order_mapping = product_mapping = staff_mapping = None

if all([file_orders, file_products, file_staff]):
    st.subheader("二、字段映射确认（自动识别，可手动修正）")
    st.caption("系统已自动匹配列名，如有错误请手动下拉选择对应字段")
    
    # 读取原始数据
    df_orders_raw = pd.read_excel(file_orders)
    df_products_raw = pd.read_excel(file_products)
    df_staff_raw = pd.read_excel(file_staff)

    # 自动匹配
    order_mapping_auto = auto_match_columns(df_orders_raw.columns.tolist(), ORDER_FIELDS)
    product_mapping_auto = auto_match_columns(df_products_raw.columns.tolist(), PRODUCT_FIELDS)
    staff_mapping_auto = auto_match_columns(df_staff_raw.columns.tolist(), STAFF_FIELDS)

    # 可编辑的映射表单
    with st.expander("📋 订单明细表字段映射", expanded=True):
        order_cols = df_orders_raw.columns.tolist()
        order_mapping = {}
        cols = st.columns(3)
        for i, (std_field, default_col) in enumerate(order_mapping_auto.items()):
            with cols[i % 3]:
                selected = st.selectbox(
                    f"{std_field}",
                    options=["-- 未匹配 --"] + order_cols,
                    index=0 if default_col is None else order_cols.index(default_col) + 1,
                    key=f"order_{std_field}"
                )
                order_mapping[std_field] = None if selected == "-- 未匹配 --" else selected

    with st.expander("📋 产品基准表字段映射", expanded=True):
        product_cols = df_products_raw.columns.tolist()
        product_mapping = {}
        cols = st.columns(2)
        for i, (std_field, default_col) in enumerate(product_mapping_auto.items()):
            with cols[i % 2]:
                selected = st.selectbox(
                    f"{std_field}",
                    options=["-- 未匹配 --"] + product_cols,
                    index=0 if default_col is None else product_cols.index(default_col) + 1,
                    key=f"product_{std_field}"
                )
                product_mapping[std_field] = None if selected == "-- 未匹配 --" else selected

    with st.expander("📋 人员信息表字段映射", expanded=True):
        staff_cols = df_staff_raw.columns.tolist()
        staff_mapping = {}
        cols = st.columns(2)
        for i, (std_field, default_col) in enumerate(staff_mapping_auto.items()):
            with cols[i % 2]:
                selected = st.selectbox(
                    f"{std_field}",
                    options=["-- 未匹配 --"] + staff_cols,
                    index=0 if default_col is None else staff_cols.index(default_col) + 1,
                    key=f"staff_{std_field}"
                )
                staff_mapping[std_field] = None if selected == "-- 未匹配 --" else selected

    # 校验必填字段
    required_order = ["订单号", "产品名称", "数量", "回款率", "签约人"]
    required_product = ["产品名称", "基准单价", "业绩系数"]
    required_staff = ["姓名", "提成比例"]

    missing = []
    for f in required_order:
        if not order_mapping.get(f):
            missing.append(f"订单表-{f}")
    for f in required_product:
        if not product_mapping.get(f):
            missing.append(f"产品表-{f}")
    for f in required_staff:
        if not staff_mapping.get(f):
            missing.append(f"人员表-{f}")

    if missing:
        st.warning(f"⚠️ 以下必填字段未匹配，请手动选择：{', '.join(missing)}")
        mapping_ready = False
    else:
        st.success("✅ 字段匹配完成，可以开始核算")
        mapping_ready = True

# ============================================================
# 奖金池参数 & 执行计算
# ============================================================
st.subheader("三、奖金池参数配置")
bonus_pool_input = st.number_input(
    "本期总奖金池金额（元）",
    min_value=0.0,
    value=0.0,
    step=1000.0,
    format="%.2f"
)

if st.button("开始核算", type="primary", use_container_width=True):
    if not mapping_ready:
        st.error("请先完成所有必填字段的映射后再核算")
    else:
        try:
            with st.spinner("正在重命名列并执行核算..."):
                # 按映射重命名为标准列名
                df_orders = df_orders_raw.rename(columns={v: k for k, v in order_mapping.items() if v})
                df_products = df_products_raw.rename(columns={v: k for k, v in product_mapping.items() if v})
                df_staff = df_staff_raw.rename(columns={v: k for k, v in staff_mapping.items() if v})

                # 数值列强制转数字
                numeric_cols_order = ["数量", "回款率"]
                numeric_cols_product = ["基准单价", "业绩系数"]
                numeric_cols_staff = ["提成比例"]
                
                for col in numeric_cols_order:
                    df_orders[col] = pd.to_numeric(df_orders[col], errors="coerce").fillna(0)
                for col in numeric_cols_product:
                    df_products[col] = pd.to_numeric(df_products[col], errors="coerce").fillna(0)
                for col in numeric_cols_staff:
                    df_staff[col] = pd.to_numeric(df_staff[col], errors="coerce").fillna(0)

                # 执行计算
                df_sheet1 = calculate_order_performance(df_orders, df_products)
                df_sheet2 = calculate_staff_bonus(df_sheet1, df_staff, bonus_pool_input)

            st.success("✅ 核算完成！")

            # 结果预览
            tab1, tab2 = st.tabs(["📋 SHEET1 订单业绩明细", "💰 SHEET2 人员奖金汇总"])
            with tab1:
                st.dataframe(df_sheet1, use_container_width=True, hide_index=True)
                st.caption(f"共计 {len(df_sheet1)} 条订单记录")
            with tab2:
                st.dataframe(df_sheet2, use_container_width=True, hide_index=True)
                st.caption(f"共计 {len(df_sheet2)-1} 名人员 + 1 行汇总统计")

            # 导出下载
            st.subheader("四、导出核算结果")
            excel_bytes = to_excel_download(df_sheet1, df_sheet2)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"业绩奖金核算结果_{timestamp}.xlsx"

            st.download_button(
                label="📥 下载结果 Excel 文件",
                data=excel_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        except Exception as e:
            st.error(f"核算出错：{str(e)}")
            st.info("请检查数值类字段是否包含非数字内容，或联系调整计算逻辑。")

# ============================================================
# 使用说明
# ============================================================
with st.expander("📖 使用说明"):
    st.markdown("""
    ### 功能说明
    1. **自动列名识别**：系统会根据关键词自动匹配表格列名，支持「产品名/品名/项目名称」等同义命名
    2. **手动修正**：自动匹配不准时，可在下拉框手动选择对应字段
    3. **必填字段**：带*为必填，其余字段缺失不影响核心计算
    4. **计算规则**：
       - 最终业绩 = 数量 × 基准单价 × 业绩系数 × 客户类型系数 × 回款率
       - 实发奖金 = 个人总业绩 × 提成比例 + 奖金池分配额
    """)
