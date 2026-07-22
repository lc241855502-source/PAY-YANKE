import streamlit as st
import pandas as pd
import io
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
st.markdown("上传 3 张基础数据表，自动核算订单业绩明细与人员奖金，一键导出结果 Excel。")

# ============================================================
# 工具函数：生成可下载的 Excel 文件
# ============================================================
def to_excel_download(sheet1_df: pd.DataFrame, sheet2_df: pd.DataFrame) -> bytes:
    """将两个 DataFrame 写入同一个 Excel 文件的不同 Sheet，返回二进制流"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        sheet1_df.to_excel(writer, sheet_name="订单业绩明细", index=False)
        sheet2_df.to_excel(writer, sheet_name="人员奖金汇总", index=False)
    return output.getvalue()

# ============================================================
# 核心计算1：每笔订单对应的业绩（SHEET1）
# ============================================================
def calculate_order_performance(
    df_orders: pd.DataFrame,
    df_products: pd.DataFrame
) -> pd.DataFrame:
    """
    输入：
    df_orders  : 订单明细表
    df_products: 产品基准信息表
    输出：每笔订单的业绩核算结果
    """
    # 1. 关联产品基准数据
    df = df_orders.merge(
        df_products[["产品名称", "基准单价", "业绩系数", "产品品类"]],
        on="产品名称",
        how="left"
    )

    # 2. 计算基准业绩
    df["基准业绩"] = df["数量"] * df["基准单价"] * df["业绩系数"]

    # 3. 客户类型系数规则（可根据实际业务调整）
    customer_coef_map = {
        "医院": 1.0,
        "经销商": 0.95,
        "直客": 1.05
    }
    df["客户类型系数"] = df["客户类型"].map(customer_coef_map).fillna(1.0)

    # 4. 按回款率折算最终业绩
    df["最终核算业绩"] = df["基准业绩"] * df["客户类型系数"] * df["回款率"]

    # 5. 数值四舍五入
    df["基准业绩"] = df["基准业绩"].round(2)
    df["最终核算业绩"] = df["最终核算业绩"].round(2)
    df["客户类型系数"] = df["客户类型系数"].round(3)

    # 定义输出列顺序
    output_cols = [
        "订单号", "订单日期", "客户名称", "客户类型", "客户类型系数",
        "产品名称", "产品品类", "数量", "基准单价", "业绩系数",
        "基准业绩", "回款率", "最终核算业绩", "签约人", "所属团队"
    ]
    # 过滤掉不存在的列，避免报错
    output_cols = [col for col in output_cols if col in df.columns]
    return df[output_cols]

# ============================================================
# 核心计算2：每个人的奖金 + 奖金池情况（SHEET2）
# ============================================================
def calculate_staff_bonus(
    df_performance: pd.DataFrame,
    df_staff: pd.DataFrame,
    total_bonus_pool: float = 0.0
) -> pd.DataFrame:
    """
    输入：
    df_performance: 订单业绩明细结果（SHEET1输出）
    df_staff      : 人员信息表
    total_bonus_pool: 总奖金池金额
    输出：人员奖金汇总表 + 奖金池汇总行
    """
    # 1. 按签约人汇总业绩
    staff_summary = (
        df_performance
        .groupby("签约人", as_index=False)
        .agg(
            订单笔数=("订单号", "count"),
            个人总业绩=("最终核算业绩", "sum")
        )
    )

    # 2. 关联人员基础信息与提成比例
    df = staff_summary.merge(
        df_staff[["姓名", "岗位", "提成比例", "所属团队"]],
        left_on="签约人",
        right_on="姓名",
        how="left"
    ).drop(columns=["姓名"])

    # 3. 计算基础奖金
    df["基础奖金"] = (df["个人总业绩"] * df["提成比例"]).round(2)

    # 4. 奖金池分配规则：按个人基础奖金占比分摊
    total_base_bonus = df["基础奖金"].sum()
    if total_bonus_pool > 0 and total_base_bonus > 0:
        df["奖金池分配额"] = (df["基础奖金"] / total_base_bonus * total_bonus_pool).round(2)
    else:
        df["奖金池分配额"] = 0.0

    # 5. 实发奖金合计
    df["实发奖金合计"] = (df["基础奖金"] + df["奖金池分配额"]).round(2)

    # 6. 追加奖金池汇总行
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
# 主界面：文件上传区
# ============================================================
st.subheader("一、上传基础数据表")
col1, col2, col3 = st.columns(3)

with col1:
    file_orders = st.file_uploader("① 订单明细表", type=["xlsx", "xls"])
with col2:
    file_products = st.file_uploader("② 产品基准表", type=["xlsx", "xls"])
with col3:
    file_staff = st.file_uploader("③ 人员信息表", type=["xlsx", "xls"])

# 参数配置区
st.subheader("二、奖金池参数配置")
bonus_pool_input = st.number_input(
    "本期总奖金池金额（元）",
    min_value=0.0,
    value=0.0,
    step=1000.0,
    format="%.2f",
    help="若无需奖金池分配，保持0即可"
)

# ============================================================
# 执行计算
# ============================================================
if st.button("开始核算", type="primary", use_container_width=True):
    if not all([file_orders, file_products, file_staff]):
        st.error("请上传全部 3 张基础表格后再执行计算")
    else:
        try:
            with st.spinner("正在读取数据并核算中..."):
                # 读取三张表（默认读取第一个Sheet）
                df_orders = pd.read_excel(file_orders)
                df_products = pd.read_excel(file_products)
                df_staff = pd.read_excel(file_staff)

                # 计算Sheet1：订单业绩明细
                df_sheet1 = calculate_order_performance(df_orders, df_products)

                # 计算Sheet2：人员奖金+奖金池
                df_sheet2 = calculate_staff_bonus(df_sheet1, df_staff, bonus_pool_input)

            st.success("✅ 核算完成！")

            # 结果预览标签页
            tab1, tab2 = st.tabs(["📋 SHEET1 订单业绩明细", "💰 SHEET2 人员奖金汇总"])
            with tab1:
                st.dataframe(df_sheet1, use_container_width=True, hide_index=True)
                st.caption(f"共计 {len(df_sheet1)} 条订单记录")
            with tab2:
                st.dataframe(df_sheet2, use_container_width=True, hide_index=True)
                st.caption(f"共计 {len(df_sheet2)-1} 名人员 + 1 行汇总统计")

            # 导出下载
            st.subheader("三、导出核算结果")
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
            st.info("请检查表头列名是否与系统要求一致，或调整对应字段名称。")

# ============================================================
# 使用说明
# ============================================================
with st.expander("📖 表格字段要求与使用说明"):
    st.markdown("""
    ### 三张输入表必填字段
    | 表格名称 | 核心字段 |
    |---------|---------|
    | 订单明细表 | 订单号、订单日期、客户名称、客户类型、产品名称、数量、回款率、签约人、所属团队 |
    | 产品基准表 | 产品名称、基准单价、业绩系数、产品品类 |
    | 人员信息表 | 姓名、岗位、提成比例、所属团队 |

    ### 业务规则说明
    1. **订单业绩核算**：`最终业绩 = 数量 × 基准单价 × 业绩系数 × 客户类型系数 × 回款率`
    2. **客户类型系数**：默认医院1.0、经销商0.95、直客1.05，可在代码中修改
    3. **人员奖金核算**：`实发奖金 = 个人总业绩 × 提成比例 + 奖金池分配额`
    4. **奖金池分配**：按个人基础奖金占总基础奖金的比例进行分摊

    ### 调整规则方法
    直接修改代码中 `calculate_order_performance` 和 `calculate_staff_bonus` 函数内的系数与逻辑即可。
    """)