import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import statsmodels.api as sm
from statsmodels.formula.api import ols
from statsmodels.stats.multicomp import pairwise_tukeyhsd
from fpdf import FPDF
import os
import tempfile

# ============================
#   字体设置（支持中文）
# ============================

FONT_PATH = "./SourceHanSansSC-Regular.otf"  # 请把字体放在和 app.py 同级目录

# Matplotlib 中文
if os.path.exists(FONT_PATH):
    plt.rcParams['font.sans-serif'] = ['Source Han Sans SC']
else:
    plt.rcParams['font.sans-serif'] = ['SimHei']

plt.rcParams['axes.unicode_minus'] = False


# ============================
#   判断因素 + 数值列
# ============================

def detect_factors(df):
    numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
    if len(numeric_cols) == 0:
        return [], None
    value_col = numeric_cols[-1]
    factor_cols = df.columns.tolist()
    factor_cols.remove(value_col)
    return factor_cols, value_col


# ============================
#   自动 ANOVA（1~3因素）
# ============================

def run_anova(df, factors, value_col):
    if len(factors) == 1:
        formula = f"{value_col} ~ C({factors[0]})"
    elif len(factors) == 2:
        f1, f2 = factors
        formula = f"{value_col} ~ C({f1}) * C({f2})"
    else:
        f1, f2, f3 = factors[:3]
        formula = f"{value_col} ~ C({f1}) * C({f2}) * C({f3})"

    model = ols(formula, data=df).fit()
    anova_table = sm.stats.anova_lm(model, typ=2)
    return anova_table, model


# ============================
#   Tukey 多重比较
# ============================

def tukey_test(df, factor, value_col):
    return pairwise_tukeyhsd(df[value_col], df[factor])


# ============================
#   可视化图
# ============================

def create_plots(df, factors, value_col):
    plots = []

    # 箱线图
    fig1, ax1 = plt.subplots(figsize=(10, 6))
    sns.boxplot(
        data=df,
        x=factors[0],
        y=value_col,
        hue=factors[1] if len(factors) > 1 else None,
        ax=ax1
    )
    ax1.set_title(f"箱线图（因素：{' × '.join(factors)}）")
    plt.xticks(rotation=30)
    plots.append(fig1)

    # 均值折线
    mean_df = df.groupby(factors)[value_col].mean().reset_index()
    fig2, ax2 = plt.subplots(figsize=(10, 6))
    if len(factors) == 1:
        sns.lineplot(data=mean_df, x=factors[0], y=value_col, marker="o", ax=ax2)
    else:
        sns.lineplot(
            data=mean_df,
            x=factors[0],
            y=value_col,
            hue=factors[1],
            style=factors[1],
            markers=True,
            ax=ax2
        )
    ax2.set_title(f"均值折线图（因素：{' × '.join(factors)}）")
    plt.xticks(rotation=30)
    plots.append(fig2)

    return plots


# ============================
#   PDF 报告（完全中文支持）
# ============================

def generate_pdf(anova_table, tukey_results, plots):
    pdf = FPDF()
    pdf.add_page()

    # 中文字体
    if os.path.exists(FONT_PATH):
        pdf.add_font("CN", "", FONT_PATH, uni=True)
        pdf.set_font("CN", size=14)
    else:
        pdf.add_font("CN", "", "./SimHei.ttf", uni=True)
        pdf.set_font("CN", size=14)

    # 标题
    pdf.cell(0, 10, "方差分析报告（自动生成）", ln=True)

    # ANOVA 表
    pdf.ln(5)
    pdf.set_font("CN", size=12)
    pdf.cell(0, 8, "一、ANOVA 检验结果：", ln=True)

    pdf.set_font("CN", size=10)
    for row in anova_table.to_string().split("\n"):
        pdf.cell(0, 6, row, ln=True)

    # Tukey 表
    pdf.ln(5)
    pdf.set_font("CN", size=12)
    pdf.cell(0, 8, "二、Tukey 事后检验：", ln=True)

    pdf.set_font("CN", size=10)
    for txt in tukey_results.split("\n"):
        pdf.cell(0, 5, txt, ln=True)

    # 图表添加进 PDF
    for fig in plots:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
            fig.savefig(tmp.name, dpi=150, bbox_inches="tight")
            pdf.add_page()
            pdf.image(tmp.name, x=10, y=20, w=180)

    pdf_path = "anova_report.pdf"
    pdf.output(pdf_path)
    return pdf_path


# ============================
#   Streamlit 主体
# ============================

st.title("📊 自动 ANOVA 方差分析工具（1~3 因素）")

uploaded = st.file_uploader("上传 Excel 文件", type=["xlsx", "xls"])

if uploaded:
    df = pd.read_excel(uploaded)
    st.success("文件读取成功！")
    st.dataframe(df.head())

    # 自动识别因素
    factors, value_col = detect_factors(df)
    if len(factors) == 0:
        st.error("❌ 未检测到因素列，请检查文件。")
        st.stop()

    st.info(f"🔍 自动检测到：{len(factors)} 个因素 → {factors}，数值列 → {value_col}")

    # ANOVA
    anova_table, model = run_anova(df, factors, value_col)
    st.subheader("📌 ANOVA 结果")
    st.dataframe(anova_table)

    # Tukey
    st.subheader("📌 Tukey 事后检验")
    tukey_output = ""
    for f in factors:
        st.write(f"### 因素：{f}")
        tukey = tukey_test(df, f, value_col)
        st.text(tukey.summary())
        tukey_output += f"\n\n因素：{f}\n{tukey.summary()}"

    # 生成图表
    plots = create_plots(df, factors, value_col)
    st.subheader("📌 图表展示")
    for fig in plots:
        st.pyplot(fig)

    # PDF 导出
    if st.button("📄 生成 PDF 报告"):
        pdf_path = generate_pdf(anova_table, tukey_output, plots)
        with open(pdf_path, "rb") as f:
            st.download_button("📥 点击下载 PDF 报告", f, file_name="anova_report.pdf")
