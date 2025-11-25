import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import os
import tempfile

import statsmodels.api as sm
from statsmodels.formula.api import ols
from statsmodels.stats.multicomp import pairwise_tukeyhsd

from fpdf import FPDF

# 中文字体路径（你必须上传到 GitHub）
FONT_PATH = "SourceHanSansSC-Regular.otf"

# Matplotlib 中文设置
if os.path.exists(FONT_PATH):
    plt.rcParams['font.sans-serif'] = ['Source Han Sans SC']
else:
    plt.rcParams['font.sans-serif'] = ['SimHei']

plt.rcParams['axes.unicode_minus'] = False


# 自动识别因素列（非数字）与数值列（最后一个数字列）
def detect_factors(df):
    numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
    if len(numeric_cols) == 0:
        return [], None
    value_col = numeric_cols[-1]             # 最后一列作为因变量
    factor_cols = df.columns.tolist()
    factor_cols.remove(value_col)
    return factor_cols, value_col


# 运行 ANOVA
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


# Tukey 事后检验
def tukey_test(df, factor, value_col):
    return pairwise_tukeyhsd(df[value_col], df[factor])


# 绘图
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
    ax1.set_title("箱线图")
    plt.xticks(rotation=25)
    plots.append(fig1)

    # 均值折线图
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
            dashes=False,
            ax=ax2
        )
    ax2.set_title("均值折线图")
    plt.xticks(rotation=25)
    plots.append(fig2)

    return plots


# PDF 生成
def generate_pdf(anova_table, tukey_text, plots):
    pdf = FPDF()
    pdf.add_page()

    # 中文字体
    if os.path.exists(FONT_PATH):
        pdf.add_font("SourceHan", "", FONT_PATH, uni=True)
        pdf.set_font("SourceHan", size=14)
    else:
        pdf.set_font("Arial", size=14)

    pdf.cell(0, 10, "方差分析报告（自动生成）", ln=True)
    pdf.ln(5)

    # ANOVA 表格
    pdf.set_font("SourceHan" if os.path.exists(FONT_PATH) else "Arial", size=12)
    pdf.cell(0, 8, "一、ANOVA 检验结果：", ln=True)

    pdf.set_font("Arial", size=9)
    for line in anova_table.to_string().split("\n"):
        pdf.cell(0, 5, line, ln=True)

    pdf.ln(5)

    # Tukey
    pdf.set_font("SourceHan" if os.path.exists(FONT_PATH) else "Arial", size=12)
    pdf.cell(0, 8, "二、Tukey 事后检验：", ln=True)

    pdf.set_font("Arial", size=9)
    for line in tukey_text.split("\n"):
        pdf.cell(0, 5, line, ln=True)

    # 插入图像
    for fig in plots:
        temp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
        fig.savefig(temp.name, dpi=150, bbox_inches="tight")

        pdf.add_page()
        pdf.image(temp.name, x=10, y=20, w=180)

    pdf_path = "anova_report.pdf"
    pdf.output(pdf_path)
    return pdf_path


# Streamlit 界面
st.title("📊 自动 ANOVA 方差分析工具（1-3 因素）")

uploaded = st.file_uploader("上传 Excel 文件", type=["xlsx", "xls"])

if uploaded:
    df = pd.read_excel(uploaded)
    st.success("文件读取成功！")
    st.dataframe(df.head())

    # 自动识别列
    factors, value_col = detect_factors(df)
    if not factors:
        st.error("未检测到因素列")
        st.stop()

    st.info(f"识别到 {len(factors)} 个因素：{factors}\n数值列：{value_col}")

    # ANOVA
    anova_table, model = run_anova(df, factors, value_col)
    st.subheader("📌 ANOVA 结果")
    st.dataframe(anova_table)

    # Tukey
    st.subheader("📌 Tukey 事后检验")
    tukey_text = ""
    for f in factors:
        st.write(f"### 因素：{f}")
        tukey = tukey_test(df, f, value_col)
        st.text(tukey.summary())
        tukey_text += f"\n\n因素：{f}\n{tukey.summary()}"

    # 图表
    st.subheader("📌 图表展示")
    plots = create_plots(df, factors, value_col)
    for fig in plots:
        st.pyplot(fig)

    # PDF
    if st.button("📥 生成 PDF 报告"):
        pdf_path = generate_pdf(anova_table, tukey_text, plots)
        with open(pdf_path, "rb") as f:
            st.download_button("点击下载 PDF", f, file_name="anova_report.pdf")

