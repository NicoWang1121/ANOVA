def generate_pdf(anova_table, tukey_results, plots):
    pdf = FPDF()
    pdf.add_page()

    # 统一中文字体
    if os.path.exists(FONT_PATH):
        pdf.add_font("CN", "", FONT_PATH, uni=True)
        pdf.set_font("CN", size=14)
    else:
        pdf.add_font("CN", "", "./SimHei.ttf", uni=True)
        pdf.set_font("CN", size=14)

    # 标题
    pdf.cell(0, 10, "方差分析报告（自动生成）", ln=True)

    # ANOVA 部分
    pdf.ln(5)
    pdf.set_font("CN", size=12)
    pdf.cell(0, 8, "一、ANOVA 检验结果：", ln=True)

    pdf.set_font("CN", size=10)
    for row in anova_table.to_string().split("\n"):
        pdf.cell(0, 5, row, ln=True)

    # Tukey 部分
    pdf.ln(5)
    pdf.set_font("CN", size=12)
    pdf.cell(0, 8, "二、Tukey 事后检验：", ln=True)

    pdf.set_font("CN", size=10)
    for txt in tukey_results.split("\n"):
        pdf.cell(0, 5, txt, ln=True)

    # 图表
    for fig in plots:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
            fig.savefig(tmp.name, dpi=150, bbox_inches="tight")
            pdf.add_page()
            pdf.image(tmp.name, x=10, y=20, w=180)

    pdf_path = "anova_report.pdf"
    pdf.output(pdf_path)
    return pdf_path
