import os
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import streamlit as st
from matplotlib.font_manager import FontProperties
from fpdf import FPDF
import io

# -----------------------
# 字体配置，确保中文显示
# -----------------------
FONT_PATH = "./NotoSansSC-Regular.otf"
if os.path.exists(FONT_PATH):
    my_font = FontProperties(fname=FONT_PATH)
else:
    my_font = None  # fallback

sns.set(style="whitegrid")
plt.rcParams['axes.unicode_minus'] = False

REQUIRED_COLS = ["姓名", "总分", "日期"]

st.title("📊 学生成绩分析工具 (含波动分析 & PDF报告)")

uploaded_file = st.file_uploader("请选择Excel文件", type=["xlsx", "xls"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"文件读取失败: {e}")
        st.stop()

    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        st.error(f"Excel缺少必要列: {missing}")
        st.stop()

    df["日期"] = pd.to_datetime(df["日期"], errors="coerce")
    df = df.dropna(subset=["姓名", "总分", "日期"])
    df = df.sort_values(by="日期")
    df["总分"] = pd.to_numeric(df["总分"], errors="coerce")
    df = df.dropna(subset=["总分"])

    st.success("✅ 文件加载成功！")
    student_name = st.selectbox("请选择学生姓名", df["姓名"].unique())

    if st.button("分析并生成报告"):
        stu = df[df["姓名"] == student_name].copy()
        if stu.empty:
            st.warning(f"未找到 {student_name} 的记录")
        else:
            # 计算波动指标
            std_dev = stu["总分"].std()
            mean_score = stu["总分"].mean()
            st.write(f"📈 {student_name} 历次成绩平均分: {mean_score:.2f}，波动标准差: {std_dev:.2f}")
            st.write("说明：标准差越小，成绩越稳定；标准差越大，成绩波动越大。")

            median_df = df.groupby("日期")["总分"].median().reset_index()

            # 绘图
            fig, ax = plt.subplots(figsize=(8, 5), dpi=120)
            dates = stu["日期"]
            scores = stu["总分"]
            median_dates = median_df["日期"]
            median_scores = median_df["总分"]

            ax.plot(dates, scores, marker='o', label=f"{student_name} 总分", color='red')
            ax.plot(median_dates, median_scores, marker='s', linestyle='--', label="班级总分中位数", color='blue')
            ax.fill_between(dates, scores - std_dev, scores + std_dev, color='red', alpha=0.2, label="成绩波动范围 (±1σ)")

            ax.set_title(f"{student_name} 历次成绩走势及波动分析", fontproperties=my_font)
            ax.set_xlabel("考试日期", fontproperties=my_font)
            ax.set_ylabel("总分", fontproperties=my_font)
            ax.grid(True)
            ax.legend(prop=my_font)

            st.pyplot(fig)

            # 历次成绩对比班级中位数表格
            compare_df = stu.merge(median_df, on="日期", suffixes=("_学生", "_班级中位数"))
            compare_df["高于中位数"] = compare_df["总分_学生"] - compare_df["总分_班级中位数"]
            compare_df["说明"] = compare_df["高于中位数"].apply(
                lambda x: "高于班级中位数" if x > 0 else ("低于班级中位数" if x < 0 else "与班级中位数持平")
            )
            st.subheader("📊 历次成绩与班级中位数对比表")
            st.dataframe(compare_df[["日期", "总分_学生", "总分_班级中位数", "说明"]])

            # -------------------
            # 生成 PDF
            # -------------------
            pdf = FPDF(orientation='P', unit='mm', format='A4')
            pdf.add_page()
            pdf.add_font("NotoSansSC", "", FONT_PATH, uni=True) if my_font else None
            pdf.set_font("NotoSansSC" if my_font else "Arial", size=14)

            pdf.cell(0, 10, f"{student_name} 成绩分析报告", ln=True, align='C')
            pdf.ln(5)
            pdf.multi_cell(0, 8, f"平均分: {mean_score:.2f}，成绩波动标准差: {std_dev:.2f}")
            pdf.ln(3)
            pdf.multi_cell(0, 8, "说明：标准差越小，成绩越稳定；标准差越大，成绩波动越大。")
            pdf.ln(5)

            # 保存图表到临时 buffer
            buf = io.BytesIO()
            fig.savefig(buf, format="png", bbox_inches="tight")
            buf.seek(0)
            pdf.image(buf, x=15, w=180)  # 插入图表
            pdf.ln(5)

            # 插入表格
            pdf.set_font("NotoSansSC" if my_font else "Arial", size=12)
            pdf.cell(0, 8, "历次成绩与班级中位数对比表", ln=True)
            pdf.ln(2)
            col_width = 45
            for idx, row in compare_df.iterrows():
                pdf.cell(col_width, 8, row["日期"].strftime("%Y-%m-%d"), border=1)
                pdf.cell(col_width, 8, str(row["总分_学生"]), border=1)
                pdf.cell(col_width, 8, str(row["总分_班级中位数"]), border=1)
                pdf.cell(col_width, 8, row["说明"], border=1)
                pdf.ln(8)

            # 提供下载
            pdf_buffer = io.BytesIO()
            pdf.output(pdf_buffer)
            pdf_buffer.seek(0)
            st.download_button(
                label="💾 下载完整 PDF 报告",
                data=pdf_buffer,
                file_name=f"{student_name}_成绩分析报告.pdf",
                mime="application/pdf"
            )
