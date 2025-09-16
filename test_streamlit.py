import os
import io
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import streamlit as st
from matplotlib.font_manager import FontProperties
from fpdf import FPDF
from datetime import datetime

# -----------------------
# 字体配置，确保中文显示
# -----------------------
FONT_PATH = "./NotoSansSC-Regular.otf"  # 请确保项目目录下有该字体文件
if os.path.exists(FONT_PATH):
    my_font = FontProperties(fname=FONT_PATH)
else:
    my_font = None  # fallback

sns.set(style="whitegrid")
plt.rcParams['axes.unicode_minus'] = False  # 负号正常显示

# -----------------------
# 必要列
# -----------------------
REQUIRED_COLS = ["姓名", "总分", "日期"]

# -----------------------
# 页面标题
# -----------------------
st.title("📊 学生成绩分析工具 (Web版)")

# -----------------------
# 上传 Excel 文件
# -----------------------
uploaded_file = st.file_uploader("请选择Excel文件", type=["xlsx", "xls"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"文件读取失败: {e}")
        st.stop()

    # 检查必要列
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        st.error(f"Excel缺少必要列: {missing}")
        st.stop()

    # 数据预处理
    df["日期"] = pd.to_datetime(df["日期"], errors="coerce")
    df = df.dropna(subset=["姓名", "总分", "日期"])
    df = df.sort_values(by="日期")
    df["总分"] = pd.to_numeric(df["总分"], errors="coerce")
    df = df.dropna(subset=["总分"])

    st.success("✅ 文件加载成功！")

    # -----------------------
    # 选择学生
    # -----------------------
    student_name = st.selectbox("请选择学生姓名", df["姓名"].unique())

    if st.button("分析并绘图"):
        stu = df[df["姓名"] == student_name].copy()
        if stu.empty:
            st.warning(f"未找到 {student_name} 的记录")
        else:
            # -----------------------
            # 历史成绩走势
            # -----------------------
            median_df = df.groupby("日期")["总分"].median().reset_index()

            fig1, ax1 = plt.subplots(figsize=(8, 5), dpi=120)
            sns.lineplot(x="日期", y="总分", data=stu, marker='o', ax=ax1, label=f"{student_name} 总分")
            sns.lineplot(x="日期", y="总分", data=median_df, marker='s', linestyle='--', ax=ax1, label="班级总分中位数")

            if my_font:
                ax1.set_title(f"{student_name} 历次成绩走势", fontproperties=my_font)
                ax1.set_xlabel("考试日期", fontproperties=my_font)
                ax1.set_ylabel("总分", fontproperties=my_font)
                ax1.legend(prop=my_font)
            else:
                ax1.set_title(f"{student_name} 历次成绩走势")
                ax1.set_xlabel("考试日期")
                ax1.set_ylabel("总分")
                ax1.legend()

            plt.xticks(rotation=45)
            st.subheader("📈 历史成绩走势")
            st.pyplot(fig1)

            # -----------------------
            # 分数趋势变化（折线图 + 中位数变化）
            # -----------------------
            stu["分数变化"] = stu["总分"].diff()
            median_df["分数变化"] = median_df["总分"].diff()

            fig2, ax2 = plt.subplots(figsize=(8, 4), dpi=120)
            sns.lineplot(x="日期", y="分数变化", data=stu, marker='o', ax=ax2, label=f"{student_name} 分数变化")
            sns.lineplot(x="日期", y="分数变化", data=median_df, marker='s', linestyle='--', ax=ax2, label="班级中位数分数变化")

            if my_font:
                ax2.set_title(f"{student_name} 分数趋势变化", fontproperties=my_font)
                ax2.set_xlabel("考试日期", fontproperties=my_font)
                ax2.set_ylabel("分数变化", fontproperties=my_font)
                ax2.legend(prop=my_font)
            else:
                ax2.set_title(f"{student_name} 分数趋势变化")
                ax2.set_xlabel("考试日期")
                ax2.set_ylabel("分数变化")
                ax2.legend()

            plt.xticks(rotation=45)
            st.subheader("📊 分数趋势变化")
            st.pyplot(fig2)

            # -----------------------
            # 历次成绩对比班级中位数表格
            # -----------------------
            compare_df = stu.merge(median_df, on="日期", suffixes=("_学生", "_班级中位数"))
            compare_df["与班级中位数差"] = compare_df["总分_学生"] - compare_df["总分_班级中位数"]

            # 添加解释列
            def explain_diff(x):
                if x > 0:
                    return "高于班级中位数"  # 高于中位数表示成绩比班级大部分同学好
                elif x < 0:
                    return "低于班级中位数"
                else:
                    return "等于班级中位数"

            compare_df["解释"] = compare_df["与班级中位数差"].apply(explain_diff)

            st.subheader("📋 历次成绩对比班级中位数")
            st.dataframe(compare_df[["日期", "总分_学生", "总分_班级中位数", "与班级中位数差", "解释"]])

            # -----------------------
            # 波动分析（标准差）
            # -----------------------
            score_std = stu["总分"].std()
            st.subheader("📏 成绩波动分析")
            st.write(f"该生历次成绩标准差（波动幅度）为: **{score_std:.2f}**，标准差越大表示成绩波动越大")

            # -----------------------
            # 生成 PDF 报告
            # -----------------------
            pdf = FPDF()
            pdf.add_page()
            pdf.set_auto_page_break(auto=True, margin=15)
            if os.path.exists(FONT_PATH):
                pdf.add_font("Noto", "", FONT_PATH, uni=True)
                pdf.set_font("Noto", "", 12)
            else:
                pdf.set_font("Arial", "", 12)

            pdf.cell(0, 10, f"{student_name} 成绩分析报告", ln=True, align="C")
            pdf.ln(5)

            pdf.cell(0, 10, f"成绩波动（标准差）: {score_std:.2f}", ln=True)
            pdf.ln(5)

            # 保存图像到内存
            buf1 = io.BytesIO()
            fig1.savefig(buf1, format="png", bbox_inches="tight")
            buf1.seek(0)

            buf2 = io.BytesIO()
            fig2.savefig(buf2, format="png", bbox_inches="tight")
            buf2.seek(0)

            pdf.image(buf1, x=10, y=None, w=180)
            pdf.ln(85)
            pdf.image(buf2, x=10, y=None, w=180)
            pdf.ln(85)

            pdf.cell(0, 10, "历次成绩对比班级中位数:", ln=True)
            pdf.ln(3)
            for idx, row in compare_df.iterrows():
                pdf.cell(0, 8, f"{row['日期'].strftime('%Y-%m-%d')} 学生:{row['总分_学生']} 班级中位数:{row['总分_班级中位数']} 差:{row['与班级中位数差']} ({row['解释']})", ln=True)

            pdf_buf = io.BytesIO()
            pdf.output(pdf_buf)
            pdf_buf.seek(0)

            st.download_button(
                label="💾 下载PDF报告",
                data=pdf_buf,
                file_name=f"{student_name}_成绩分析报告.pdf",
                mime="application/pdf"
            )
