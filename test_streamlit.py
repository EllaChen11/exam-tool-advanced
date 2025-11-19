import os
import io
import tempfile
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import streamlit as st
from matplotlib.font_manager import FontProperties
from fpdf import FPDF

# -----------------------
# 字体配置，确保中文显示
# -----------------------
FONT_PATH = "./NotoSansSC-Regular.otf"
if os.path.exists(FONT_PATH):
    my_font = FontProperties(fname=FONT_PATH)
else:
    my_font = None

sns.set(style="whitegrid")
plt.rcParams['axes.unicode_minus'] = False

REQUIRED_COLS = ["姓名", "总分", "日期", "选择", "材料"]

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

    # 必要列检查
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        st.error(f"Excel缺少必要列: {missing}")
        st.stop()

    # 数据预处理
    df["日期"] = pd.to_datetime(df["日期"], errors="coerce")
    df = df.dropna(subset=["姓名", "总分", "日期"])
    df["总分"] = pd.to_numeric(df["总分"], errors="coerce")
    df["选择"] = pd.to_numeric(df["选择"], errors="coerce")
    df["材料"] = pd.to_numeric(df["材料"], errors="coerce")
    df = df.dropna()

    st.success("✅ 文件加载成功！")

    # -----------------------
    # 日期前 5 名 & 后 5 名
    # -----------------------
    st.subheader("📆 按日期查看前五名 & 后五名")

    available_dates = sorted(df["日期"].dt.date.unique())
    selected_date = st.selectbox("请选择日期", available_dates)

    df_day = df[df["日期"].dt.date == selected_date].sort_values(by="总分", ascending=False)

    st.markdown("### 🥇 前五名")
    st.table(df_day.head(5)[["姓名", "选择", "材料", "总分"]])

    st.markdown("### 🪫 后五名")
    st.table(df_day.tail(5)[["姓名", "选择", "材料", "总分"]])

    st.markdown("---")

    # -----------------------
    # 选择学生
    # -----------------------
    student_name = st.selectbox("请选择学生姓名", df["姓名"].unique())

    if st.button("分析并绘图"):
        stu = df[df["姓名"] == student_name].copy()

        # =======================
        # 总分走势
        # =======================
        st.subheader("📈 总分历史走势")

        median_df = df.groupby("日期")["总分"].median().reset_index()

        fig1, ax1 = plt.subplots(figsize=(8, 5), dpi=120)
        sns.lineplot(x="日期", y="总分", data=stu, marker='o', ax=ax1, label=f"{student_name} 总分")
        sns.lineplot(x="日期", y="总分", data=median_df, marker='s', linestyle='--',
                     ax=ax1, label="班级总分中位数")
        plt.xticks(rotation=45)
        st.pyplot(fig1)

        # =======================
        # 分数变化
        # =======================
        st.subheader("📉 分数变化趋势")
        stu = stu.sort_values(by="日期")
        median_df = median_df.sort_values(by="日期")
        stu["分数变化"] = stu["总分"].diff()
        median_df["分数变化"] = median_df["总分"].diff()

        fig2, ax2 = plt.subplots(figsize=(8, 4), dpi=120)
        sns.lineplot(x="日期", y="分数变化", data=stu, marker='o', ax=ax2)
        sns.lineplot(x="日期", y="分数变化", data=median_df, marker='s', linestyle='--', ax=ax2)
        plt.xticks(rotation=45)
        st.pyplot(fig2)

        # =======================
        # 新增：选择题折线图
        # =======================
        st.subheader("📘 选择题成绩走势（含中位数）")

        median_sel = df.groupby("日期")["选择"].median().reset_index()
        fig_sel, ax_sel = plt.subplots(figsize=(8, 4), dpi=120)
        sns.lineplot(x="日期", y="选择", data=stu, marker='o', ax=ax_sel, label=f"{student_name} 选择题")
        sns.lineplot(x="日期", y="选择", data=median_sel, marker='s', linestyle='--',
                     ax=ax_sel, label="班级选择题中位数")
        plt.xticks(rotation=45)
        st.pyplot(fig_sel)

        # =======================
        # 新增：材料题折线图
        # =======================
        st.subheader("📙 材料题成绩走势（含中位数）")

        median_mat = df.groupby("日期")["材料"].median().reset_index()
        fig_mat, ax_mat = plt.subplots(figsize=(8, 4), dpi=120)
        sns.lineplot(x="日期", y="材料", data=stu, marker='o', ax=ax_mat, label=f"{student_name} 材料题")
        sns.lineplot(x="日期", y="材料", data=median_mat, marker='s', linestyle='--',
                     ax=ax_mat, label="班级材料题中位数")
        plt.xticks(rotation=45)
        st.pyplot(fig_mat)

        # =======================
        # 对比表格
        # =======================
        compare_df = stu.merge(median_df, on="日期", suffixes=("_学生", "_班级中位数"))
        compare_df["差值"] = compare_df["总分_学生"] - compare_df["总分_班级中位数"]

        st.subheader("📋 历次成绩对比班级中位数")
        st.dataframe(compare_df)

        # =======================
        # 波动分析
        # =======================
        st.subheader("📏 成绩波动分析")
        st.write(f"学生成绩标准差: **{stu['总分'].std():.2f}**")
        st.write(f"班级中位数标准差: **{median_df['总分'].std():.2f}**")

        # =======================
        # 生成 PDF
        # =======================
        pdf = FPDF()
        pdf.add_page()

        if os.path.exists(FONT_PATH):
            pdf.add_font("Noto", "", FONT_PATH, uni=True)
            pdf.set_font("Noto", "", 12)
        else:
            pdf.set_font("Arial", "", 12)

        pdf.cell(0, 10, f"{student_name} 成绩分析报告", ln=True, align="C")

        # 保存图片到临时文件
        tmp_files = []
        for fig in [fig1, fig2]:
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
            fig.savefig(tmp.name, dpi=150, bbox_inches="tight")
            tmp.close()
            tmp_files.append(tmp.name)

        # 插入 PDF
        for img in tmp_files:
            pdf.image(img, x=10, w=190)
            pdf.ln(5)

        pdf_buf = io.BytesIO()
        pdf.output(pdf_buf)
        pdf_buf.seek(0)

        st.download_button(
            "📥 下载 PDF 报告",
            pdf_buf,
            f"{student_name}_成绩报告.pdf",
            mime="application/pdf"
        )

        for f in tmp_files:
            os.remove(f)
