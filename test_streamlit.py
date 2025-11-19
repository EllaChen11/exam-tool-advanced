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
    # 功能 1：按日期查看前 5 / 后 5
    # -----------------------
    st.subheader("📌 按日期查看成绩排名（前 5 名 / 后 5 名）")

    all_dates = sorted(df["日期"].unique())
    select_date = st.selectbox("请选择考试日期", all_dates, format_func=lambda x: x.strftime("%Y-%m-%d"))

    df_one = df[df["日期"] == select_date].sort_values("总分", ascending=False)

    col1, col2 = st.columns(2)

    with col1:
        st.write("🏆 **前 5 名（Top 5）**")
        st.dataframe(df_one.head(5)[["姓名", "总分"]])

    with col2:
        st.write("🐢 **后 5 名（Last 5）**")
        st.dataframe(df_one.tail(5)[["姓名", "总分"]])

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

            ax1.set_title(f"{student_name} 历次成绩走势")
            ax1.set_xlabel("考试日期")
            ax1.set_ylabel("总分")
            plt.xticks(rotation=45)
            st.subheader("📈 历史成绩走势")
            st.pyplot(fig1)

            # -----------------------
            # 分数趋势变化（折线 + 中位数变化）
            # -----------------------
            stu = stu.sort_values(by="日期")  # 确保有序
            median_df = median_df.sort_values(by="日期")
            stu["分数变化"] = stu["总分"].diff()
            median_df["分数变化"] = median_df["总分"].diff()

            fig2, ax2 = plt.subplots(figsize=(8, 4), dpi=120)
            sns.lineplot(x="日期", y="分数变化", data=stu, marker='o', ax=ax2, label=f"{student_name} 分数变化")
            sns.lineplot(x="日期", y="分数变化", data=median_df, marker='s', linestyle='--', ax=ax2,
                         label="班级中位数分数变化")

            ax2.set_title(f"{student_name} 分数趋势变化")
            ax2.set_xlabel("考试日期")
            ax2.set_ylabel("分数变化")
            plt.xticks(rotation=45)
            st.subheader("📊 分数趋势变化")
            st.pyplot(fig2)

            # -----------------------
            # 功能 2：材料题 / 选择题折线图
            # -----------------------
            st.subheader("📝 材料题 & 选择题 成绩走势")

            need_cols = ["材料", "选择"]
            missing_sub_cols = [c for c in need_cols if c not in df.columns]

            if missing_sub_cols:
                st.warning(f"Excel 缺少以下题型列，无法绘制材料题/选择题折线图：{missing_sub_cols}")
            else:
                df["材料题"] = pd.to_numeric(df["材料题"], errors="coerce")
                df["选择题"] = pd.to_numeric(df["选择题"], errors="coerce")
                stu["材料题"] = pd.to_numeric(stu["材料题"], errors="coerce")
                stu["选择题"] = pd.to_numeric(stu["选择题"], errors="coerce")

                median_sub = df.groupby("日期")[["材料题", "选择题"]].median().reset_index()

                # -------- 材料题 --------
                fig_mat, ax_mat = plt.subplots(figsize=(8, 4), dpi=120)
                sns.lineplot(x="日期", y="材料题", data=stu, marker='o', ax=ax_mat, label=f"{student_name} 材料题")
                sns.lineplot(x="日期", y="材料题", data=median_sub, marker='s', linestyle='--',
                             ax=ax_mat, label="班级材料题中位数")

                ax_mat.set_title("材料题成绩走势")
                ax_mat.set_xlabel("考试日期")
                ax_mat.set_ylabel("材料题分数")
                plt.xticks(rotation=45)
                st.pyplot(fig_mat)

                # -------- 选择题 --------
                fig_sel, ax_sel = plt.subplots(figsize=(8, 4), dpi=120)
                sns.lineplot(x="日期", y="选择题", data=stu, marker='o', ax=ax_sel, label=f"{student_name} 选择题")
                sns.lineplot(x="日期", y="选择题", data=median_sub, marker='s', linestyle='--',
                             ax=ax_sel, label="班级选择题中位数")

                ax_sel.set_title("选择题成绩走势")
                ax_sel.set_xlabel("考试日期")
                ax_sel.set_ylabel("选择题分数")
                plt.xticks(rotation=45)
                st.pyplot(fig_sel)

            # -----------------------
            # 历次成绩对比班级中位数表格
            # -----------------------
            compare_df = stu.merge(median_df, on="日期", suffixes=("_学生", "_班级中位数"))
            compare_df["与班级中位数差"] = compare_df["总分_学生"] - compare_df["总分_班级中位数"]
            compare_df["解释"] = compare_df["与班级中位数差"].apply(
                lambda x: "高于班级中位数" if x > 0 else ("低于班级中位数" if x < 0 else "等于班级中位数")
            )
            st.subheader("📋 历次成绩对比班级中位数")
            st.dataframe(compare_df[["日期", "总分_学生", "总分_班级中位数", "与班级中位数差", "解释"]])

            # -----------------------
            # 波动分析
            # -----------------------
            student_std = float(stu["总分"].std())
            median_std = float(median_df["总分"].std())
            st.subheader("📏 成绩波动分析")
            st.write(f"学生历次成绩标准差: **{student_std:.2f}**")
            st.write(f"班级中位数标准差: **{median_std:.2f}**")

            # -----------------------
            # PDF 导出
            # -----------------------
            pdf = FPDF()
            pdf.add_page()
            pdf.set_auto_page_break(auto=True, margin=15)

            if os.path.exists(FONT_PATH):
                try:
                    pdf.add_font("Noto", "", FONT_PATH, uni=True)
                    pdf.set_font("Noto", "", 12)
                except:
                    pdf.set_font("Arial", "", 12)
            else:
                pdf.set_font("Arial", "", 12)

            pdf.cell(0, 10, f"{student_name} 成绩分析报告", ln=True, align="C")
            pdf.ln(5)
            pdf.cell(0, 8, f"学生成绩标准差: {student_std:.2f}", ln=True)
            pdf.cell(0, 8, f"班级中位数标准差: {median_std:.2f}", ln=True)
            pdf.ln(5)

            # 保存图表为图片并插入 PDF
            tmp_files = []
            try:
                def save_fig_temp(fig):
                    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                    fig.savefig(tmp.name, dpi=150, bbox_inches="tight")
                    tmp.close()
                    tmp_files.append(tmp.name)
                    return tmp.name

                img1 = save_fig_temp(fig1)
                img2 = save_fig_temp(fig2)

                pdf.image(img1, x=pdf.l_margin, w=pdf.w - 2 * pdf.l_margin)
                pdf.ln(5)
                pdf.image(img2, x=pdf.l_margin, w=pdf.w - 2 * pdf.l_margin)
                pdf.ln(5)

                pdf_buf = io.BytesIO()
                pdf.output(pdf_buf)
                pdf_buf.seek(0)

                st.download_button(
                    label="💾 下载 PDF 成绩报告",
                    data=pdf_buf,
                    file_name=f"{student_name}_成绩报告.pdf",
                    mime="application/pdf"
                )

            finally:
                for f in tmp_files:
                    try:
                        os.remove(f)
                    except:
                        pass
