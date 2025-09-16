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
            # 分数趋势变化（折线 + 中位数变化）
            # -----------------------
            stu = stu.sort_values(by="日期")  # 确保有序
            median_df = median_df.sort_values(by="日期")
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

            def explain_diff(x):
                if x > 0:
                    return "高于班级中位数"
                elif x < 0:
                    return "低于班级中位数"
                else:
                    return "等于班级中位数"

            compare_df["解释"] = compare_df["与班级中位数差"].apply(explain_diff)
            st.subheader("📋 历次成绩对比班级中位数")
            st.dataframe(compare_df[["日期", "总分_学生", "总分_班级中位数", "与班级中位数差", "解释"]])

            # -----------------------
            # 波动分析（学生标准差 + 中位数标准差）
            # -----------------------
            student_std = float(stu["总分"].std())
            median_std = float(median_df["总分"].std())  # 所有考试中位数的标准差

            st.subheader("📏 成绩波动分析")
            st.write(f"学生历次成绩标准差: **{student_std:.2f}**")
            st.write(f"班级每次考试中位数标准差: **{median_std:.2f}**")
            st.write("说明：学生标准差大于中位数标准差 -> 学生成绩波动幅度高于班级中位数的波动，可能较不稳定；反之则较稳定。")

            # -----------------------
            # 生成 PDF 报告（修正：兼容 output 返回类型 & 用临时文件插入图片）
            # -----------------------
            pdf = FPDF()
            pdf.add_page()
            pdf.set_auto_page_break(auto=True, margin=15)

            # 加载中文字体（若存在），否则回退 Arial
            if os.path.exists(FONT_PATH):
                try:
                    pdf.add_font("Noto", "", FONT_PATH, uni=True)
                    pdf.set_font("Noto", "", 12)
                except Exception:
                    pdf.set_font("Arial", "", 12)
            else:
                pdf.set_font("Arial", "", 12)

            # 标题与波动信息
            pdf.cell(0, 10, f"{student_name} 成绩分析报告", ln=True, align="C")
            pdf.ln(5)
            pdf.cell(0, 8, f"学生历次成绩标准差: {student_std:.2f}", ln=True)
            pdf.cell(0, 8, f"班级中位数标准差: {median_std:.2f}", ln=True)
            pdf.ln(5)

            # 将图保存到临时文件（避免某些 fpdf 版本对 file-like 不支持）
            tmp_files = []
            try:
                # fig1 -> temp file
                tmp1 = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                fig1.savefig(tmp1.name, format="png", bbox_inches="tight", dpi=150)
                tmp1.close()
                tmp_files.append(tmp1.name)

                # fig2 -> temp file
                tmp2 = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                fig2.savefig(tmp2.name, format="png", bbox_inches="tight", dpi=150)
                tmp2.close()
                tmp_files.append(tmp2.name)

                # 插入图片（按页宽缩放）
                page_w = pdf.w - 2 * pdf.l_margin
                # first figure
                pdf.image(tmp1.name, x=pdf.l_margin, w=page_w)
                pdf.ln(5)
                # second figure
                pdf.image(tmp2.name, x=pdf.l_margin, w=page_w)
                pdf.ln(5)

                # 写入成绩对比表（每行用 multi_cell，限制宽度）
                # 添加成绩对比表格
                pdf.cell(0, 10, "历次成绩对比班级中位数:", ln=True)
                pdf.ln(3)
                for idx, row in compare_df.iterrows():
                    pdf.cell(0, 8, f"{row['日期'].strftime('%Y-%m-%d')} 学生:{row['总分_学生']} 班级中位数:{row['总分_班级中位数']} 差:{row['与班级中位数差']} ({row['解释']})", ln=True)
    
                pdf_buf = io.BytesIO()
                pdf.output(pdf_buf)
                pdf_buf.seek(0)
                
                st.download_button(
                    label="💾 下载完整 PDF 报告",
                    data=pdf_buf,
                    file_name=f"{student_name}_成绩分析报告.pdf",
                    mime="application/pdf"
                )

            finally:
                # 清理临时文件
                for f in tmp_files:
                    try:
                        os.remove(f)
                    except Exception:
                        pass
