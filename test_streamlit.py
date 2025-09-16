import os
import io
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import streamlit as st
from matplotlib.font_manager import FontProperties

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

            fig, ax = plt.subplots(figsize=(8, 5), dpi=120)
            sns.lineplot(x="日期", y="总分", data=stu, marker='o', ax=ax, label=f"{student_name} 总分")
            sns.lineplot(x="日期", y="总分", data=median_df, marker='s', linestyle='--', ax=ax, label="班级总分中位数")

            # 中文显示
            if my_font:
                ax.set_title(f"{student_name} 历次成绩走势", fontproperties=my_font)
                ax.set_xlabel("考试日期", fontproperties=my_font)
                ax.set_ylabel("总分", fontproperties=my_font)
                ax.legend(prop=my_font)
            else:
                ax.set_title(f"{student_name} 历次成绩走势")
                ax.set_xlabel("考试日期")
                ax.set_ylabel("总分")
                ax.legend()

            plt.xticks(rotation=45)
            st.subheader("📈 历史成绩走势")
            st.pyplot(fig)

            # -----------------------
            # 分数趋势变化
            # -----------------------
            stu["分数变化"] = stu["总分"].diff()
            st.subheader("📊 分数趋势变化")
            st.line_chart(stu.set_index("日期")["分数变化"])

            # -----------------------
            # 历次成绩对比班级中位数表格
            # -----------------------
            compare_df = stu.merge(median_df, on="日期", suffixes=("_学生", "_班级中位数"))
            compare_df["与班级中位数差"] = compare_df["总分_学生"] - compare_df["总分_班级中位数"]

            # 添加解释列
            def explain_diff(x):
                if x > 0:
                    return "高于班级中位数"
                elif x < 0:
                    return "低于班级中位数"
                else:
                    return "等于班级中位数"

            compare_df["解释"] = compare_df["与班级中位数差"].apply(explain_diff)

            # 显示表格
            st.subheader("📋 历次成绩对比班级中位数")
            st.dataframe(compare_df[["日期", "总分_学生", "总分_班级中位数", "与班级中位数差", "解释"]])
