import os
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import streamlit as st
from matplotlib.font_manager import FontProperties

# -----------------------
# 字体配置，确保中文显示
# -----------------------
# 请确保项目目录下有 NotoSansSC-Regular.otf 文件
FONT_PATH = "./NotoSansSC-Regular.otf"
if os.path.exists(FONT_PATH):
    my_font = FontProperties(fname=FONT_PATH)
    plt.rcParams['font.family'] = my_font.get_name()
else:
    my_font = None  # fallback

sns.set(style="whitegrid")
plt.rcParams['axes.unicode_minus'] = False

# -----------------------
# Streamlit 页面标题
# -----------------------
st.title("📊 学生成绩分析工具")

# -----------------------
# 上传 Excel 文件
# -----------------------
REQUIRED_COLS = ["姓名", "总分", "日期"]
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
            # 历史成绩走势折线图
            # -----------------------
            fig, ax = plt.subplots(figsize=(8, 5), dpi=120)
            ax.plot(stu["日期"], stu["总分"], marker='o', label=f"{student_name} 总分")
            ax.set_title(f"{student_name} 历次成绩走势")
            ax.set_xlabel("考试日期")
            ax.set_ylabel("总分")
            ax.grid(True)
            ax.legend()
            st.pyplot(fig)

            # -----------------------
            # 分数趋势散点图
            # -----------------------
            median_df = df.groupby("日期")["总分"].median().reset_index()
            fig2, ax2 = plt.subplots(figsize=(8, 5), dpi=120)
            sns.scatterplot(x="日期", y="总分", data=stu, color='red', s=100, label=student_name)
            sns.scatterplot(x="日期", y="总分", data=median_df, color='blue', s=100, label="班级中位数")
            ax2.set_title(f"{student_name} 分数趋势变化")
            ax2.set_xlabel("考试日期")
            ax2.set_ylabel("总分")
            ax2.grid(True)
            ax2.legend()
            st.pyplot(fig2)

            # -----------------------
            # 历次成绩对比班级中位数表格
            # -----------------------
            compare_df = pd.merge(
                stu[['日期','总分']],
                median_df.rename(columns={"总分":"班级中位数"}),
                on="日期"
            )
            compare_df["与班级中位数差"] = compare_df["总分"] - compare_df["班级中位数"]

            # 添加解释列
            def explain_diff(x):
                if x > 0:
                    return "高于班级中位数"  # 说明该次成绩比班级中位数高，表现较好
                elif x < 0:
                    return "低于班级中位数"  # 说明该次成绩比班级中位数低，需要改进
                else:
                    return "等于班级中位数"  # 说明该次成绩与班级中位数持平

            compare_df["解释"] = compare_df["与班级中位数差"].apply(explain_diff)

            # 彩色显示表格
            def highlight_row(row):
                if row["解释"] == "高于班级中位数":
                    return ['background-color: #c6f6d5']*len(row)  # 绿色
                elif row["解释"] == "低于班级中位数":
                    return ['background-color: #fed7d7']*len(row)  # 红色
                else:
                    return ['background-color: #fefcbf']*len(row)  # 黄色

            st.subheader("📋 历次成绩对比班级中位数")
            st.dataframe(compare_df.style.apply(highlight_row, axis=1))
