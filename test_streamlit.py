import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import streamlit as st
import io
import os
from matplotlib.font_manager import FontProperties

# -----------------------
# 字体配置，确保中文显示
# -----------------------
# 请确保项目目录下有 NotoSansSC-Light.otf 文件
FONT_PATH = "NotoSansSC-Light.otf"
if os.path.exists(FONT_PATH):
    my_font = FontProperties(fname=FONT_PATH)
else:
    my_font = None  # fallback

sns.set(style="whitegrid")
plt.rcParams['axes.unicode_minus'] = False

# -----------------------
# 页面标题
# -----------------------
st.title("📊 学生成绩分析工具 (Web版)")

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

    # 检查列名
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        st.error(f"Excel缺少必要列: {missing}")
        st.stop()

    # 数据预处理
    df["日期"] = pd.to_datetime(df["日期"], errors="coerce")
    df["总分"] = pd.to_numeric(df["总分"], errors="coerce")
    df = df.dropna(subset=["姓名", "总分", "日期"])
    df = df.sort_values(by="日期")
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
            # 1️⃣ 成绩趋势图 + 班级中位数
            median_df = df.groupby("日期")["总分"].median().reset_index()

            fig, ax = plt.subplots(figsize=(8, 5), dpi=120)
            ax.plot(stu["日期"], stu["总分"], marker='o', label=f"{student_name} 总分")
            ax.plot(median_df["日期"], median_df["总分"], marker='s', linestyle='--', label="班级中位数")
            ax.set_title(f"{student_name} 历次成绩走势", fontproperties=my_font)
            ax.set_xlabel("考试日期", fontproperties=my_font)
            ax.set_ylabel("总分", fontproperties=my_font)
            ax.grid(True)
            ax.legend(prop=my_font)
            st.pyplot(fig)

            # 2️⃣ 成绩增长率表格
            stu['上次成绩'] = stu['总分'].shift(1)
            stu['增长率(%)'] = ((stu['总分'] - stu['上次成绩']) / stu['上次成绩'] * 100).round(2)
            st.write(f"📈 {student_name} 各次考试成绩变化：")
            st.dataframe(stu[['日期', '总分', '上次成绩', '增长率(%)']].fillna('-'))

            # 3️⃣ 班级成绩分布（箱线图）
            fig2, ax2 = plt.subplots(figsize=(8,5), dpi=120)
            sns.boxplot(x='日期', y='总分', data=df, ax=ax2)
            sns.scatterplot(x='日期', y='总分', data=stu, color='red', s=100, label=student_name)
            ax2.set_title(f"{student_name} 与班级成绩分布", fontproperties=my_font)
            ax2.set_xlabel("考试日期", fontproperties=my_font)
            ax2.set_ylabel("总分", fontproperties=my_font)
            ax2.legend(prop=my_font)
            st.pyplot(fig2)

            # 4️⃣ 成绩区间统计
            bins = [0, 60, 70, 80, 90, 100]
            labels = ['<60','60-69','70-79','80-89','90-100']
            stu['成绩区间'] = pd.cut(stu['总分'], bins=bins, labels=labels, right=False)
            count_interval = stu['成绩区间'].value_counts().reindex(labels)
            fig3, ax3 = plt.subplots(figsize=(6,4), dpi=120)
            count_interval.plot(kind='bar', color='skyblue', ax=ax3)
            ax3.set_title(f"{student_name} 成绩区间统计", fontproperties=my_font)
            ax3.set_xlabel("分数区间", fontproperties=my_font)
            ax3.set_ylabel("次数", fontproperties=my_font)
            st.pyplot(fig3)

            # 5️⃣ 排名百分比
            df['排名'] = df.groupby('日期')['总分'].rank(ascending=False, method='min')
            df['总人数'] = df.groupby('日期')['总分'].transform('count')
            df['百分比'] = ((df['总人数'] - df['排名']) / df['总人数'] * 100).round(2)
            stu_rank = df[df['姓名']==student_name][['日期','总分','排名','总人数','百分比']]
            st.write(f"🏆 {student_name} 各次考试排名与百分比：")
            st.dataframe(stu_rank)

            # -----------------------
            # 导出图表和报告
            # -----------------------
            # 图表 PNG
            buf_fig = io.BytesIO()
            fig.savefig(buf_fig, format='png', bbox_inches='tight')
            buf_fig2 = io.BytesIO()
            fig2.savefig(buf_fig2, format='png', bbox_inches='tight')
            buf_fig3 = io.BytesIO()
            fig3.savefig(buf_fig3, format='png', bbox_inches='tight')

            st.download_button(
                label="💾 下载成绩趋势图",
                data=buf_fig.getvalue(),
                file_name=f"{student_name}_成绩趋势.png",
                mime="image/png"
            )

            st.download_button(
                label="💾 下载班级分布图",
                data=buf_fig2.getvalue(),
                file_name=f"{student_name}_班级分布.png",
                mime="image/png"
            )

            st.download_button(
                label="💾 下载成绩区间统计图",
                data=buf_fig3.getvalue(),
                file_name=f"{student_name}_成绩区间统计.png",
                mime="image/png"
            )

            # Excel 报告
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                stu.to_excel(writer, sheet_name='成绩变化', index=False)
                stu_rank.to_excel(writer, sheet_name='排名百分比', index=False)
            st.download_button(
                label="💾 下载分析报告 (Excel)",
                data=output.getvalue(),
                file_name=f"{student_name}_成绩分析报告.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
