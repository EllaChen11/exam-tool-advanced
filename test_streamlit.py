import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.font_manager import FontProperties
import seaborn as sns
import numpy as np
from datetime import datetime
import os
import io
import tempfile
from fpdf import FPDF

# ==================== 1. 终极字体方案（永不报错）====================
# 项目根目录放以下任意一个字体文件即可完美显示中文：
# NotoSansSC-Regular.otf（推荐） / SimHei.ttf / msyh.ttc

pdf_font_path = "./NotoSansSC-Regular.otf"  # 请确保项目目录下有该字体文件
if os.path.exists(pdf_font_path):
    my_font = FontProperties(fname=pdf_font_path)
else:
    my_font = None  # fallback

sns.set(style="whitegrid")
plt.rcParams['axes.unicode_minus'] = False  # 负号正常显示


# ==================== 2. Streamlit 页面配置 ====================
st.set_page_config(page_title="学生成绩智能分析系统", page_icon="📊", layout="wide")
st.title("📊 学生成绩智能分析系统")
st.markdown("### 支持多次考试 | 自动偏科诊断 | 一键导出PDF报告")

# ==================== 3. 文件上传 ====================
uploaded_file = st.file_uploader("请上传 Excel 成绩单（必须包含列：姓名、选择、材料、总分、日期）", type=["xlsx", "xls"])

if not uploaded_file:
    st.info("请上传数据文件开始分析～")
    st.stop()


# ==================== 4. 数据加载与清洗 ====================
@st.cache_data
def load_data(file):
    df = pd.read_excel(file)
    required = ["姓名", "选择", "材料", "总分", "日期"]
    if not all(col in df.columns for col in required):
        st.error(f"缺少必要列：{[c for c in required if c not in df.columns]}")
        st.stop()

    df["日期"] = pd.to_datetime(df["日期"], errors="coerce")
    for c in ["选择", "材料", "总分"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df = df.dropna(subset=required)
    df["日期_str"] = df["日期"].dt.strftime("%Y-%m-%d")
    return df.sort_values("日期").reset_index(drop=True)


df = load_data(uploaded_file)
st.success(f"成功加载 {len(df)} 条记录，涉及 {df['姓名'].nunique()} 名学生")

# ==================== 5. 班级概览 ====================
st.markdown("---")
st.header("班级成绩概览")
date_options = sorted(df["日期_str"].unique(), reverse=True)
selected_date = st.selectbox("选择考试日期查看排名", date_options)

exam_df = df[df["日期_str"] == selected_date].copy()
exam_df = exam_df.sort_values("总分", ascending=False).reset_index(drop=True)
exam_df["排名"] = exam_df.index + 1

c1, c2 = st.columns(2)
with c1:
    st.subheader("总分前五名")
    st.dataframe(exam_df.head(5)[["排名", "姓名", "选择", "材料", "总分"]], use_container_width=True)
with c2:
    st.subheader("需重点关注（后五名）")
    st.dataframe(exam_df.tail(5)[["排名", "姓名", "选择", "材料", "总分"]], use_container_width=True)

# ==================== 6. 学生个人深度分析 ====================
st.markdown("---")
st.header("👤 学生个人深度诊断")
student_name = st.selectbox("请选择学生", sorted(df["姓名"].unique()))

stu = df[df["姓名"] == student_name].sort_values("日期").reset_index(drop=True)
if len(stu) == 0:
    st.warning("该学生无成绩记录")
    st.stop()

# 计算排名历史
ranks = []
for date in stu["日期"]:
    day_data = df[df["日期"] == date]
    rank = day_data[day_data["姓名"] == student_name].index[0] - \
           day_data.sort_values("总分", ascending=False).index[0] + 1
    ranks.append(rank)
stu["排名"] = ranks

# 关键指标
latest = stu.iloc[-1]
progress = stu["总分"].diff().iloc[-1] if len(stu) > 1 else 0
progress_text = f"较上次 +{progress:.0f}分" if progress > 0 else f"较上次 {progress:.0f}分" if progress < 0 else "持平"

# 偏科诊断
choice_gap = stu["选择"].mean() - stu["材料"].mean()
if abs(choice_gap) >= 8:
    bias = f"严重偏科（{'材料题' if choice_gap > 0 else '选择题'}弱）"
elif abs(choice_gap) >= 5:
    bias = f"轻度偏科（{'材料题' if choice_gap > 0 else '选择题'}较弱）"
else:
    bias = "成绩均衡"

# 指标卡
col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("最近总分", f"{latest['总分']}", progress_text)
col2.metric("历史平均", f"{stu['总分'].mean():.1f}")
col3.metric("历史最高", f"{stu['总分'].max():.0f}")
col4.metric("最近排名", f"第 {latest.name + 1} 名" if len(exam_df) > 0 else "未知")
col5.metric("偏科诊断", bias)

# ==================== 图表1：双Y轴 成绩+排名趋势 ====================
st.subheader("📈 成绩与排名趋势")
fig, ax1 = plt.subplots(figsize=(12, 6))
ax1.plot(stu["日期"], stu["总分"], 'o-', label="总分", linewidth=3, markersize=8)
ax1.plot(stu["日期"], stu["选择"], 's--', label="选择题", alpha=0.8)
ax1.plot(stu["日期"], stu["材料"], '^--', label="材料题", alpha=0.8)
ax1.set_ylabel("分数")
ax1.legend(loc="upper left")
if zh_font:
    ax1.set_title(f"{student_name} 成绩趋势", fontproperties=zh_font, fontsize=16)

ax2 = ax1.twinx()
ax2.plot(stu["日期"], stu["排名"], 'D-', color="#9b59b6", label="排名", linewidth=3)
ax2.invert_yaxis()
ax2.set_ylabel("排名（数值越小越好）")
ax2.legend(loc="upper right")

ax1.xaxis.set_major_formatter(mdates.DateFormatter('%m-%d'))
plt.xticks(rotation=45)
plt.tight_layout()
st.pyplot(fig)

# ==================== 图表2：雷达图 ====================
st.subheader("🕸️ 能力雷达图（与班级平均对比）")
categories = ['选择题', '材料题']
stu_vals = [stu["选择"].mean(), stu["材料"].mean()]
class_vals = [df["选择"].mean(), df["材料"].mean()]

angles = np.linspace(0, 2 * np.pi, len(categories), endpoint=False).tolist()
angles += angles[:1]
stu_vals += stu_vals[:1]
class_vals += class_vals[:1]

fig2, ax = plt.subplots(figsize=(6, 6), subplot_kw=dict(projection='polar'))
ax.plot(angles, stu_vals, 'o-', linewidth=3, label=student_name, color='#e74c3c')
ax.fill(angles, stu_vals, alpha=0.25, color='#e74c3c')
ax.plot(angles, class_vals, 's--', linewidth=2, label='班级平均', color='#3498db')
ax.set_xticks(angles[:-1])
ax.set_xticklabels(categories)
if zh_font:
    ax.set_title("能力雷达图", fontproperties=zh_font, pad=20)
ax.legend(loc='upper right', bbox_to_anchor=(1.3, 1.0))
st.pyplot(fig2)

# ==================== 智能建议 ====================
st.subheader("🧠 智能诊断建议")
suggestions = []
if latest["总分"] < df["总分"].quantile(0.3):
    suggestions.append("🔴 成绩位于班级下游，建议制定专项提升计划")
if latest["总分"] > df["总分"].quantile(0.8):
    suggestions.append("🟢 成绩优秀！已进入第一梯队，继续保持可冲击年级前3！")
if abs(choice_gap) >= 8:
    weak = "材料题" if choice_gap > 0 else "选择题"
    suggestions.append(f"🔴 严重偏科！{weak}拖后腿明显，需重点突破")
if len(stu) >= 2 and progress >= 5:
    suggestions.append("🟢 最近进步显著！学习状态极佳，继续加油！")

for s in suggestions:
    st.markdown(f"**{s}**")

# ==================== PDF 报告生成 ====================
st.markdown("---")
st.subheader("📄 一键生成并下载PDF诊断报告")

if st.button("🚀 生成个人PDF报告", type="primary"):
    with st.spinner("正在生成精美PDF报告..."):
        pdf = FPDF()
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=15)

        # 添加中文字体
        if pdf_font_path and os.path.exists(pdf_font_path):
            pdf.add_font("Chinese", "", pdf_font_path, uni=True)
            pdf.set_font("Chinese", size=12)
        else:
            pdf.set_font("Arial", size=12)

        pdf.set_font(size=18, style='B')
        pdf.cell(0, 15, f"{student_name} 成绩诊断报告", ln=1, align='C')
        pdf.set_font(size=12)
        pdf.cell(0, 10, f"生成时间：{datetime.now().strftime('%Y年%m月%d日 %H:%M')}", ln=1)
        pdf.ln(5)

        # 保存图片
        tmp_files = []
        fig.savefig("temp_trend.png", dpi=150, bbox_inches='tight')
        fig2.savefig("temp_radar.png", dpi=150, bbox_inches='tight')
        tmp_files.extend(["temp_trend.png", "temp_radar.png"])

        pdf.image("temp_trend.png", w=180)
        pdf.ln(10)
        pdf.image("temp_radar.png", w=100, x=50)
        pdf.ln(15)

        pdf.set_font(size=14, style='B')
        pdf.cell(0, 10, "智能诊断结论", ln=1)
        pdf.set_font(size=12)
        for s in suggestions:
            pdf.multi_cell(0, 8, "• " + s.replace("🔴", "警告").replace("🟢", "优秀"))

        # 输出
        pdf_buffer = io.BytesIO()
        pdf.output(pdf_buffer)
        pdf_buffer.seek(0)

        st.success("PDF报告生成成功！")
        st.download_button(
            label="📥 点击下载完整报告",
            data=pdf_buffer,
            file_name=f"{student_name}_成绩诊断报告_{datetime.now().strftime('%Y%m%d')}.pdf",
            mime="application/pdf"
        )

        # 清理临时图片
        for f in tmp_files:
            if os.path.exists(f):
                os.remove(f)
