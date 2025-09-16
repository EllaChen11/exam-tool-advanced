# -----------------------
# 波动分析（标准差）完善版
# -----------------------
# 学生成绩标准差
score_std = stu["总分"].std()

# 每次考试班级标准差（反映班级波动）
class_std_df = df.groupby("日期")["总分"].std().reset_index().rename(columns={"总分": "班级标准差"})

# 合并到学生成绩表中
compare_df = compare_df.merge(class_std_df, on="日期")

st.subheader("📏 成绩波动分析")
st.write(f"- 该生历次成绩标准差（波动幅度）为: **{score_std:.2f}**")
st.write("- 每次考试班级成绩标准差如下表，学生成绩波动大于班级标准差可能表示波动幅度较大：")
st.dataframe(compare_df[["日期", "总分_学生", "班级标准差"]])
