import pandas as pd

xlsx1 = './data/20240601-202406280000十楼睿云联考勤.xlsx'
xlsx2 = './data/fz.xlsx'
# 读取第一个表格
df1 = pd.read_excel(xlsx1)

# 读取第二个表格
df2 = pd.read_excel(xlsx2)

# 选择第一个表格的列
columns_df1 = ['工号', '名字', '部门名称', '状态', '假别', '出勤', '上班时间', '下班时间', '签到时间', '签退时间']

# 选择第二个表格的列
columns_df2 = ['工号', '名字', '部门名称', '状态', '假别', '出勤', '上班时间', '下班时间', '签到时间', '签退时间']

# 筛选两个数据框中共同的列
common_columns = set(columns_df1).intersection(columns_df2)

# 根据共同的列创建新的数据框
df1_common = df1[list(common_columns)]
df2_common = df2[list(common_columns)]

# 合并两个数据框
merged_df = pd.concat([df1_common, df2_common], ignore_index=True)

# 保存到新的 XLSX 文件
merged_df.to_excel('merge.xlsx', index=False)

print('合成后的表格已保存为 "merge.xlsx"')
