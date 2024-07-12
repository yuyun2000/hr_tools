import pandas as pd
from openpyxl import load_workbook

file = './data/福州2024年06月考勤.xls'
# 读取原始xls表格
df = pd.read_excel(file, skiprows=1)

# 数据处理
# 确认列名
date_col = '刷卡日期'
time_col = '刷卡时间'
dept_col = '部门'
id_col = '工号'
name_col = '姓名'
datafrom = '数据来源'

# 将刷卡日期和刷卡时间合并成一个完整的时间戳
df[date_col] = pd.to_datetime(df[date_col])
df[time_col] = df[time_col].astype(str)  # 将刷卡时间转换为字符串格式
df['完整时间'] = pd.to_datetime(df[date_col].dt.strftime('%Y-%m-%d') + ' ' + df[time_col])

# 找出同一天的第一和最后一次打卡时间
df_grouped = df.groupby([id_col, name_col, dept_col, date_col]).agg(
    签到时间=('完整时间', 'first'),
    签退时间=('完整时间', 'last')
).reset_index()

# 构造新的DataFrame
new_df = pd.DataFrame()
new_df['工号'] = df_grouped[id_col]
new_df['名字'] = df_grouped[name_col]
new_df['部门名称'] = df_grouped[dept_col]
new_df['状态'] = df_grouped[date_col].apply(lambda x: '[周末]' if x.weekday() >= 5 else '正常')
new_df['假别'] = ''
new_df['出勤'] = df_grouped[date_col].dt.strftime('%Y-%m-%d')
new_df['上班时间'] = '09:00'
new_df['下班时间'] = '18:30'

# 将签到时间和签退时间转换为字符串
new_df['签到时间'] = df_grouped['签到时间'].dt.strftime('%Y-%m-%d %H:%M:%S')
new_df['签退时间'] = df_grouped['签退时间'].dt.strftime('%Y-%m-%d %H:%M:%S')

# 保存为新的xlsx表格
new_df.to_excel('./data/fz.xlsx', index=False)
