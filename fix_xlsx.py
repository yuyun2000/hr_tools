'''
修复考勤表的工号问题
'''
import pandas as pd

# 读取工号对照品
df_ref = pd.read_excel('./data/人员工号信息.xlsx')
# 读取10楼考勤
df_data = pd.read_excel('./merge.xlsx')

# 创建一个名字到部门和工号的映射字典
name_dict = df_ref.groupby('姓名').apply(lambda x: x.to_dict(orient='records')).to_dict()

# 定义一个函数来修正工号和部门
def correct_data(row):
    name = row['名字']
    if name in name_dict:
        records = name_dict[name]
        if len(records) == 1:
            row['工号'] = records[0]['工号']
            row['部门名称'] = records[0]['部门']
        else:
            # 处理重名的情况
            for record in records:
                if record['部门'] == row['部门名称']:
                    row['工号'] = record['工号']
                    break
    return row

# 应用修正函数到每一行
df_corrected = df_data.apply(correct_data, axis=1)

df_corrected.to_excel('./merge_fixed.xlsx', index=False)
