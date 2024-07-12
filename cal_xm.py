import pandas as pd
from datetime import datetime, timedelta

# 读取考勤数据
df = pd.read_excel('./merge_fixed.xlsx')

# 初始化结果数据框、异常情况记录和餐补记录
results = []
abnormal_records = []
meal_allowance_records = []

# 定义处理函数
def process_attendance(group):
    overtime_total = 0
    abnormal_count = 0
    meal_allowance = 0
    prev_sign_out = None

    for i, row in group.iterrows():
        on_time = datetime.strptime(row['上班时间'], "%H:%M")
        off_time = datetime.strptime(row['下班时间'], "%H:%M")
        sign_in = datetime.strptime(row['签到时间'], "%Y-%m-%d %H:%M:%S") if not pd.isna(row['签到时间']) else None
        sign_out = datetime.strptime(row['签退时间'], "%Y-%m-%d %H:%M:%S") if not pd.isna(row['签退时间']) else None

        overtime = 0

        # 计算加班时间
        if sign_out:
            if sign_out.time() > off_time.time():
                overtime += (sign_out - off_time).seconds / 3600.0

        if sign_in:
            if sign_in.time() < on_time.time():
                overtime += (on_time - sign_in).seconds / 3600.0

        # 计算餐补
        if "周末" in row['状态']:
            if overtime >= 3 and overtime < 8:
                meal_allowance += 15
                meal_allowance_records.append({
                    '工号': row['工号'],
                    '姓名': row['名字'],
                    '日期': row['出勤'],
                    '餐补金额': 15,
                    '原因': '周末加班 3-8 小时'
                })
            elif overtime >= 8:
                meal_allowance += 30
                meal_allowance_records.append({
                    '工号': row['工号'],
                    '姓名': row['名字'],
                    '日期': row['出勤'],
                    '餐补金额': 30,
                    '原因': '周末加班 8 小时及以上'
                })
        elif overtime >= 2.5 or (sign_out and sign_out.time().hour >= 21):
            meal_allowance += 15
            meal_allowance_records.append({
                '工号': row['工号'],
                '姓名': row['名字'],
                '日期': row['出勤'],
                '餐补金额': 15,
                '原因': '加班超过 21:00 或加班达到 2.5 小时'
            })

        # 计算迟到早退异常
        if "周末" not in row['状态']:
            if prev_sign_out and prev_sign_out.time() > datetime.strptime("21:00",
                                                                          "%H:%M").time() and sign_in and sign_in.time() <= datetime.strptime(
                    "09:30", "%H:%M").time():
                pass
            else:
                if sign_in and sign_in.time() > (on_time + timedelta(minutes=5)).time() and sign_in.time() < (on_time + timedelta(minutes=20)).time():
                    abnormal_count += 1
                    abnormal_records.append({
                        '工号': row['工号'],
                        '姓名': row['名字'],
                        '日期': row['出勤'],
                        '时间': sign_in.strftime("%H:%M:%S"),
                        '异常类型': '迟到'
                    })

            if sign_out and sign_out.time() > (off_time - timedelta(minutes=20)).time() and sign_out.time() < off_time.time():
                abnormal_count += 1
                abnormal_records.append({
                    '工号': row['工号'],
                    '姓名': row['名字'],
                    '日期': row['出勤'],
                    '时间': sign_out.strftime("%H:%M:%S"),
                    '异常类型': '早退'
                })

        overtime_total += overtime
        prev_sign_out = sign_out

    results.append({
        '部门': group['部门名称'].iloc[0],
        '工号': group['工号'].iloc[0],
        '姓名': group['名字'].iloc[0],
        '加班总时数': round(overtime_total, 2),
        '考勤异常数': abnormal_count,
        '餐补金额': meal_allowance
    })

# 按照工号分组处理数据
df.groupby('工号').apply(process_attendance)

# 输出结果到新的Excel文件
result_df = pd.DataFrame(results)
result_df.to_excel('./merge_summary.xlsx', index=False)

# 输出异常记录到Excel文件
abnormal_df = pd.DataFrame(abnormal_records)
abnormal_df.to_excel('./attendance_abnormal.xlsx', index=False)

# 输出餐补记录到Excel文件
meal_allowance_df = pd.DataFrame(meal_allowance_records)
meal_allowance_df.to_excel('./meal_allowance.xlsx', index=False)

print("考勤汇总已保存为 merge_summary.xlsx")
print("考勤异常记录已保存为 attendance_abnormal.xlsx")
print("餐补记录已保存为 meal_allowance.xlsx")
