# 依托钉钉导出考勤打卡excel表格，实现考勤统计和可视化
# 这个脚本统计一段时间内的打卡情况，全自动
#
# 1. 统计总的工作时长，分两个类别：a.不包括休息日 和 b.包括休息日
# 2. 统计 c.工作日平均工作时长
# 3. d.统计工作日出勤天数 和 e.应出勤天数
# 4. 统计工作日，每个人，f.每个打卡节点（例如上午上班，上午下班，下午上班等，一天6个打卡节点）的平均打卡时间
# 以上结果可视化

# # # matplotlib 画图, 注意 linux 下可能不包含对于字体，需要手动装
# # # 字体问题在win环境下实测可行
###############################################
###############################################

import pandas as pd
import os
from datetime import datetime, timedelta
import matplotlib.pyplot as plt

plt.rcParams['font.sans-serif'] = ['SimHei']


def read_total_worktime(file_path):
    # return
    # 出勤天数 yes
    # 工作总时间 yes
    # 应出勤天数
    # 总天数

    df = pd.read_excel(file_path, sheet_name='月度汇总')
    row = df.shape[0]
    column = df.shape[1]
    for i in range(3, row):  # 行3开始，每行一人
        name = df.iloc[i][0]
        attendance_days = pd.to_numeric(df.iloc[i][6], errors='coerce')
        total_worktime = pd.to_numeric(df.iloc[i][8], errors='coerce')
        # print(name)

        if name in data_dict:
            if not pd.isna(attendance_days):
                # print(attendance_days)
                data_dict[name]['出勤天数'] += attendance_days
            if not pd.isna(total_worktime):
                # print(total_worktime)
                data_dict[name]['工作总时间'] += total_worktime
        else:
            print(f'ERROR: Name {name} not found!')

    for j in range(32, column):  # 列32开始记录每天的情况
        if not pd.isna(df.iloc[3, j]):  # 是一天
            total_days[0] += 1
            if not df.iloc[3, j] in ['休息', '休息并打卡']:  # 是工作日
                work_days[0] += 1


def read_workday_worktime(file_path):
    # return
    # 工作日早上打卡平均时间
    # 工作日日平均工作时间
    # !!! 上述二者返回的是总时间，然后计算平均时间 !!!
    df = pd.read_excel(file_path, sheet_name='每日统计')
    row = df.shape[0]
    column = df.shape[1]
    for i in range(3, row):  # 行3开始，每行一人
        name = df.iloc[i][0]
        shift = df.iloc[i][8]
        morning_checktime = df.iloc[i][9]
        daily_worktime = pd.to_numeric(df.iloc[i][24], errors='coerce')
        # print(name)
        # print(shift)
        # print(morning_checktime)
        # print(daily_worktime)
        if shift != '休息':
            if not pd.isna(morning_checktime):
                time_obj = datetime.strptime(morning_checktime, '%H:%M')
                start_time = datetime.strptime('05:00', '%H:%M')
                end_time = datetime.strptime('10:30', '%H:%M')
                if start_time <= time_obj <= end_time:
                    data_dict[name]['工作日早出勤打卡平均时间'] += time_obj - datetime.strptime('00:00', '%H:%M')
                    data_dict[name]['早打卡天数'] += 1

            if not pd.isna(daily_worktime):
                data_dict[name]['工作日工作总时间'] += daily_worktime
        # else:
        #     print('今天休息')


def read_excels(excel_dir):
    excel_files = os.listdir(excel_dir)
    for excel_path in excel_files:
        read_one_excel(os.path.join(excel_dir, excel_path))


def read_one_excel(file_path):
    read_total_worktime(file_path)
    read_workday_worktime(file_path)


def datetime_to_str(datetime_obj):
    return datetime_obj.strftime('%Y-%m-%d')


def visualize(datadict):
    data = datadict
    attributes = ['出勤天数', '工作总时间', '工作日工作总时间', '工作日日平均工作时间', '工作日早出勤打卡平均时间',
                  '早打卡天数']

    # 每个属性一张表
    for i, attribute in enumerate(attributes):
        # 提取当前属性的值
        attribute_data = {person: data[person][attribute] for person in data}
        sorted_data = sorted(attribute_data.items(), key=lambda x: x[1])
        # print(sorted_data)
        keys = [item[0] for item in sorted_data]    # keys is name in order
        attribute_values = [item[1] for item in sorted_data]

        plt.figure(figsize=(10, 6))

        # 天数
        if attribute in ['出勤天数', '早打卡天数']:
            plt.bar(keys, attribute_values, color='skyblue')
            plt.title(schedule + ' ' + attribute + f'(工作日天数{work_days[0]})')
            for index, value in enumerate(attribute_values):
                plt.text(keys[index], value, str(value), ha='center', va='bottom')

        # 时间
        elif attribute in ['工作总时间', '工作日工作总时间', '工作日日平均工作时间']:
            plt.bar(keys, attribute_values, color='skyblue')
            plt.title(schedule + ' ' + attribute + '(小时)')
            for index, value in enumerate(attribute_values):
                plt.text(keys[index], value, f'{value:.1f}', ha='center', va='bottom')
        # 时刻
        elif attribute in ['工作日早出勤打卡平均时间']:
            for index in range(len(attribute_values)):  # timedelta -> str
                delta = attribute_values[index]
                hours = delta.seconds // 3600
                minutes = (delta.seconds % 3600) // 60
                attribute_values[index] = "{:02d}:{:02d}".format(hours, minutes)

            plt.bar(keys, attribute_values, color='skyblue')
            for index, value in enumerate(attribute_values):
                plt.text(keys[index], value, str(value), ha='center', va='bottom')
            plt.title(schedule + ' ' + attribute)
        median_value = attribute_values[int(len(attribute_values) / 2)]
        if attribute in ['工作日早出勤打卡平均时间', '出勤天数', '早打卡天数']:
            median_str = f'中位数={median_value}'
        else:
            # in ['工作总时间', '工作日工作总时间', '工作日日平均工作时间']:
            median_str = f'中位数={median_value:.1f}'

        plt.text(keys[0], median_value, median_str, ha='center', va='bottom', color='red')
        plt.axhline(y=median_value, color='red', linestyle='--', label='Median')
        plt.xticks(rotation=45)
        plt.tight_layout()

        # 保存图形到文件
        plt.savefig(f'{output_dir}/{attribute}.png')

        # 关闭当前图形窗口
        plt.close()


if __name__ == '__main__':
    # 设置参数
    excel_dir = './3.10-4.17'
    schedule = '2024.3.10-2024.4.17'

    output_dir = schedule
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    data_dict = {}
    work_days = [0]  # 工作日天数
    total_days = [0]  # 统计天数

    # load people's name
    with open('person.txt', 'r', encoding='utf-8') as file:
        names = [line.strip() for line in file.readlines()]

    for name in names:
        data_dict[name] = {'出勤天数': 0,
                           '工作总时间': 0.0,
                           '工作日工作总时间': 0.0,
                           '工作日日平均工作时间': 0.0,
                           '工作日早出勤打卡平均时间': datetime.strptime('00:00', '%H:%M'),
                           '早打卡天数': 0}

    read_excels(excel_dir)

    # 打印dict
    for name, attributes in data_dict.items():
        data_dict[name]['工作总时间'] /= 60
        data_dict[name]['工作日工作总时间'] /= 60
        data_dict[name]['工作日日平均工作时间'] = data_dict[name]['工作日工作总时间'] / work_days[0]
        data_dict[name]['工作日早出勤打卡平均时间'] = data_dict[name]['工作日早出勤打卡平均时间'] - datetime.strptime('00:00', '%H:%M')
        data_dict[name]['工作日早出勤打卡平均时间'] /= data_dict[name]['早打卡天数'] + 0.000001

        # print dict
        # print(f"{name}:")
        # for attribute, value in attributes.items():
        #     print(f"    {attribute}: {value}")

    print('\n')
    print(f'工作日天数: {work_days[0]}')
    print(f'统计天数：{total_days[0]}')

    # pop someone, do it to ignore someone
    data_dict.pop('a')
    data_dict.pop('b')

    visualize(data_dict)






