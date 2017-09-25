"""
1 化疗及靶向治疗=>(治疗目的==一线治疗)=>本次治疗开始时间 != ''（1）最早
  内分泌治疗 => (治疗目的==一线治疗)=>开始用药日期 != ''（2）最早
  （1）（2）最早 [一线开始时间] $3
"""

import xlwt
import json
from functools import reduce
from datetime import datetime

wb = xlwt.Workbook()
ws_data = wb.add_sheet('PFS_new')
data_all = {}
row = 0
row_4s = 0


# 查看基本信息最全的 pid
def base_info_pid(store, sheet):
    pid_set = set()
    for line in data_all[store][sheet]:
        pid_set.add(float(line['pid']))
    return pid_set


# 多条分组再比较，一个dict
def one_parameter(ku, sheet, field, filter_name='True', filter_value='True', compared='Reverse'):

    key_value = {}
    # 将data_all 替换为多个可变的字符相连
    all_sheet_name = sheet.split(',')
    # 这样定义存在问题？
    all_sheet = []
    for li in all_sheet_name:
        all_sheet += data_all[ku][li]

    for line in all_sheet:
        # 过滤不需要的
        if line['pid'] == '' or line['pid'] == 'pid':
            continue
        if line[field] == '':
            continue
        if line[field] == '1970-01-01':
            continue

        if filter_name == 'True':
            line[filter_name] = 'True'
        if line[filter_name] == filter_value:
            pid = line['pid']
            if type(pid) == str:
                pid = float(pid)
            try:
                arr = key_value[pid]
            except:
                arr = []
                key_value[pid] = arr
            arr.append(line[field])
    return compared_list(key_value, compared)


# 比较分组后的大小，'Reverse'认为最小／早值
def compared_list(key_value, compared='Reverse'):
    # 进行比较选择最小的那个
    def compared_early(x, y):
        # 格式转换
        def parse_ymd(s):
            year_s, mon_s, day_s = s.split('-')
            return datetime(int(year_s), int(mon_s), int(day_s))

        def float_str(s):
            if type(s) != float:
                return s.strftime('%Y-%m-%d')
            else:
                return s
        # 比较
        if type(x) == str:
            x = parse_ymd(x)
            y = parse_ymd(y)

        if compared != 'Reverse':
            if x >= y:
                return float_str(x)
            else:
                return float_str(y)
        else:
            if x <= y:
                return float_str(x)
            else:
                return float_str(y)

    # 进行唯一值的确定
    def filter_f(s):
        if type(s) == float:
            return s
        else:
            return s and s.strip()

    last_date = {}
    for k, v in key_value.items():
        # 数组内的日期数值比较
        has_date = list(filter(filter_f, v))

        if len(has_date) == 0:
            # last_date[k] = ''
            continue
        elif len(has_date) == 1:
            last_date[k] = has_date[0]
        else:
            last_date[k] = reduce(compared_early, has_date)
    return last_date


# 2个dict分组, 若一个存在，另一个不存在则取存在的那个
def get_all_compared(first_one, second_one, store='307.xlsx', sheet='基本信息-基本信息'):
    # 将2个日期进行比较
    one_map = {}
    pid_set = base_info_pid(store, sheet)
    for pid in pid_set:
        arr = []
        try:
            first_one[pid]
            second_one[pid]
        # 有一个不存在或2个不存在
        except:
            try:
                first_one[pid]
            # not 1, look 2
            except:
                try:
                    second_one[pid]
                    # not 1, 2
                except:
                    continue
                else:
                    one_map[pid] = arr
                    if type(second_one[pid]) == list:
                        for lis in second_one[pid]:
                            arr.append(lis)
                    else:
                        arr.append(second_one[pid])
            # 1, in
            else:
                one_map[pid] = arr
                if type(first_one[pid]) == list:
                    for lis in first_one[pid]:
                        arr.append(lis)
                else:
                    arr.append(first_one[pid])
        else:
            one_map[pid] = arr
            if type(first_one[pid]) == list:
                for lis in first_one[pid]:
                    arr.append(lis)
            else:
                arr.append(first_one[pid])

            if type(second_one[pid]) == list:
                for lis in second_one[pid]:
                    arr.append(lis)
            else:
                arr.append(second_one[pid])
    return one_map


def write_data(one_start, two_start):
    global row
    if row == 0:
        ws_data.write(row, 0, 'pid')
        ws_data.write(row, 1, '一线开始时间')
        ws_data.write(row, 2, '二线开始时间')
        pid_set = base_info_pid('307.xlsx', '基本信息-基本信息')
    else:
        pid_set = base_info_pid('4s.xlsx', '基本信息_jibenxinxi')
    for pid in pid_set:
        row += 1
        ws_data.write(row, 0, pid)
        try:
            one_start[pid]
        except:
            ws_data.write(row, 1, '')
        else:
            ws_data.write(row, 1, one_start[pid])

        try:
            two_start[pid]
        except:
            ws_data.write(row, 2, '')
        else:
            ws_data.write(row, 2, two_start[pid])


# 获取307参数
def get_parameter_307():
    # 1 一线开始时间
    first_one = one_parameter('307.xlsx', '化疗及靶向治疗-化疗及靶向治疗', '本次治疗开始日期', '治疗目的', '一线治疗')
    second_one = one_parameter('307.xlsx', '内分泌治疗-内分泌治疗', '开始用药日期', '治疗目的', '一线治疗')
    one_start_group = get_all_compared(first_one, second_one)
    one_start = compared_list(one_start_group)
    # 2 二线开始时间
    first_two = one_parameter('307.xlsx', '化疗及靶向治疗-化疗及靶向治疗', '本次治疗开始日期', '治疗目的', '二线治疗')
    second_two = one_parameter('307.xlsx', '内分泌治疗-内分泌治疗', '开始用药日期', '治疗目的', ' 二线治疗')
    two_start_group = get_all_compared(first_two, second_two)
    two_start = compared_list(two_start_group)

    write_data(one_start, two_start)


def get_paramter_4s():
    # 一线开始时间
    one_time = one_parameter('4s.xlsx', '复发转移_recum_BC', '用药开始日期', '治疗目的', '一线治疗')
    # 二线开始时间
    two_time = one_parameter('4s.xlsx', '复发转移_recum_BC', '用药开始日期', '治疗目的', '二线治疗')
    write_data(one_time, two_time)

# 代码开始的地方
if __name__ == '__main__':
    try:
        f = open('data_all', 'r')
        data_all = json.load(f)
        get_parameter_307()
        get_paramter_4s()
        wb.save('data_out/一线PFS（新增）.xls')
        f.close()
        print('over，一线PFS（新增）')
    except Exception as e:
        print(e)