import math
import json
import xlwt
import re
from functools import reduce
from datetime import datetime

"""
307：
1 化疗及靶向治疗／内分泌治疗 => 治疗目的==二线治疗（是否PFS？是：‘’）
2 疗效评估 => 评估前状态	!= “新辅助治疗”、“辅助治疗”、空，(是否PFS？是：‘’）
是
3 化疗及靶向治疗 => 治疗目的==一线治疗 => 本次治疗开始时间!=1970-01-01（1）
             => 治疗目的==二线治疗 => 本次治疗开始时间!=1970-01-01（2）
             （PFS：（2）-（1））
3空则4              
4 内分泌治疗 => 治疗目的==一线治疗 => 开始用药日期!=1970-01-01	(1)
          => 治疗目的==二线治疗 => 开始用药日期!=1970-01-01（2）
          （PFS：（2）-（1））
否          
5 化疗及靶向治疗 => 治疗目的==一线治疗 => 本次治疗开始时间!=1970-01-01（1）
             最后一次随访日期!=1970-01-01（2)
             （PFS：（2）-（1））
5空则6              
6 内分泌治疗 => 治疗目的==一线治疗 => 开始用药日期!=1970-01-01 (1)
            最后一次随访日期!=1970-01-01（2）
          （PFS：（2）-（1））

4s:
1复发转移 => 字段名字有‘进展’|治疗开始日期!=1970-01-01|治疗结束日期!=1970-01-01(是否PFS？是：‘’)
是：
2复发转移 => 靶向治疗开始日期 ！= 1970-01-01（1）
        => 治疗开始日期 ！= 1970-01-01（2）
        （PFS：（2）-（1）
否        
3复发转移 => 靶向治疗开始日期 ！= 1970-01-01（1）
            => 最后一次随访日期 ！= 1970-01-01（2）
            （PFS：（2）-（1）
"""

wb = xlwt.Workbook()
ws_data = wb.add_sheet('PFS')
data_all = {}
row = 0


def group(ku, sheet, field, filter_name='True', filter_value='True', flags=0):
    key_value = {}
    for line in data_all[ku][sheet]:
        if filter_name == 'True':
            line[filter_name] = 'True'
        if line[filter_name] == filter_value:
            #  将空值和1970舍弃
            if line[field] == '' or line[field] == '1970-01-01':
                continue
            try:
                arr = key_value[line['pid']]
            except:
                arr = []
                key_value[line['pid']] = arr
            arr.append(line[field])
    return compared_list(key_value, flags)


def compared_list(key_value, flags):
    # 格式转换
    def parse_ymd(s):
        year_s, mon_s, day_s = s.split('-')
        return datetime(int(year_s), int(mon_s), int(day_s))

    # 进行比较选择最小的那个
    def compared_early(x, y):
        # 比较
        if type(x) == str:
            x = parse_ymd(x)
        if type(y) == str:
            y = parse_ymd(y)
        if flags:
            if x <= y:
                return x
            else:
                return y
        else:
            if x >= y:
                return x
            else:
                return y

    # 进行唯一值的确定
    last_date = {}
    for k, v in key_value.items():
        # 数组内的日期数值比较
        if len(v) == 1:
            last_date[k] = parse_ymd(v[0])
        elif len(v) > 1:
            last_date[k] = reduce(compared_early, v)
        else:
            print('日期长度为0')
    return last_date


def during(is_pfs_307, start_map, end_map, whether='yes'):
    during_time = {}

    def yes_or_not():
        if whether == 'yes':
            return k in is_pfs_307
        else:
            return k not in is_pfs_307

    for k, start_v in start_map.items():
        # 是的情况下，且开始和结束值都存在
        if yes_or_not():
            if k in end_map.keys():
                end_v = end_map[k]
                during_days = (end_v - start_v).days
                if during_days < 0:
                    print(k)
                during_ = str(during_days // 30) + '.' + str(during_days % 30)
                during_time[k] = during_

    return during_time


def get_pfs_307():
    # 获取是否为pfs
    is_pfs_one = set()
    is_pfs_two = set()
    all_pfs_one = set()
    for line in data_all['307.xlsx']['化疗及靶向治疗-化疗及靶向治疗'] + data_all['307.xlsx']['内分泌治疗-内分泌治疗']:
        # 判断是否为PFS
        all_pfs_one.add(line['pid'])
        if line['治疗目的'] == '二线治疗':
            is_pfs_one.add(line['pid'])

    for line in data_all['307.xlsx']['疗效评价-疗效评价']:
        all_pfs_one.add(line['pid'])
        if line['评估前状态'] != '新辅助治疗' and line['评估前状态'] != '新辅助治疗' and line['评估前状态'] != '':
            is_pfs_two.add(line['pid'])

    is_pfs_307 = is_pfs_one | is_pfs_two
    not_pfs_307 = all_pfs_one - is_pfs_307

    # 查找需要的数值
    start_chemotherapy_map = group('307.xlsx', '化疗及靶向治疗-化疗及靶向治疗', '本次治疗开始日期', '治疗目的', '一线治疗')
    end_chemotherapy_map = group('307.xlsx', '化疗及靶向治疗-化疗及靶向治疗', '本次治疗开始日期', '治疗目的', '二线治疗')
    start_endocrine_map = group('307.xlsx', '内分泌治疗-内分泌治疗', '开始用药日期', '治疗目的', '一线治疗')
    end_endocrine_map = group('307.xlsx', '内分泌治疗-内分泌治疗', '开始用药日期', '治疗目的', '二线治疗')
    visit_map = group('307.xlsx', '随访-复发及生存随访', '随访日期', flags='not_Reverse')

    # 若是pfs
    # 找到第一个条件成立的所有pid
    during_time_first = during(is_pfs_307, start_chemotherapy_map, end_chemotherapy_map)
    not_during_time = is_pfs_307 - during_time_first.keys()
    # 第一个条件不成立 找条件二成立的条件
    during_time_second = during(not_during_time, start_endocrine_map, end_endocrine_map)
    # 还有都不存在的写入时为空
    # 合并是pfs所有
    during_time_first.update(during_time_second)

    # 若不是pfs 同上
    during_time_f = during(not_pfs_307, start_chemotherapy_map, visit_map)
    not_during = not_pfs_307 - during_time_f.keys()
    during_time_s = during(not_during, start_endocrine_map, visit_map)
    # 合并是非pfs所有
    during_time_f.update(during_time_s)
    # 将所有有时间的合并
    # during_time_first.update(during_time_f)
    return is_pfs_307, not_pfs_307, during_time_first, during_time_f


def get_pfs_4s():
    def condition(li, name):
        return li[name] != '1970-01-01' and li[name] != ''

    is_pfs_4s = set()
    not_pfs_4s = set()
    #  判断是否为pfs
    fields = ['进展后选择', '治疗间疾病是否进展[其他]', '疾病进展时间[符号]', '进展后选择[其他]', '进展日期', '治疗间疾病是否进展', '疾病进展时间']
    for line in data_all['4s.xlsx']['复发转移_recum_BC']:
        for field in fields:
            if condition(line, field) or condition(line, '治疗开始日期') or condition(line, '治疗结束日期'):
                is_pfs_4s.add(line['pid'])
            else:
                not_pfs_4s.add(line['pid'])

    start_change_map = group('4s.xlsx', '复发转移_recum_BC', '靶向治疗开始日期')
    end_change_map = group('4s.xlsx', '复发转移_recum_BC', '治疗开始日期')
    visit_map = group('4s.xlsx', '随访信息1_foup_BC', '末次随访日期', flags='not_Reverse')
    # 判断条件1成立，否则看条件2
    change_first = during(is_pfs_4s, start_change_map, end_change_map)

    change_second = during(is_pfs_4s, start_change_map, visit_map, whether='no')
    # change_first.update(change_second)
    return is_pfs_4s, not_pfs_4s, change_first, change_second


def write_data(is_pfs_307, not_pfs_307, pfs_307_first, pfs_307_second):
    global row
    ws_data.write(row, 0, 'pid')
    ws_data.write(row, 1, '是否为PFS')
    ws_data.write(row, 2, 'PFS')
    for pid in is_pfs_307:
        row += 1
        ws_data.write(row, 0, pid)
        ws_data.write(row, 1, '是')
        try:
            pfs_307_first[pid]
        except:
            ws_data.write(row, 2, '')
        else:
            ws_data.write(row, 2, pfs_307_first[pid])

    for pid in not_pfs_307:
        row += 1
        ws_data.write(row, 0, pid)
        ws_data.write(row, 1, '不是')
        try:
            pfs_307_second[pid]
        except:
            ws_data.write(row, 2, '')
        else:
            ws_data.write(row, 2, pfs_307_second[pid])


# 代码开始的地方
if __name__ == '__main__':
    f = open('data_all', 'r')
    data_all = json.load(f)
    is_307, pfs_307, during_f_307, during_s_307 = get_pfs_307()
    is_4s, pfs_4s, during_f_4s, during_s_4s = get_pfs_4s()
    is_307.update(is_4s)
    pfs_307.update(pfs_4s)
    during_f_307.update(during_f_4s)
    during_s_307.update(during_s_4s)
    write_data(is_307, pfs_307, during_f_307, during_s_307)
    wb.save('data_out/PFS.xls')
    f.close()
