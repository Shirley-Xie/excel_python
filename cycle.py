# -*- coding: utf-8 -*-
import os
import xlrd
import json
import xlwt
import copy
import re

"""
参数17：返回 PID, 术前化疗总周期数

计算获取：术前化疗总周期数
判断逻辑
    307，表单：化疗及靶向治疗，
    '治疗目的'== 新辅助治疗，and '化疗方案名称'！= 空，
        若一个PID有多次记录，直接相加
            若'该方案实际周期数值'=空则取'化疗方案名称'中的数字相加（除方案IP-16外：为0）

    4S，表单：术前化疗_pre0-chem，总周期数
        若一个PID有多次记录，序号为1的总周期数相加  

"""

data_all = {}
wb = xlwt.Workbook()
ws_data = wb.add_sheet('data')
row = 0


# 术前化疗总周期数的获取
def get_cycle():
    cycle_map_307 = {}
    try:
        # 307库，表单：化疗及靶向治疗
        for line in data_all['307.xlsx']['化疗及靶向治疗-化疗及靶向治疗']:
            # '治疗目的'== 新辅助治疗，and '化疗方案名称'！= 空，
            if line['治疗目的'] == '新辅助治疗' and line['化疗方案名称'] != '':
                # 对同一PID进行分组
                date = line['本次治疗开始日期']
                pid = int(line['pid'])
                try:
                    arr = cycle_map_307[pid]
                except:
                    arr = []
                    cycle_map_307[pid] = arr
                # 有值则直接取值否则看'化疗方案名称'
                if line['该方案实际周期数'] != '':
                    arr.append(int(line['该方案实际周期数']))
                else:
                    # 进行正则匹配找到数字相加
                    str = line['化疗方案名称']
                    # 判断是否有VP-16,有为0
                    sum_all = 0
                    if re.findall(r"VP-16|vp-16|vp16|VP16", str):
                        sum_all = 0
                    else:
                        num_li = re.findall(r"\d+", str)
                        sum_all = sum(map(lambda a: int(a), num_li))
                    arr.append(sum_all)

        # cycle_map的周期数相加
        last_map_307 = {}
        for k, v in cycle_map_307.items():
            last_map_307[k] = sum(v)

        # write_data(last_map_307, '术前化疗总周期数')

        # 4S库：表单：术前化疗_preo_chem，字段：'序号'为1'总周期数'
        cycle_map_4s = {}
        for line in (data_all['4s.xlsx']['术前化疗_preo_chem']):
            # 同一PID，序号为1的总周期数相加
            if line['序号'] == 1:
                pid = int(line['pid'])
                try:
                    arr = cycle_map_4s[pid]
                except:
                    arr = []
                    cycle_map_4s[pid] = arr
                if line['总周期数'] == '':
                    line['总周期数'] = 0
                arr.append(int(line['总周期数']))

        last_map_4s = {}
        for k, v in cycle_map_4s.items():
            last_map_4s[k] = sum(v)

        # 合并2个dict
        last_map_4s.update(last_map_307)
        write_data(last_map_4s, '术前化疗总周期数')

    except Exception as e:
        print(e)


def write_data(data_list, name):
    global row
    ws_data.write(row, 0, 'pid')
    ws_data.write(row, 1, name)
    for k, v in data_list.items():
        row += 1
        ws_data.write(row, 0, str(k))
        ws_data.write(row, 1, v)


if __name__ == '__main__':
    try:
        f = open('data_all', 'r')
        data_all = json.load(f)
        get_cycle()
        wb.save('data_out/cycle_all.xls')
        f.close()
    except Exception as e:
        print(e)