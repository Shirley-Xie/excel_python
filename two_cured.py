"""
1 化疗及靶向治疗=>(治疗目的==二线治疗 and
  (药物通用名-商品名X or 化疗方案名称 )==(曲妥珠单抗or 赫赛汀or H )

  内分泌治疗 => (治疗目的==二线治疗 and
  药物通用名（其他）== “赫赛汀” or “曲妥珠单抗”
  有一个就是‘二线赫赛汀治疗’

2 化疗及靶向治疗=>(治疗目的==二线治疗 and
  药物通用名-商品名==（拉帕替尼”或“L”或“lapatinib”）and（“卡培他滨”或“X”），且非（“赫赛汀”或“曲妥珠单抗”或H）

  内分泌治疗 => (治疗目的==二线治疗 and
  药物通用名-商品名==“拉帕替尼”或“L”或“lapatinib”）且（“卡培他滨”或“X”），且（“赫赛汀”或“曲妥珠单抗”或H）
  有一个就是‘二线拉帕替尼且卡培他滨治疗
"""

import xlwt
import json
import re

wb = xlwt.Workbook()
ws_data = wb.add_sheet('二线治疗_赫赛汀_帕替尼且卡培他滨')
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
def yes_or_not(ku, sheet, names, cond_re):

    hst_set = set()

    for line in data_all[ku][sheet]:
        # 过滤不需要的
        if line['pid'] == '' or line['pid'] == 'pid':
            continue
        if line['治疗目的'] == '二线治疗':
            for k, v in line.items():
                if k in names:
                    if re.findall(cond_re, v):
                        hst_set.add(line['pid'])
                        break
    return hst_set


def write_data(one_, two_):
    global row
    if row == 0:
        ws_data.write(row, 0, 'pid')
        ws_data.write(row, 1, '二线赫赛汀治疗')
        ws_data.write(row, 2, '二线拉帕替尼且卡培他滨治疗')
        pid_set = base_info_pid('307.xlsx', '基本信息-基本信息')
    else:
        pid_set = base_info_pid('4s.xlsx', '基本信息_jibenxinxi')
    for pid in pid_set:
        row += 1
        ws_data.write(row, 0, pid)
        if pid not in one_:
            ws_data.write(row, 1, '')
        else:
            ws_data.write(row, 1, '是')

        if pid not in two_:
            ws_data.write(row, 2, '')
        else:
            ws_data.write(row, 2, '是')


# 获取307参数
def get_parameter():
    goods_names = ('药物通用名-商品名', '药物通用名-商品名-2', '药物通用名-商品名-3', '药物通用名-商品名-4', '药物通用名-商品名-5',
                   '药物通用名-商品名-6', '药物通用名-商品名-7', '药物通用名-商品名-8', '化疗方案名称')
    names_another = ('药物通用名-其他', '药物通用名-其他-2', '药物通用名-其他-3', '药物通用名-其他-4')
    re_1 = r"(曲妥珠单抗|赫赛汀|H)"
    re_2 = r"(拉帕替尼|L|lapatinib)(卡培他滨|X)(^赫赛汀|^曲妥珠单抗|^H|\w*)"
    re_3 = r"(拉帕替尼|L|lapatinib)(卡培他滨|X)(赫赛汀|曲妥珠单抗|H)"
    chemothy_names = goods_names + ('化疗方案名称',)

    # 1 二线赫赛汀治疗
    yes_hst1 = yes_or_not('307.xlsx', '化疗及靶向治疗-化疗及靶向治疗', chemothy_names, re_1)
    yes_hst2 = yes_or_not('307.xlsx', '内分泌治疗-内分泌治疗', names_another, r"(赫赛汀|曲妥珠单抗)")
    # 求并集
    two_hst = yes_hst1 | yes_hst2

    # 2 二线拉帕替尼且卡培他滨治疗
    yes_lptn1 = yes_or_not('307.xlsx', '化疗及靶向治疗-化疗及靶向治疗', goods_names, re_2)
    yes_lptn2 = yes_or_not('307.xlsx', '内分泌治疗-内分泌治疗', goods_names, re_3)
    # 求并集
    two_lptn = yes_lptn1 | yes_lptn2
    write_data(two_hst, two_lptn)

    # 4s
    # 二线赫赛汀治疗
    yes_lptn1_4s = yes_or_not('4s.xlsx', '复发转移_recum_BC', '治疗方案', re_1)
    # 二线拉帕替尼且卡培他滨治疗
    yes_lptn2_4s = yes_or_not('4s.xlsx', '复发转移_recum_BC', '治疗方案', re_2)
    write_data(yes_lptn1_4s, yes_lptn2_4s)

# 代码开始的地方
if __name__ == '__main__':
    try:
        f = open('data_all', 'r')
        data_all = json.load(f)
        get_parameter()
        wb.save('data_out/二线治疗（新增）.xls')
        f.close()
        print('over')
    except Exception as e:
        print(e)