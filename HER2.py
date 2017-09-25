import json
import xlwt

wb = xlwt.Workbook()
ws_307 = wb.add_sheet('HRE2')
# ws_4s = wb.add_sheet('HER2_4s')
data_all = {}
row = 0
row_4s = 0


def group(ku, sheet, filter_name, filter_value, flag=0):
    key_value = {}

    for line in data_all[ku][sheet]:
        if line[filter_name] in filter_value:
            try:
                arr = key_value[line['pid']]
            except:
                arr = []
                key_value[line['pid']] = arr
            arr.append(line[filter_name])
    # 计算次数
    count_map = {}
    same_map = {}
    for k, v in key_value.items():
        same = set()
        count_map[k] = len(v)
        for li in v:
            if flag:
                if li in ('+',  '-', '±', 1, 2, 3):
                    li = '---'
            same.add(li)
        # 判断是否为单值
        if len(v) == 1:
            continue
        elif len(same) == 1:
            same_map[k] = '一致'
        elif len(same) == 2:
            same_map[k] = '不一致'
        else:
            print('长度有问题')
    return count_map, same_map


def get_307():
    count_map1, same_map1 = group('307.xlsx', '分子病理检测-分子病理检测', '原位杂交测结果', ('扩增阳性', '扩增阴性'))
    count_map2, same_map2 = group('307.xlsx', '分子病理检测-分子病理检测', 'HER-2免疫组化结果', ('+++', '+',  '-',  '±'), flag=1)
    write_data(count_map1, same_map1, count_map2, same_map2)


def get_4s():
    count_map1, same_map1 = group('4s.xlsx', '肿瘤标记检测_biom_BC', 'HER2FISH检测', (2, 3))
    count_map2, same_map2 = group('4s.xlsx', '肿瘤标记检测_biom_BC', 'HER2免疫组化结果', (1, 2, 3, 5), flag=1)
    write_data(count_map1, same_map1, count_map2, same_map2)


# 查看基本信息最全的 pid
def base_info_pid(store, sheet):
    pid_set = set()
    for line in data_all[store][sheet]:
        pid_set.add(line['pid'])
    return pid_set


def write_data(count_map1, same_map1, count_map2, same_map2):
    global row
    if row == 0:
        ws_307.write(row, 0, 'HER2（FISH）次数')
        ws_307.write(row, 1, 'HER2（FISH）次数')
        ws_307.write(row, 2, 'HER2（FISH）一致性')
        ws_307.write(row, 3, 'HER2（免疫组化）次数')
        ws_307.write(row, 4, 'HER2（免疫组化）一致性')
        pid_set = base_info_pid('307.xlsx', '基本信息-基本信息')
    else:
        print(row)
        pid_set = base_info_pid('4s.xlsx', '基本信息_jibenxinxi')
    for pid in pid_set:
        row += 1
        ws_307.write(row, 0, pid)
        try:
            count_map1[pid]
        except:
            ws_307.write(row, 1, '')
        else:
            ws_307.write(row, 1, count_map1[pid])

        try:
            same_map1[pid]
        except:
            ws_307.write(row, 2, '')
        else:
            ws_307.write(row, 2, same_map1[pid])

        try:
            count_map2[pid]
        except:
            ws_307.write(row, 3, '')
        else:
            ws_307.write(row, 3, count_map2[pid])

        try:
            same_map2[pid]
        except:
            ws_307.write(row, 4, '')
        else:
            ws_307.write(row, 4, same_map2[pid])


# def write_data_4s(count_map1, same_map1, count_map2, same_map2):
#     global row_4s
#     row_4s.write(row_4s, 0, 'HER2（FISH）次数')
#     row_4s.write(row_4s, 1, 'HER2（FISH）次数')
#     row_4s.write(row_4s, 2, 'HER2（FISH）一致性')
#     row_4s.write(row_4s, 3, 'HER2（免疫组化）次数')
#     row_4s.write(row_4s, 4, 'HER2（免疫组化）一致性')
#     pid_set = base_info_pid('307.xlsx', '基本信息-基本信息')
#     for pid in pid_set:
#         row_4s += 1
#         row_4s.write(row_4s, 0, pid)
#
#         try:
#             count_map1[pid]
#         except:
#             row_4s.write(row_4s, 1, '')
#         else:
#             row_4s.write(row_4s, 1, count_map1[pid])
#
#         try:
#             same_map1[pid]
#         except:
#             row_4s.write(row_4s, 2, '')
#         else:
#             row_4s.write(row_4s, 2, same_map1[pid])
#
#         try:
#             count_map2[pid]
#         except:
#             ws_307.write(row_4s, 3, '')
#         else:
#             ws_307.write(row_4s, 3, count_map2[pid])
#
#         try:
#             same_map2[pid]
#         except:
#             ws_307.write(row_4s, 4, '')
#         else:
#             ws_307.write(row_4s, 4, same_map2[pid])

if __name__ == '__main__':
    f = open('data_all', 'r')
    data_all = json.load(f)
    get_307()
    get_4s()
    wb.save('data_out/HER2.xls')
    f.close()
