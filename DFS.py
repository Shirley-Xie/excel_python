import os
import xlrd
import xlwt
import copy
import json
from functools import reduce
from datetime import datetime
"""
将内分泌治疗目的 改为 治疗目的
参数16，DFS
307：
1 疗效评价-疗效评价 =>(评估前状态==辅助治疗)=>评估日期(复发转移时间)

2 化疗及靶向治疗-化疗及靶向治疗=>(治疗目的==一线治疗)=>本次治疗开始日期()
内分泌治疗=>(内分泌治疗目的==一线治疗)=>开始用药日期	()
选择最早的时间（一线开始时间）

3 手术治疗-手术治疗=>手术日期

4 化疗及靶向治疗-化疗及靶向治疗=>(治疗目的==新辅助治疗)=>本次治疗开始日期
内分泌治疗-内分泌治疗=>(内分泌治疗目的==新辅助治疗)=>开始用药日期	
选择最早的时间（新辅助开始时间）

5 基本信息-基本信息=>首次确诊日期

6 随访-复发及生存随访=>随访日期	

7 化疗及靶向治疗-化疗及靶向治疗 =>（治疗目的==辅助治疗）=>本次治疗开始日期(1)最大
  内分泌治疗-内分泌治疗=>(治疗目的==辅助治疗)=>开始用药日期(2)最大
  手术治疗-手术治疗=>手术日期 （3） 最大
  (1) (2)最大（最后一次治疗时间）

4s:
1 复发转移_recum_BC => 复发转移日期（复发转移时间）
  5        => DFS（DFS）
2 手术治疗_surgry => 手术日期	（手术时间）

3 术前化疗_pre0-chem => 治疗开始日期
术前内分泌治疗_preo_ecthero => 治疗开始日期
术前靶向治疗_preo_targ => 治疗开始日期
选择最早的时间（新辅助开始时间)

4 随访信息_foup_BC => 末次随访日期（最后一次随访日期）

6  取最大的日期
手术治疗_surgry => 手术日期 （1）
 辅助化疗_adchem_BC => 化疗开始日期（索引）（2）
 辅助靶向治疗_adjutar_BC => 靶向治疗开始日期（索引）(3)
 辅助内分泌治疗_adjuec_BC => TAM类开始日期 (4)
 ... 					   AI类开始日期 (5)
 1-5（最后一次治疗时间）
"""

wb = xlwt.Workbook()
ws_307 = wb.add_sheet('DFS_307')
ws_4s = wb.add_sheet('DFS_4s')
data_all = {}
row = 0
row_4s = 0


# 查看基本信息最全的 pid
def base_info_pid(store, sheet):
    pid_set = set()
    for line in data_all[store][sheet]:
        pid_set.add(float(line['pid']))
    return pid_set


# 转化为json格式方便读取
def init_data():
    global data_all
    # data_all = {}
    # 录入文件，输出文件
    path_dir = './data'
    out_dir = './data_out'
    for filename in os.walk(path_dir).__next__()[2]:
        data_all[filename] = {}
        path = path_dir + '/' + filename
        workbook = xlrd.open_workbook(path)
        for sheet_name in workbook.sheet_names():
            data_all[filename][sheet_name] = []
            sheet = workbook.sheet_by_name(sheet_name)
            nrows = sheet.nrows
            ncols = sheet.ncols
            headers = {}
            # if filename == '4s.xlsx' and sheet_name == '复发转移_recum_BC':
            if True:
                for col_index in range(ncols):
                    # value = sheet.cell_value(0, col_index).encode('utf-8')
                    value = sheet.cell_value(0, col_index)
                    headers[col_index] = value
                for row_index in range(1, nrows):
                    line = {}
                    for col_index in range(ncols):
                        # value = sheet.cell_value(row_index, col_index).encode('utf-8')
                        try:
                            value = sheet.cell_value(row_index, col_index)
                        except Exception as e:
                            print('filename:', filename, ' sheet_name:', sheet_name, ' row_index:',row_index, ' col_index:', col_index)
                        header = headers[col_index]
                        line[header] = value
                    data_all[filename][sheet_name].append(line)
        out_file_name = out_dir + '/' + filename
        f = open(out_file_name, 'w')
        f.write(json.dumps(data_all[filename]))
        f.close()

    f = open('data_all', 'w')
    f.write(json.dumps(data_all))
    f.close()


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
        if line[field] != '' and line[field] == '1970-01-01':
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


def write_307(change_time, operate_map, new_start, visit_time, one_start, diagnosis_time, cured_map):
    global row
    # data_all = DATA_ALL
    pid_set = base_info_pid('307.xlsx', '基本信息-基本信息')
    ws_307.write(row, 0, 'pid')
    ws_307.write(row, 1, '复发转移时间')
    ws_307.write(row, 2, '手术时间')
    ws_307.write(row, 3, '新辅助开始时间')
    ws_307.write(row, 4, '一线治疗开始时间')
    ws_307.write(row, 5, '诊断时间')
    ws_307.write(row, 6, '最后一次随访日期')
    ws_307.write(row, 7, '最后一次治疗时间')
    for pid in pid_set:
        row += 1
        ws_307.write(row, 0, pid)

        try:
            change_time[pid]
        except:
            ws_307.write(row, 1, '')
        else:
            ws_307.write(row, 1, change_time[pid])

        try:
            operate_map[pid]
        except:
            ws_307.write(row, 2, '')
        else:
            ws_307.write(row, 2, operate_map[pid])

        try:
            new_start[pid]
        except:
            ws_307.write(row, 3, '')
        else:
            ws_307.write(row, 3, new_start[pid])

        try:
            one_start[pid]
        except:
            ws_307.write(row, 4, '')
        else:
            ws_307.write(row, 4, one_start[pid])

        try:
            diagnosis_time[pid]
        except:
            ws_307.write(row, 5, '')
        else:
            ws_307.write(row, 5, diagnosis_time[pid])

        try:
            visit_time[pid]
        except:
            ws_307.write(row, 6, '')
        else:
            ws_307.write(row, 6, visit_time[pid])

        try:
            cured_map[pid]
        except:
            ws_307.write(row, 7, '')
        else:
            ws_307.write(row, 7, cured_map[pid])


def write_4s(change_time, operate_map, new_start, visit_time, dfs, cured_time):
    global row_4s
    # data_all = DATA_ALL
    ws_4s.write(row_4s, 0, 'pid')
    ws_4s.write(row_4s, 1, '复发转移时间')
    ws_4s.write(row_4s, 2, '手术时间')
    ws_4s.write(row_4s, 3, '新辅助开始时间')
    ws_4s.write(row_4s, 4, '原DFS')
    ws_4s.write(row_4s, 5, '最后一次随访日期')
    ws_4s.write(row_4s, 6, '最后一次治疗时间')
    pid_set = base_info_pid('4s.xlsx', '基本信息_jibenxinxi')
    for pid in pid_set:
        row_4s += 1
        ws_4s.write(row_4s, 0, pid)

        try:
            change_time[pid]
        except:
            ws_4s.write(row_4s, 1, '')
        else:
            ws_4s.write(row_4s, 1, change_time[pid])

        try:
            operate_map[pid]
        except:
            ws_4s.write(row_4s, 2, '')
        else:
            ws_4s.write(row_4s, 2, operate_map[pid])

        try:
            new_start[pid]
        except:
            ws_4s.write(row_4s, 3, '')
        else:
            ws_4s.write(row_4s, 3, new_start[pid])

        try:
            dfs[pid]
        except:
            ws_4s.write(row_4s, 4, '')
        else:
            ws_4s.write(row_4s, 4, dfs[pid])

        try:
            visit_time[pid]
        except:
            ws_4s.write(row_4s, 5, '')
        else:
            ws_4s.write(row_4s, 5, visit_time[pid])

        try:
            cured_time[pid]
        except:
            ws_4s.write(row_4s, 6, '')
        else:
            ws_4s.write(row_4s, 6, cured_time[pid])


# 比较2个dict分组后的不同
def different(one, same):
    different_map = {}
    # 找各自不同的
    for same_k, same_v in one.items():
        try:
            same[same_k]
        except:
            different_map[same_k] = same_v
    return different_map


# 获取307参数
def get_parameter_307():
    # 获取307的数据
    # 1 复发转移时间
    change_time = one_parameter('307.xlsx', '疗效评价-疗效评价', '评估日期', '评估前状态', '辅助治疗')

    # 2 一线治疗开始时间
    first_one = one_parameter('307.xlsx', '化疗及靶向治疗-化疗及靶向治疗', '本次治疗开始日期', '治疗目的', '一线治疗')
    second_one = one_parameter('307.xlsx', '内分泌治疗-内分泌治疗', '开始用药日期', '治疗目的', '一线治疗')
    # 找到相同的pid
    one_start_same = get_all_compared(first_one, second_one)
    one_start_same = compared_list(one_start_same,)
    #  找到不同的
    first_diff_one = different(first_one, one_start_same)
    second_diff_one = different(second_one, one_start_same)
    # 合并所有
    first_diff_one.update(one_start_same)
    first_diff_one.update(second_diff_one)
    # 3 手术时间
    operate_map = one_parameter('307.xlsx', '手术治疗-手术治疗', '手术日期')
    # 函数的key转换为float类型

    # 4 新辅助开始时间
    first_new = one_parameter('307.xlsx', '化疗及靶向治疗-化疗及靶向治疗', '本次治疗开始日期', '治疗目的', '新辅助治疗')
    second_new = one_parameter('307.xlsx', '内分泌治疗-内分泌治疗', '开始用药日期', '治疗目的', '新辅助治疗')
    new_start_same = get_all_compared(first_new, second_new)
    new_start_same = compared_list(new_start_same)
    #  找到不同的
    first_diff_new = different(first_new, new_start_same)
    second_diff_new = different(second_new, new_start_same)
    # 合并所有
    first_diff_new.update(new_start_same)
    first_diff_new.update(second_diff_new)

    # 5 诊断时间
    diagnosis_time = one_parameter('307.xlsx', '基本信息-基本信息', '首次确诊日期')
    # 6 最后一次随访日期
    visit_time = one_parameter('307.xlsx', '随访-复发及生存随访', '随访日期', compared='not_Reverse')
    # 7 最后一次治疗时间
    cured_time_first = one_parameter('307.xlsx', '化疗及靶向治疗-化疗及靶向治疗', '本次治疗开始日期', '治疗目的', '辅助治疗', compared='not_Reverse')
    cured_time_second = one_parameter('307.xlsx', '内分泌治疗-内分泌治疗', '开始用药日期', '治疗目的', '辅助治疗', compared='not_Reverse')
    cured_time_third = one_parameter('307.xlsx', '手术治疗-手术治疗', '手术日期', compared='not_Reverse')
    # 3组比较
    cured_same_1 = get_all_compared(cured_time_first, cured_time_second)
    cured_same_2 = get_all_compared(cured_same_1, cured_time_third)
    cured_map = compared_list(cured_same_2, compared='not_Reverse')
    # 将获取的数值按pid进行整合,复发转移时间,一线开始时间,手术时间,新辅助开始时间,诊断时间,最后一次随访日期
    write_307(change_time, operate_map, first_diff_new, visit_time, first_diff_one, diagnosis_time, cured_map)


def get_parameter_4s():
    # 1 复发转移时间
    change_time = one_parameter('4s.xlsx', '复发转移_recum_BC', '复发转移日期')
    # 2 手术时间
    operate_map = one_parameter('4s.xlsx', '手术治疗_surgry', '手术日期')
    # 3 新辅助开始时间
    new_start = one_parameter('4s.xlsx', '术前化疗_preo_chem,术前内分泌治疗_preo_ecthero,术前靶向治疗_preo_targ', '治疗开始日期')
    # 4 最后一次随访日期
    visit_time = one_parameter('4s.xlsx', '随访信息1_foup_BC', '末次随访日期')
    # 5 DFS
    DFS = one_parameter('4s.xlsx', '复发转移_recum_BC', 'DFS', compared='not_Reverse')
    # 6 最后一次治疗时间
    cured_time1 = one_parameter('4s.xlsx', '手术治疗_surgry', '手术日期', compared='not_Reverse')
    cured_time2 = one_parameter('4s.xlsx', '辅助化疗_adchem_BC', '化疗开始日期（索引）', compared='not_Reverse')
    cured_time3 = one_parameter('4s.xlsx', '辅助靶向治疗_adjutar_BC', '靶向治疗开始日期（索引）', compared='not_Reverse')
    cured_time4 = one_parameter('4s.xlsx', '辅助内分泌治疗_adjuec_BC', 'TAM类开始日期', compared='not_Reverse')
    cured_time5 = one_parameter('4s.xlsx', '辅助内分泌治疗_adjuec_BC', 'AI类开始日期', compared='not_Reverse')
    # 分组
    cured_same1 = get_all_compared(cured_time1, cured_time2, store='4s.xlsx', sheet='基本信息_jibenxinxi')
    cured_same2 = get_all_compared(cured_time3, cured_time4, store='4s.xlsx', sheet='基本信息_jibenxinxi')
    cured_same3 = get_all_compared(cured_same1, cured_same2, store='4s.xlsx', sheet='基本信息_jibenxinxi')
    cured_same4 = get_all_compared(cured_same3, cured_time5, store='4s.xlsx', sheet='基本信息_jibenxinxi')
    cured_time = compared_list(cured_same4, compared='not_Reverse')
    # 参数写入
    write_4s(change_time, operate_map, new_start, visit_time, DFS, cured_time)


# 代码开始的地方
if __name__ == '__main__':
    try:
        init_data()
        # f = open('data_all', 'r')
        # data_all = json.load(f)
        # get_parameter_307()
        # get_parameter_4s()
        # wb.save('data_out/all_DFS.xls')
        # f.close()
        print('over')
    except Exception as e:
        print('报错', e)