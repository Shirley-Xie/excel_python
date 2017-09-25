import xlwt
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

7 化疗及靶向治疗-化疗及靶向治疗=>（治疗目的==辅助治疗 and 本次治疗开始日期!=''）(1)
  内分泌治疗-内分泌治疗=>=>（治疗目的==辅助治疗 and 开始用药日期!='')(2)
  (1) (2)（辅助开始时间）

4s:
1 复发转移_recum_BC => 复发转移日期（复发转移时间）
  5        => DFS（DFS）
2 手术治疗_surgry => 手术日期	（手术时间）

3 术前化疗_pre0-chem => 治疗开始日期
术前内分泌治疗_preo_ecthero => 治疗开始日期
术前靶向治疗_preo_targ => 治疗开始日期
选择最早的时间（新辅助开始时间)

4 随访信息_foup_BC => 末次随访日期（最后一次随访日期）

6 辅助化疗_adchem_BC => 化疗开始日期（索引） (1)
  辅助内分泌治疗_adjuec_BC => TAM类开始日期 (2)
  。。。							AI类开始日期，（3）
  辅助靶向治疗_adjutar_BC => 靶向治疗开始日期（索引） (3)
  1 2 3（辅助开始时间
7 辅助开始时间
8诊断时间 
"""

wb = xlwt.Workbook()
ws_307 = wb.add_sheet('DFS_s_307')
ws_4s = wb.add_sheet('DFS_s_4s')
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


def write_307(change_time, operate_map, new_start, visit_time, one_start, diagnosis_time, assist_cured):
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
    ws_307.write(row, 7, '辅助开始时间')
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
            assist_cured[pid]
        except:
            ws_307.write(row, 7, '')
        else:
            ws_307.write(row, 7, assist_cured[pid])


def write_4s(change_time, operate_map, new_start, visit_time, dfs, assist_cured, look_time):
    global row_4s
    # data_all = DATA_ALL
    ws_4s.write(row_4s, 0, 'pid')
    ws_4s.write(row_4s, 1, '复发转移时间')
    ws_4s.write(row_4s, 2, '手术时间')
    ws_4s.write(row_4s, 3, '新辅助开始时间')
    ws_4s.write(row_4s, 4, '原DFS')
    ws_4s.write(row_4s, 5, '最后一次随访日期')
    ws_4s.write(row_4s, 6, '辅助开始时间')
    ws_4s.write(row_4s, 7, '诊断时间')
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
            assist_cured[pid]
        except:
            ws_4s.write(row_4s, 6, '')
        else:
            ws_4s.write(row_4s, 6, assist_cured[pid])

        try:
            look_time[pid]
        except:
            ws_4s.write(row_4s, 7, '')
        else:
            ws_4s.write(row_4s, 7, look_time[pid])


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
    one_start_same = compared_list(one_start_same)
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
    # 7 （辅助开始时间）
    cured_time_first = one_parameter('307.xlsx', '化疗及靶向治疗-化疗及靶向治疗', '本次治疗开始日期', '治疗目的', '辅助治疗')
    cured_time_second = one_parameter('307.xlsx', '内分泌治疗-内分泌治疗', '开始用药日期', '治疗目的', '辅助治疗')
    cured_same = get_all_compared(cured_time_first, cured_time_second)
    cured_map = compared_list(cured_same)
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
    # 7 辅助开始时间
    assist_1 = one_parameter('4s.xlsx', '辅助化疗_adchem_BC', '化疗开始日期（索引）')
    assist_2 = one_parameter('4s.xlsx', '辅助内分泌治疗_adjuec_BC', 'TAM类开始日期')
    assist_3 = one_parameter('4s.xlsx', '辅助内分泌治疗_adjuec_BC', 'AI类开始日期')
    assist_4 = one_parameter('4s.xlsx', '辅助靶向治疗_adjutar_BC', '靶向治疗开始日期（索引）')
    assist_cmp_1 = get_all_compared(assist_1, assist_4, store='4s.xlsx', sheet='基本信息_jibenxinxi')
    assist_cmp_2 = get_all_compared(assist_2, assist_3, store='4s.xlsx', sheet='基本信息_jibenxinxi')
    assist_cmp_3 = get_all_compared(assist_cmp_1, assist_cmp_2, store='4s.xlsx', sheet='基本信息_jibenxinxi')
    assist_cured = compared_list(assist_cmp_3)
    # 8诊断时间
    look_time = one_parameter('4s.xlsx', '初诊病理检测_init_path', '检查日期')
    # 参数写入
    write_4s_1(look_time)
    write_4s(change_time, operate_map, new_start, visit_time, DFS, assist_cured, look_time)


def write_4s_1(look_time):
    global row_4s
    # data_all = DATA_ALL
    ws_4s.write(row_4s, 0, 'pid')
    ws_4s.write(row_4s, 1, '诊断时间')
    pid_set = base_info_pid('4s.xlsx', '基本信息_jibenxinxi')
    for pid in pid_set:
        row_4s += 1
        ws_4s.write(row_4s, 0, pid)

        try:
            look_time[pid]
        except:
            ws_4s.write(row_4s, 1, '')
        else:
            ws_4s.write(row_4s, 1, look_time[pid])

# 代码开始的地方
if __name__ == '__main__':
    try:
        f = open('data_all', 'r')
        data_all = json.load(f)
        # get_parameter_307()
        # get_parameter_4s()
        # 8诊断时间
        look_time = one_parameter('4s.xlsx', '初诊病理检测_init_path', '检查日期')
        # 参数写入
        write_4s_1(look_time)
        wb.save('data_out/DFS_诊断时间.xls')
        f.close()
        print('DFS_DFS_诊断时间')
    except Exception as e:
        print(e)