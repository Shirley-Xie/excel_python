import os
import xlrd
import xlwt
import json
import abvd
import writeExcel
data_all = {}
import cscobc, yangsen
import lungcancer
import T

# 转化为json格式方便读取
def init_data(indir, outdir, name):
    global data_all
    # data_all = {}
    # 录入文件，输出文件
    path_dir = indir
    out_dir = outdir
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
                            print('filename:', filename, ' sheet_name:', sheet_name, ' row_index:', row_index, ' col_index:', col_index)
                        header = headers[col_index]
                        line[header] = value
                    data_all[filename][sheet_name].append(line)
        # out_file_name = out_dir + '/' + filename
        # f = open(out_file_name, 'w')
        # f.write(json.dumps(data_all[filename]))
        # f.close()

    f = open(out_dir + '/' + name, 'w')
    f.write(json.dumps(data_all))
    f.close()


# 和并多字典
def merge_dicts(merge_map, cure_type, names):
    name_map = {'1360': 'ABVD', '1363': 'CHOP', '1414': 'Gemox', '1416': 'R+OBP', '1536': 'CHOP-B', '1537': 'R-Gemox',
                '1603': 'AVD',
                '1605': 'ECOP', '1606': 'GP', '1643': 'EP', '1672': 'R-CHOP', '1673': 'RGL', '2039': 'R-COP',
                '2066': 'R-ABVD',
                '2141': 'BVD', '2647': '盐酸米托蒽醌脂质体试验'}
    for pid, vls in cure_type.items():
        for k, v in vls.items():
            if pid not in merge_map.keys():
                new_dict = [''] * 11
                merge_map[pid] = new_dict
            else:
                new_dict = merge_map[pid]
            # 按照名字排列,格式
            if k in names.keys():
                if v in name_map.keys():
                    v = name_map[v]
                new_dict[names[k]] = v


def csc(data_all):
    cs = cscobc.Cscobc(data_all)
    # base_dict = cs.direct_p()
    # head = ['pid', '性别', '医院-科室', '出生日期', '是否绝经', '月经史', '病理报告日期（索引）', '手术日期', '首次确诊年龄', '临床N分期']
    # wr = writeExcel.WriteExcel(base_dict.keys(), '基本信息')
    # wr.write_arr(head, base_dict)
    # wr.last('基本信息')
    qian_map = cs.qian_shao()
    # dict_16_23 = cs.undirect_16_23()
    head1 = ['pid', '前哨阳性个数']
    # , '新辅助AI内分泌治疗', '辅助AI内分泌治疗', '是否复发转移', '复发转移后是否内分泌治疗',
    #          '微转移', '临床病理类型不为肉瘤', '初诊颅外病灶', '腋窝临床检查阳性', '一线内分泌治疗'
    wr = writeExcel.WriteExcel(data_all, [['cscobc.xlsx', '基本信息_jibenxinxi_4578'], ['bc-s_new.xlsx', '基本信息-基本信息']],'qianshao_num')
    wr.write_excel(head1, qian_map)
    wr.last('qianshao_num')


def starting(data_all):
    # 进行操作
    read = abvd.Parameter(data_all)
    # 性别年龄，数组字典
    sex_old_dict = read.sex_age()
    # orr,化疗线数，治疗方案， pd,非pd
    orr, pd, npd, cure_type = read.orr_pd()
    for pid, v in orr.items():
        if len(v) < 1:
            orr[pid] = ''
        else:
            for l in v:
                if l in ['3', 3]: orr[pid] = 'SD'
                elif l in [2, '2']: orr[pid] = 'PR'
                elif l in [1, '1']: orr[pid] = 'CR'
                elif l in [4, '4']: orr[pid] = 'PD'
                else: orr[pid] = l
    # pip评估
    pip_grade = read.pip_()
    # 基线的临床分期,一般字典
    clinical_stage = {}
    read.clinical_stages(clinical_stage, 'ABVD.xls', '入院诊断-临床诊断')
    read.clinical_stages(clinical_stage, 'ABVD4S.xlsx', '--临床诊断与分期')
    # 进行间隔的提取,一般字典
    abvd_gap = {}
    read.time_gap(abvd_gap, 'ABVD4S.xlsx', '--系统治疗', '治疗方案', '本周期开始时间', 'ABVD')
    read.time_gap(abvd_gap, 'ABVD.xls', '治疗-化疗', '化疗方案', '化疗时间', 'ABVD')
    abvd_two = read.gap_two(abvd_gap)
    # 免疫组化结果取值 字典跟着字典形式
    immune_dict = {}
    read.immune('ABVD4S.xlsx', '病理检查lym--淋巴结或组织活检', immune_dict)
    read.immune_307('ABVD.xls', '入院诊断-病理诊断', immune_dict)
    pfs_dict = read.pfs()
    # 将数转变为数组形式4+6
    merge_map = {}
    names = {'一线': 0, '1': 0, '二线': 1, 2: 1, '三线': 2, '3': 2,  '四线': 3,  '4': 3,  '四线及以上': 4, 'CD3': 5, 'CD5': 6, 'CD15': 7, 'CD20': 8, 'BCL-6': 9, 'Ki67':10, 'KI-67':10}
    merge_dicts(merge_map, cure_type, names)
    merge_dicts(merge_map, immune_dict, names)

    head1 = ['pid', '性别', '年龄', 'pfs', 'orr', 'pd', '非pd', 'IPI评估(分数)', '时间间隔', '临床分期', '一线', '二线', '三线', '四线', '四线及以上','CD3', 'CD5', 'CD15', 'CD20', 'BCL-6', 'Ki67']
    # 写入
    wr = writeExcel.WriteExcel(data_all, [['ABVD4S.xlsx', '--基本信息'], ['ABVD.xls', '基本信息-基本信息']], 'abvd')
    wr.write_excel(head1, sex_old_dict, pfs_dict, orr, pd, npd, pip_grade, abvd_two, clinical_stage, merge_map)
    wr.last('abvd综合表1')

if __name__ == '__main__':
    # Driver program
    i = -1
    j = 4

    if i == 0:
        init_data('./data', './data_read', 'data_all')
    elif i == 1:
        init_data('./data/data1', './data_read', 'data_all_1')
    elif i == 2:
        init_data('./data/yangsen', './data_read', 'data_ys')
    elif i == 3:
        init_data('./data/data_lung', './data_read', 'data_lung')
    elif i == 4:
        init_data('./data/dataT', './data_read', 'dataT')
    print('data')

    if j == 0:
        f = open('./data_read/data_all', 'r')
        data_all = json.load(f)
        starting(data_all)
    elif j == 1:
        f = open('./data_read/data_all_1', 'r')
        data_all = json.load(f)
        csc(data_all)
    elif j == 2:
        f = open('./data_read/data_ys', 'r')
        data_all = json.load(f)
        ys = yangsen.YangSen(data_all)
        ys.yangSen_starting()
    elif j == 3:
        f = open('./data_read/data_lung', 'r')
        data_all = json.load(f)
        lung = lungcancer.LungCancer(data_all)
        lung.stages()
    elif j == 4:
        f = open('./data_read/dataT', 'r')
        data_all = json.load(f)
        t = T.T(data_all)
        t.driver()

    print('over')