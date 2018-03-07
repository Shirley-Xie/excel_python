import re
import writeExcel

class T:
    def __init__(self, data_all):
        self.data_all = data_all
        self.re_little = r'来曲唑|阿那曲唑|依西美坦|AI|EXE|Let|瑞宁得|芙瑞|弗隆|阿诺新'
        self.re_many = r'来曲唑|阿那曲唑|依西美坦|氟维司群|甲羟孕酮|孕激素|他莫昔芬|TAM|托瑞米芬|AI|EXE|Let|瑞宁得|芙瑞|弗隆|芙仕得|阿诺新'
        self.ai = ['来曲唑', '阿那曲唑', '依西美坦', "AI", "EXE", "Let", "瑞宁得", "芙瑞", "弗隆", "阿诺新"]
        self.hou = ['来曲唑','阿那曲唑','依西美坦','氟维司群','甲羟孕酮', '孕激素', '他莫昔芬', 'TAM','托瑞米芬','AI','EXE','Let','瑞宁得', '芙瑞', '弗隆', '芙仕得','阿诺新']
    # 临床T分期
    def clinical(self):
        # 这里格式不同，后面需要处理
        clinical = {}

        def clini_T(exl, sheet, _name, rule_found, rule_choose):
            # 找到最大值,不写参数可以识别吗
            def find_max(rule):
                init = 0
                for name in line:
                    if re.findall(rule, name):
                        # 找到最大值
                        if int(line['name']) > init:
                            init = line['name']
                if init <= 20:
                    init = 'T1'
                elif 50 <= init < 20:
                    init = 'T2'
                elif init > 50:
                    init = 'T3'
                clinical[pid].appent([init, check])

            # 规则操作
            for line in self.data_all[exl][sheet]:
                pid = line['pid']
                name_T = line[_name]
                if name_T == '':
                    continue
                # 若找到4种值则直接返回,判断是否进行不是4种情况的操作
                flag = False
                if '307' in exl:
                    found = re.findall(rule_found, name_T)
                    if found:
                        flag = True
                        if pid not in clinical.keys():
                            clinical[pid] = found[0]
                else:
                    dict_T = {1: 'Tx', 2: 'T0', 3: 'Tis', 11: 'T4'}
                    # 若找到4种值则直接返回
                    if name_T in rule_found:
                        flag = True
                        if pid not in clinical.keys():
                            clinical[pid] = dict_T[line[name_T]]
                # 存在则直接给值否则取最大值，因为不相信自己填写的数据
                if not flag:
                    # 多次pid则找到最早的那条,将其保留进行下次比较
                    check = line['检查日期']
                    if not check:
                        continue
                    if pid not in clinical.keys():
                        # 找到字段带有"病灶-宽""病灶-高""病灶-最长径"的最大值,
                        find_max(rule_choose)
                    else:
                        # 进行比较
                        if clinical[pid][1] > check:
                            find_max(rule_choose)
        clini_T('307.xlsx', '临床诊断-临床诊断',  '原发肿瘤分期（T）', r'Tx|Tis|T0|T4', r'病灶-宽|病灶-高|病灶-最长径')
        clini_T('4s.xlsx', '初诊肿瘤临床分期_init_cstage_7455',  'T分期-BC（AJCC第七版）', [1, 2, 3, 4], r'肿瘤直径')
        return clinical

    def p3_6(self):
        # 参数3：初诊颅外病灶, luwai
        luwai = set()
        # 307
        for line in self.data_all['307.xlsx']['临床诊断-临床诊断']:
            pid = line['pid']
            if pid == '' or pid == 'pid':
                continue
            if line['转移灶部位-选项[对侧乳腺]'] or line['转移灶部位-选项[肝]'] or line['转移灶部位-选项[胸膜]'] or\
                    line['转移灶部位-选项[肺]'] or line['转移灶部位-选项[骨]'] or line['转移灶部位-选项[软组织]']:
                luwai.add(pid)
        # 4s
        for line in self.data_all['4s.xlsx']['初诊肿瘤临床分期_init_cstage_7455']:
            pid = line['pid']
            if pid == '' or pid == 'pid':
                continue
            if line['新增转移部位'] not in [1, '1']:
                luwai.add(pid)

        # 参数4和5的函数
        def ai_cure(style, sheet, cure):
            for line in self.data_all['307.xlsx']['内分泌治疗-内分泌治疗']:
                pid = line['pid']
                if pid == '' or pid == 'pid':
                    continue
                if line['内分泌治疗目的'] != style:
                    continue
                else:
                    # 找到匹配的字段名称
                    for li in line:
                        head = re.findall(r'药物通用名-', li)
                        if head:
                            if line[li] in self.ai:
                                cure.add(pid)

            for line in self.data_all['4s.xlsx'][sheet]:
                pid = line['pid']
                if pid == '' or pid == 'pid':
                    continue
                if line['内分泌治疗方案'] in [2, '2']:
                    cure.add(pid)
        # 参数4：新辅助AI内分泌治疗, new_cure_ai
        new_cure_ai = set()
        # 参数5：辅助AI内分泌治疗,cure_ai
        cure_ai = set()
        ai_cure('新辅助治疗', '术前内分泌治疗_preo_ecthero_28594', new_cure_ai)
        ai_cure('辅助治疗', '辅助内分泌治疗_adjuec_BC_4637', cure_ai)

        # 参数6：复发转移后是否内分泌治疗
        # 307
        recrudesce = set()
        for line in self.data_all['307.xlsx']['内分泌治疗-内分泌治疗']:
            pid = line['pid']
            if pid == '' or pid == 'pid':
                continue
            if line['内分泌治疗目的'] not in ['新辅助治疗', '辅助治疗', '']:
                for li in line:
                    if line[li] in self.ai:
                        recrudesce.add(pid)
        # 4s
        for line in self.data_all['4s.xlsx']['复发转移_recum_BC_36112']:
            pid = line['pid']
            if pid == '' or pid == 'pid':
                continue
            txt = line['治疗方案（文本型）']
            if txt != '':
                if re.findall(r'来曲唑|阿那曲唑|依西美坦|AI|EXE|Let|瑞宁得|芙瑞|弗隆|阿诺新', str(txt)):
                    recrudesce.add(pid)
        return luwai, new_cure_ai, cure_ai, recrudesce

    # 返回判断为有值的字段,
    def para_yes(self, exl, sheet, k_v, con, bingli, wu=True):
        for line in self.data_all[exl][sheet]:
            pid = line['pid']
            if pid == '' or pid == 'pid':
                continue
            if wu:
                v = line[k_v[0]] == k_v[1]
            else:
                v = True
            if v:
                for i in con:
                    if line[i]:
                        if pid in bingli.keys():
                            name = bingli[pid]
                            name = name + "," + i
                            bingli[pid] = name
                        else:
                            bingli[pid] = i

    def add_str(self, cure_out, pid, value):

        if pid not in cure_out.keys():
            cure_out[pid] = value
        else:
            if value != cure_out[pid]:
                cure_out[pid] = cure_out[pid] + ',' + value

    def ai_cure(self, style, sheet, name, cure_out):
        for line in self.data_all['307.xlsx']['内分泌治疗-内分泌治疗']:
            pid = line['pid']
            if pid == '' or pid == 'pid':
                continue
            if line['内分泌治疗目的'] != style:
                continue
            else:
                # 找到匹配的字段名称
                for li in line:
                    head = re.findall(r'药物通用名-', li)
                    if head:
                        if line[li] in self.ai:
                            self.add_str(cure_out, pid, line[li])
        # dict
        dict_ = {1: '来曲唑', 2: '阿那曲唑', 3: '依西美坦', 999: '', '': ''}
        for line in self.data_all['4s.xlsx'][sheet]:
            pid = line['pid']
            if pid == '' or pid == 'pid':
                continue
            if line['内分泌治疗方案'] in [2, '2']:
                self.add_str(cure_out, pid, dict_[line[name]])

    def get_date(self, exl, sheet, field, name, date_dict):
        for line in self.data_all[exl][sheet]:
            pid = line['pid']
            if pid == '' or pid == 'pid':
                continue
            exc = '一线治疗'
            if field != 0:
                exc == line[field]
            if exc == '一线治疗':
                cure_time = line[name]
                if cure_time not in ['', '1970-01-01']:
                    # 取最早的值
                    if pid not in date_dict.keys():
                        date_dict[pid] = cure_time
                    else:
                        # 比较大小，取大
                        if name in ['无病生存（DFS）', '随访日期']:
                            if cure_time > date_dict[pid]:
                                date_dict[pid] = cure_time
                        # 取小
                        elif cure_time < date_dict[pid]:
                            date_dict[pid] = cure_time

    def group_a(self, pid, d_map, i, value):
        # 第一次使用
        if pid not in d_map.keys():
            new_dict = ['']*9
            d_map[pid] = new_dict
        else:
            new_dict = d_map[pid]
        new_dict[i] = value

    def direct(self, base_dict):
        # 直接提取
        for line in self.data_all['4s.xlsx']['基本信息_jibenxinxi_4578']:
            pid = line['pid']
            if pid == '' or pid == 'pid':
                continue
            sex_num = {1: '男', 2: '女', 998: '', 999: '', '': ''}
            m_num = {1: '是', 2: '否', 998: '', 999: '', '': ''}
            self.group_a(pid, base_dict, 0, sex_num[line['性别']])
            self.group_a(pid, base_dict, 1, line['出生日期'])
            self.group_a(pid, base_dict, 2, m_num[line['是否绝经']])

        # 临床分期
        for line in self.data_all['4s.xlsx']['初诊肿瘤临床分期_init_cstage_7455']:
            pid = line['pid']
            if pid == '' or pid == 'pid':
                continue
            namemap = {1: 'NX  区域淋巴结不能确定（例如曾经切除）',
            2: 'N0  区域淋巴结无转移',
            3: 'N1  同侧腋窝淋巴结转移，可活动',
            4: 'N2  同侧腋窝淋巴结转移，固定或相互融合或缺乏同侧腋窝淋巴结转移的临床证据，但临床上发现*有同侧内乳淋巴结转移',
            5: 'N2a  同侧腋窝淋巴结转移，固定或相互融合',
            6: 'N2b  仅临床上发现*同侧腋窝淋巴结转移，而无同侧腋窝淋巴结转移的临床证据',
            7: 'N3 同侧锁骨下淋巴结转移伴或不伴有腋窝淋巴结转移；或临床上发现*同侧内乳淋巴结转移和腋窝淋巴结转移的临床证据；或同侧锁骨上淋巴结转移伴或不伴腋窝或内乳淋巴结转移',
            8: 'N3a  同侧锁骨下淋巴结转移',
            9: 'N3b  同侧内乳淋巴结及腋窝淋巴结转移',
            10: 'N3c  同侧锁骨上淋巴结转移',
            999: '', 998: '', '': ''}
            self.group_a(pid, base_dict, 3, namemap[line['N分期-BC（AJCC第七版）']])

        for line in self.data_all['4s.xlsx']['手术治疗_surgry_4687']:
            pid = line['pid']
            if pid == '' or pid == 'pid':
                continue
            self.group_a(pid, base_dict, 4, line['病理报告日期（索引）'])
            self.group_a(pid, base_dict, 5, line['手术日期'])
            # 参数32：阳性淋巴结
            if line['淋巴结术式'] in ['1', 1]:
                su = line['阳性淋巴结数量']
                if su:
                    if isinstance(su, float) or isinstance(su, int):
                        self.group_a(pid, base_dict, 8, '是')

        for line in self.data_all['307.xlsx']['手术治疗-手术治疗']:
            pid = line['pid']
            if pid == '' or pid == 'pid':
                continue
            self.group_a(pid, base_dict, 5, line['手术日期'])

        # 307
        for line in self.data_all['307.xlsx']['基本信息-基本信息']:
            pid = line['pid']
            if pid == '' or pid == 'pid':
                continue
            self.group_a(pid, base_dict, 0, line['性别'])
            self.group_a(pid, base_dict, 6, line['首次确诊年龄'])
            yue = {'未绝经': '否', '已绝经': '是', '不确定': '', '其他': '', '': ''}
            self.group_a(pid, base_dict, 2, yue[line['月经史']])

        #  临床诊断，字段：区域淋巴结分期（N）4
        for line in self.data_all['307.xlsx']['临床诊断-临床诊断']:
            pid = line['pid']
            if pid == '' or pid == 'pid':
                continue
            self.group_a(pid, base_dict, 3, line['区域淋巴结分期（N）'])

        for line in self.data_all['307.xlsx']['病理诊断-病理诊断']:
            pid = line['pid']
            if pid == '' or pid == 'pid':
                continue
            # 前哨阳性307
            if line['前哨淋巴结转移情况-选项[阳性]'] == '是':
                self.group_a(pid, base_dict, 8, '是')

        for line in self.data_all['4s.xlsx']['初诊肿瘤临床分期_init_cstage_7455']:
            pid = line['pid']
            if pid == '' or pid == 'pid':
                continue
            if line['送检淋巴结部位'] in [3, '3'] and line['淋巴结大小'] not in ['', 0, 0.0, '0']:
                self.group_a(pid, base_dict, 7, '是')

    def add_all(self, qian_map, pid, su):
        if pid not in qian_map.keys():
            new_dict = 0
        else:
            new_dict = qian_map[pid]
        qian_map[pid] = new_dict + su

    # 参数37：前哨阳性数量
    def qian_shao(self):
        # 307
        qian_map = {}
        for line in self.data_all['307.xlsx']['病理诊断-病理诊断']:
            pid = line['pid']
            if pid == '' or pid == 'pid':
                continue
            # 前哨阳性307
            if line['阳性(个)-1']:
                self.add_all(qian_map, pid, line['阳性(个)-1'])

        # 4s
        for line in self.data_all['4s.xlsx']['手术治疗_surgry_4687']:
            pid = line['pid']
            if pid == '' or pid == 'pid':
                continue
            if line['淋巴结术式'] in ['1', 1]:
                su = line['阳性淋巴结数量']
                if su:
                    if isinstance(su, float) or isinstance(su, int):
                        # 将有的值相加
                        self.add_all(qian_map, pid, su)
        return qian_map

    def add_set(self, pid, line_32, start):
        if pid not in line_32.keys():
            a = set()
            line_32[pid] = a
            a.add(start)
        else:
            x = line_32[pid]
            x.add(start)

    def para32_38(self):
        # 参数35：AI治疗失败后后续内分泌治疗开始用药日期-，start_medi
        # 参数36： AI治疗失败后后续内分泌治疗结束用药日期-1，end_medi
        # 参数37：AI治疗失败后后续内分泌治疗结束治疗原因-1, TT
        # 参数38：AI治疗失败后后续内分泌治疗疗效评估-1, evaluate
        # 参数32：AI治疗失败后后续内分泌治疗总线数-1, line_32_last
        # 参数33：AI治疗失败后后续内分泌治疗分线数-1, case_34
        # 参数34：AI治疗失败后后续内分泌治疗方案-1,
        start_medi = {}
        end_medi= {}
        TT = {}
        evaluate_before = {}
        evaluate_after = {}
        case_34 = {}
        for line in self.data_all['307.xlsx']['内分泌治疗-内分泌治疗']:
            pid = line['pid']
            if pid == '' or pid == 'pid':
                continue
            if line['内分泌治疗目的'] not in ['', '新辅助治疗', '辅助治疗']:
                # 找到匹配的字段名称，参数35，36
                for li in line:
                    head = re.findall(r'药物通用名-', li)
                    if head:
                        if line[li] in self.hou:
                            # 参数34,307
                            self.add_str(case_34, pid, line[li])

                            if line['开始用药日期'] not in ['', '1970-1-1', '1970-01-01']:
                                if pid not in start_medi.keys():
                                    start_medi[pid] = line['开始用药日期']

                            end_time = line['结束用药日期']
                            if end_time not in ['', '1970-1-1', '1970-01-01']:
                                if pid not in end_medi.keys():
                                    end_medi[pid] = end_time

                            # 参数37
                            if line['更换内分泌治疗原因-选项[疾病进展]']:
                                TT[pid] = 'TTP'
                            else:
                                for i in ['更换内分泌治疗原因-选项[满5年]', '更换内分泌治疗原因-选项[满10年]', '更换内分泌治疗原因-选项[按既定方案换药]', '更换内分泌治疗原因-选项[其他]', '更换内分泌治疗原因-选项[依从性]']:
                                    if line[i]:
                                        TT[pid] = 'TTF'
        # 参数38的307
        for line in self.data_all['307.xlsx']['疗效评价-疗效评价']:
            pid = line['pid']
            if pid == '' or pid == 'pid':
                continue
            # 参数38，在35和36之间
            pg_date, pg_g = line['评估日期'], line['目标病灶评估']
            if pid in start_medi.keys():
                if pg_date and pg_g and pg_date < start_medi[pid]:
                    self.add_str(evaluate_before, pid, pg_g)
            # 参数38，在35和36之间, 求交集即可，为了防止前一个有后一个没有的判断
            if pid in end_medi.keys():
                pg_date, pg_g = line['评估日期'], line['目标病灶评估']
                if pg_date and pg_g and pg_date > end_medi[pid]:
                    self.add_str(evaluate_after, pid, pg_g)

        # 此处将两个集合取交集
        pids = set(evaluate_after.keys()) & set(evaluate_before.keys())

        # 参数38最后数据集
        evaluate = {}; start_medi_set = {}
        for pid in pids:
            evaluate[pid] = evaluate_after[pid]

        for line in self.data_all['4s.xlsx']['复发转移_recum_BC_36112']:
            pid = line['pid']
            if pid == '' or pid == 'pid':
                continue
            # 参数32的4S部分，获取开始的所有日期
            start_date = line['用药开始日期']
            if start_date not in ['', '1970-1-1', '1970-01-01']:
                self.add_set(pid, start_medi_set, start_date)

            txt = line['治疗方案（文本型）']
            if txt != '':
                if re.findall(r'来曲唑|阿那曲唑|依西美坦|氟维司群|甲羟孕酮|孕激素|他莫昔芬|TAM|托瑞米芬|AI|EXE|Let|瑞宁得|芙瑞|弗隆|芙仕得|阿诺新', str(txt)):
                    # 参数34，4s
                    self.add_str(case_34, pid, txt)
                    if start_date not in ['', '1970-1-1', '1970-01-01']:
                        if pid not in start_medi.keys():
                            start_medi[pid] = start_date
                    if line['用药结束日期'] not in ['', '1970-1-1', '1970-01-01']:
                        if pid not in end_medi.keys():
                            end_medi[pid] = line['用药结束日期']
                    # 参数37
                    if line['疾病进展时间'] not in ['', '1970-1-1', '1970-01-01']:
                        TT[pid] = 'TTP'
                    else:
                        if line['治疗失败时间'] not in ['', '1970-1-1', '1970-01-01']:
                            TT[pid] = 'TTF'
                    # 参数38， 返回第一条满足的值
                    pg = line['总体疗效评估']
                    if pg:
                        map = {1: '完全缓解（CR）', 2: '部分缓解(PR)', 3: '稳定(SD)', 4: '进展(PD)', 5: '非CR/非PD（no-CR/no-PD）',
                         6: '不能确定的CR(CRu)', 7: 'VGPR', 997: '', 8: '完全缓解伴血细胞未完全回复（CRi）', 998: '', 999: ''}
                        if pid not in evaluate.keys():
                            evaluate[pid] = map[pg]

        """
        参数32，line_32_last
        """
        # 参数32：AI治疗失败后后续内分泌治疗总线数-1,307
        line_32_set = {}
        line_32_last = {}
        for line in self.data_all['307.xlsx']['化疗及靶向治疗-化疗及靶向治疗']:
            pid = line['pid']
            if pid == '' or pid == 'pid':
                continue
            if line['治疗目的'] not in ['新辅助治疗', '辅助治疗', '']:
                start = line['本次治疗开始日期']
                if start:
                    # 将所有的值放入集合
                    self.add_set(pid, line_32_set, start)

        # 参数32 35：start_medi
        self.find_index(start_medi, line_32_set, line_32_last)  # 307 结束
        # 4s
        self.find_index(start_medi, start_medi_set, line_32_last)
        return line_32_last, case_34, start_medi, end_medi, TT, evaluate

    # 307
    def p307(self, para_date, para_line, condition, end=False, end_early=False):
        for line in self.data_all['307.xlsx']['内分泌治疗-内分泌治疗']:
            pid = line['pid']
            if pid == '' or pid == 'pid':
                continue
            goal = line['内分泌治疗目的']
            # 找到用药最早时间
            start = line['开始用药日期']
            if condition(goal):
                flag = False
                for li in line:
                    head = re.findall(r'药物通用名-', li)
                    if head:
                        if line[li] in self.ai:
                            # 若有则继续
                            flag = True
                            break
                if flag:
                    if start:
                        if pid not in para_date.keys():
                            para_date[pid] = start
                        else:
                            if para_date[pid] < start:
                                para_date[pid] = start
                                # 转化为数字,一开始返回线数
                                if not (end or end_early):
                                    lies = {'一线治疗': 1, '二线治疗': 2, '三线治疗': 3, '四线治疗': 4, '五线及以上治疗': 5}
                                    para_line[pid] = lies[goal]

                                # 最后取线数,此处为汉字,区别线数和最后日期
                                if end:
                                    para_line[pid] = line[end]

                    # 返回最晚日期的最早的满足条件
                    if end_early:
                        end_date = line[end_early]
                        if pid not in para_line.keys():
                            para_line[pid] = end_date
                        else:
                            if para_line[pid] < end_date:
                                para_line[pid] = end_date

    def p4s(self, sheet, starting, ending, start_dates, end_dates):
        for line in self.data_all['4s.xlsx'][sheet]:
            pid = line['pid']
            if pid == '' or pid == 'pid':
                continue
            if line['内分泌治疗方案'] in [2, '2']:
                # 字段2
                start = line[starting]
                end = line[ending]
                if start not in ['', '1970-1-1', '1970-01-01']:
                    if pid not in start_dates.keys():
                        start_dates[pid] = start
                # 字段3
                if end not in ['', '1970-1-1', '1970-01-01']:
                    if pid not in end_dates.keys():
                        end_dates[pid] = start


    # v, set, out
    def find_index(self, one, sets, all_last):
        for k, v in one.items():
            if k in sets.keys():
                b = sorted(list(sets[k]))
                # v在b的位置
                if v in b:
                    all_last[k] = b.index(v)+1


    # 所有在此运行
    def driver(self):

        if True:
            base_dict = {}
            self.direct(base_dict)
            # '前哨阳性个数'4s
            qianshao = self.qian_shao()
            head = ['pid', '前哨阳性数量', '性别', '出生日期', '是否绝经', '区域淋巴结分期(N)', '病理报告日期', '手术日期', '首次确诊年龄', '腋窝临床检查阳性', '前哨阳性']
            wr = writeExcel.WriteExcel(self.data_all, [['4s.xlsx', '基本信息_jibenxinxi_4578'], ['307.xlsx', '基本信息-基本信息']],
                                       '前哨')
            wr.write_excel(head, qianshao, base_dict)
            wr.last('基本信息')
        if False:
            # 多个条件选择用数据遍历解决
            con1 = ['病理类型-选项[乳头Paget病]', '病理类型-选项[伴印戒细胞分化的癌]',
            '病理类型-选项[伴大汗腺分化的癌]', '病理类型-选项[伴髓样特征的癌]', '病理类型-选项[分泌性癌]', '病理类型-选项[导管内乳头状癌]',
            '病理类型-选项[导管原位癌]','病理类型-选项[小叶原位癌]', '病理类型-选项[小管癌]', '病理类型-选项[微小浸润性癌]','病理类型-选项[恶性叶状肿瘤]', '病理类型-选项[浸润性乳头状癌]',
            '病理类型-选项[浸润性导管癌（非特殊型）]', '病理类型-选项[浸润性小叶癌]', '病理类型-选项[浸润性微乳头状癌]', '病理类型-选项[浸润性癌]', '病理类型-选项[炎症性癌]', '病理类型-选项[筛状癌]',
            '病理类型-选项[非特殊型化生性癌]', '病理类型-选项[非特殊型浸润性癌]', '病理类型-选项[黏液癌]']

            con2 = ['复发/进展病灶部位-其他值',
             '复发/进展病灶部位-选项[其他]', '复发/进展病灶部位-选项[其他区域淋巴结]', '复发/进展病灶部位-选项[同侧乳腺]',
             '复发/进展病灶部位-选项[同侧胸壁]', '复发/进展病灶部位-选项[同侧腋窝]', '复发/进展病灶部位-选项[对侧乳腺]','复发/进展病灶部位-选项[肝]',
             '复发/进展病灶部位-选项[肺]', '复发/进展病灶部位-选项[胸膜]', '复发/进展病灶部位-选项[脑]', '复发/进展病灶部位-选项[远处软组织]', '复发/进展病灶部位-选项[骨]']

            con3 = ['新增转移部位', '新增转移部位[其他]', '复发转移部位（bc）', '复发转移部位（bc）[其他]']

            # 参数13：病理类型
            bl = {}
            self.para_yes('307.xlsx', '病理诊断-病理诊断', ['病理诊断方法', '手术后病理'],  con1, bl)

            # 参数22：复发转移部位
            fufa = {}
            self.para_yes('307.xlsx', '疗效评价-疗效评价', {0: 'wuyiyi'}, con2, fufa, wu=False)
            self.para_yes('4s.xlsx', '复发转移_recum_BC_36112', {0: 'wuyiyi'}, con3, fufa, wu=False)

            # # 参数18：辅助AI内分泌治疗, cure_out
            # cure_out = {}
            # self.ai_cure('辅助治疗', '辅助内分泌治疗_adjuec_BC_4637', 'AI类药物', cure_out)

            # 19 复发转移日期
            date19 = {}
            # 返回最早值
            self.get_date('307.xlsx', '化疗及靶向治疗-化疗及靶向治疗', '治疗目的', '本次治疗开始日期', date19)
            self.get_date('307.xlsx', '内分泌治疗-内分泌治疗', '内分泌治疗目的', '开始用药日期', date19)
            self.get_date('4s.xlsx', '复发转移_recum_BC_36112', 0, '复发转移日期', date19)
            # 20 4SDFS 取大
            dfs = {}
            self.get_date('4s.xlsx', '复发转移_recum_BC_36112', 0, '无病生存（DFS）', dfs)
            # 参数21：末次随访日期，取大
            visit = {}
            self.get_date('307.xlsx', '随访-复发及生存随访', 0, '随访日期', visit)
            self.get_date('4s.xlsx', '随访信息_foup_BC_4624', 0, '末次随访日期', visit)

            head = ['pid', '病理类型', '复发转移部位', '复发转移日期', '4SDFS', '末次随访日期']
            wr = writeExcel.WriteExcel(self.data_all, [['4s.xlsx', '基本信息_jibenxinxi_4578'], ['307.xlsx', '基本信息-基本信息']], 'para_13_22')
            wr.write_excel(head, bl, fufa, date19, dfs, visit)
            # wr.last('para13_22')
        if False:
            # 参数3——6， '初诊颅外病灶',：luwai
            # 参数18：辅助AI内分泌治疗, cure_out
            cure_out = {}
            self.ai_cure('辅助治疗', '辅助内分泌治疗_adjuec_BC_4637', 'AI类药物', cure_out)

            luwai, new_cure_ai, cure_ai, recrudesce = self.p3_6()
            head = ['pid', '新辅助AI内分泌治疗', '辅助AI内分泌治疗', '复发转移后AI内分泌治疗', '辅助AI内分泌治疗方案', '初诊颅外病灶']
            wr = writeExcel.WriteExcel(self.data_all, [['4s.xlsx', '基本信息_jibenxinxi_4578'], ['307.xlsx', '基本信息-基本信息']], 'AI内分泌')
            wr.write_excel(head,  new_cure_ai, cure_ai, recrudesce, cure_out, luwai)
            wr.last('综合表')
        if False:
            # 参数32：AI治疗失败后后续内分泌治疗总线数-1, line_32_last
            # 参数34：AI治疗失败后后续内分泌治疗方案-1,case_34
            # 参数35：AI治疗失败后后续内分泌治疗开始用药日期-，start_medi
            # 参数36： AI治疗失败后后续内分泌治疗结束用药日期-1，end_medi
            # 参数37：AI治疗失败后后续内分泌治疗结束治疗原因-1, TT
            # 参数38：AI治疗失败后后续内分泌治疗疗效评估-1, evaluate

            line_32_last, case_34, start_medi, end_medi, TT, evaluate = self.para32_38()
            head = ['pid', 'AI失败后内分泌总线数-1',  'AI失败后内分泌分线数方案-1', 'AI失败后内分泌开始用药日期-1', 'AI失败后内分泌结束用药日期-1', 'AI失败后内分泌结束治疗原因-1', 'AI失败后内分泌疗效评估-1']
            wr = writeExcel.WriteExcel(self.data_all, [['4s.xlsx', '基本信息_jibenxinxi_4578'], ['307.xlsx', '基本信息-基本信息']], 'AI治疗失败后后续内分泌治疗')
            wr.write_excel(head, line_32_last, case_34, start_medi, end_medi, TT, evaluate)
            wr.last('AI治疗失败后后续内分泌治疗32_38')
        if False:
            # 参数24：晚期解救治疗初次AI治疗内分泌线数，para_24_line
            # 参数25：新辅助AI内分泌治疗开始日期，para_25_start
            # 参数26：新辅助AI内分泌治疗结束日期，para_26_end
            # 参数27：辅助AI内分泌治疗开始日期，para_27_start
            # 参数28：辅助AI内分泌治疗结束日期，para_28_end

            para_24_start = {}  # 日期
            para_sets = {}
            # 参数24的结果
            para_24_line = {}

            # 参数24的筛选
            condi_24 = lambda x: x not in ['', '新辅助治疗', '辅助治疗']
            self.p307(para_24_start, para_24_line, condi_24)

            # 参数24，4s
            para_date = {}
            for line in self.data_all['4s.xlsx']['复发转移_recum_BC_36112']:
                pid = line['pid']
                if pid == '' or pid == 'pid':
                    continue
                txt = line['治疗方案（文本型）']
                if txt != '':
                    start_date = line['用药开始日期']
                    if re.findall(self.re_little, str(txt)):
                        # 4s
                        if start_date not in ['', '1970-1-1', '1970-01-01']:
                            # 满足第一条日期(1)
                            if pid not in para_date.keys():
                                para_date[pid] = start_date
                            # （2）
                            if re.findall(self.re_many, str(txt)):
                                self.add_set(pid, para_sets, start_date)
            # 找到位置
            self.find_index(para_date, para_sets, para_24_line)

            # 参数25,26的筛选，结束用药日期对应的
            para_25_start = {}
            para_26_end = {}
            condi_25 = lambda x: x == '新辅助治疗'
            self.p307(para_25_start, para_26_end, condi_25, '结束用药日期')
            self.p4s('术前内分泌治疗_preo_ecthero_28594', '治疗开始日期', '治疗结束日期', para_25_start, para_26_end)

            # 参数27，28,结束用药日期最早的
            para_27_start = {}
            para_28_end = {}
            condi_27 = lambda x: x == '辅助治疗'
            self.p307(para_27_start, para_28_end, condi_27, False, '结束用药日期')
            self.p4s('辅助内分泌治疗_adjuec_BC_4637', 'AI类开始日期', 'AI类结束日期', para_27_start, para_28_end)

            head = ['pid', '晚期解救治疗初次AI治疗内分泌线数',  '新辅助AI内分泌治疗开始日期', '新辅助AI内分泌治疗结束日期', '辅助AI内分泌治疗开始日期', '辅助AI内分泌治疗结束日期']
            wr = writeExcel.WriteExcel(self.data_all, [['4s.xlsx', '基本信息_jibenxinxi_4578'], ['307.xlsx', '基本信息-基本信息']], 'AI治疗失败后后续内分泌治疗')
            wr.write_excel(head, para_24_line, para_25_start, para_26_end, para_27_start, para_28_end)
            wr.last('AI治疗失败后后续内分泌治疗24_28')