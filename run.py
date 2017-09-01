import json
import re
import os
import xlrd
import xlwt
"""
找出成绩占全班10%且每门超80分，
有参加活动或者特长生找出来
"""


# 转换格式
def init_data():
    global DATA_ALL
    # 录入文件，输出文件
    path_dir = './data'
    out_dir = './data_out'
    data_all = {}
    for filename in os.walk(path_dir).__next__()[2]:
        data_all[filename] = {}
        path = path_dir + '/' + filename
        # 根据路径打开excel
        workbook = xlrd.open_workbook(path)
        # 循环表单，表名
        for sheet_name in workbook.sheet_names():
            data_all[filename][sheet_name] = []
            # if sheet_name == '要查询的表单':
            if True:
                # 根据表名获取表单内容
                sheet = workbook.sheet_by_name(sheet_name)
                nrows = sheet.nrows
                ncols = sheet.ncols
                headers = {}
                for row_index in range(nrows):
                    line = {}
                    for col_index in range(ncols):
                        try:
                            if row_index == 0:
                                value = sheet.cell_value(0, col_index)
                                headers[col_index] = value
                                continue
                            else:
                                value = sheet.cell_value(row_index, col_index)
                        except Exception as e:
                            print('filename:', filename, ' sheet_name:', sheet_name, ' row_index:', row_index,
                                  ' col_index:', col_index)
                        header = headers[col_index]
                        line[header] = value
                    if row_index != 0:
                        data_all[filename][sheet_name].append(line)
        out_file_name = out_dir + '/' + filename
        f = open(out_file_name, 'w')
        f.write(json.dumps(data_all[filename]))
        f.close()
    f = open('data_all', 'w')
    f.write(json.dumps(data_all))
    f.close()
    DATA_ALL = data_all

wb = xlwt.Workbook()
ws_data = wb.add_sheet('data')


def write_data(data_list, name):
    global row
    ws_data.write(row, 0, 'pid')
    ws_data.write(row, 1, name)
    for k, v in data_list.items():
        row += 1
        ws_data.write(row, 0, str(k))
        ws_data.write(row, 1, v)
    row += 1
# 代码开始的地方
if __name__ == '__main__':
    init_data()


