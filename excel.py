import xlwt
import xlrd
import os
import json


class ExcelHandle:
    #  读取excel
    # 在当前路径data下的文件filename
    wb = xlwt.Workbook()
    ws_data = wb.add_sheet('data')
    data_all = {}

    # 转换格式
    def init_data(self):
        global data_all
        # 录入文件，输出文件
        path_dir = './data'
        # out_dir = './data_out'
        for filename in os.walk(path_dir).__next__()[2]:
            data_all[filename] = {}
            path = path_dir + '/' + filename
            # 根据路径打开excel
            workbook = xlrd.open_workbook(path)
            # 循环表单，表名
            for sheet_name in workbook.sheet_names():
                data_all[filename][sheet_name] = []
                # 根据表名获取表单内容
                sheet = workbook.sheet_by_name(sheet_name)
                n_rows = sheet.nrows
                n_cols = sheet.ncols
                headers = {}
                # 默认是第一行为字段名
                for row_index in range(n_rows):
                    line = {}
                    for col_index in range(n_cols):
                        try:
                            if row_index == 0:
                                value = sheet.cell_value(0, col_index)
                                headers[col_index] = value
                                continue
                            else:
                                value = sheet.cell_value(row_index, col_index)
                        except Exception as e:
                            print('error:', e, 'filename:', filename, ' sheet_name:', sheet_name,
                                  ' row_index:', row_index, ' col_index:', col_index)
                        header = headers[col_index]
                        line[header] = value
                    if row_index != 0:
                        data_all[filename][sheet_name].append(line)
    #     out_file_name = out_dir + '/' + filename
    #     f = open(out_file_name, 'w')
    #     f.write(json.dumps(data_all[filename]))
    #     f.close()
        f = open('data_all', 'w')
        f.write(json.dumps(data_all))
        f.close()

    # def write_data(data_list, name):
    #     global row
    #     ws_data.write(row, 0, 'pid')
    #     ws_data.write(row, 1, name)
    #     for k, v in data_list.items():
    #         row += 1
    #         ws_data.write(row, 0, str(k))
    #         ws_data.write(row, 1, v)
    #     row += 1
    #
    # # 写入新的excel表中
    # wb = xlwt.Workbook()
    # ws_data = wb.add_sheet('data')
    # wb.save('data_out/cycle_all.xls')
    #
    # def write_data(data_list, name):
    #     global row
    #     ws_data.write(row, 0, 'pid')
    #     ws_data.write(row, 1, name)
    #     for k, v in data_list.items():
    #         row += 1
    #         ws_data.write(row, 0, str(k))
    #         ws_data.write(row, 1, v)
    #     row += 1



