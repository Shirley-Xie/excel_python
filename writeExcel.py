import xlwt


class WriteExcel:
    wb = xlwt.Workbook()
    # col = 0

    def __init__(self, data_all, exls, sheetname):
        self.data_all = data_all
        self.ws_data = self.wb.add_sheet(sheetname)
        self.row = 0
        self.sheetname = sheetname
        self.ALL_PIDS = self.get_pids(exls)
        self.col = 0
    # def __init__(self, data_all, sheetname, pids):
    #     self.data_all = data_all
    #     self.ws_data = self.wb.add_sheet(sheetname)
    #     self.row = 0
    #     self.sheetname = sheetname
    #     self.ALL_PIDS = pids
    #     self.col = 0

    def get_pids(self, exl_sheet):
        pids = set()
        for sheet in exl_sheet:
            for line in self.data_all[sheet[0]][sheet[1]]:
                pid = line['pid']
                if pid == '' or pid == 'pid':
                    continue
                pids.add(pid)
        return pids

    # 针对多个简单字典的操作
    def write_excel(self, head, *datas):
        # 头部写好
        if self.row == 0:
            for k, v in enumerate(head):
                self.ws_data.write(self.row, k, v)

        for pid in self.ALL_PIDS:
            self.row += 1
            self.ws_data.write(self.row, 0, pid)
            for k, data in enumerate(datas):
                self.col += 1
                self.one_row(data, self.row, self.col, pid)
            self.col = 0

    # 写入一行,字典的数据类型还有dict，多数组，因为存在空值，必须找到list最长的值
    def one_row(self, data, row, col, pid):
        if pid in data:
            if isinstance(data, set):
                self.ws_data.write(row, col, '是')
            elif isinstance(data, dict):
                if isinstance(data[pid], list):
                    for i in range(col, len(data[pid])+col):
                        self.ws_data.write(row, i, data[pid][i-col])
                    self.col += (len(data[pid])-1)
                elif isinstance(data[pid], dict):
                    pass
                else:
                    self.ws_data.write(row, col, data[pid])

        else:
            self.ws_data.write(row, col, '')

    # 针对字典，值为数组
    def write_arr(self, head, datas):
        # 头部写好
        if self.row == 0:
            for k, v in enumerate(head):
                self.ws_data.write(self.row, k+self.col, v)
        for pid in self.ALL_PIDS:
            self.row += 1
            if self.col == 0:
                self.ws_data.write(self.row, 0, pid)
            if pid in datas.keys():
                self.row_arr(datas[pid], self.row, head)
        self.col += len(datas)

    def row_arr(self, v_arr, row, head):
        if isinstance(v_arr, list):
            for i in range(self.col+1, len(v_arr)+1):
                self.ws_data.write(row, i, v_arr[i-1])
        elif isinstance(v_arr, dict):
            for i, h in enumerate(head):
                if h in v_arr:
                    self.ws_data.write(row, self.col+i, v_arr[h])
                else:
                    self.ws_data.write(row, self.col + i, '')
        else:
            print('格式问题')

    def last(self, name):
        WriteExcel.wb.save('data_out/' + name + '.xls')



