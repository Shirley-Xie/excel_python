# 在当前路径data下的文件filename
path_dir = './data'
filename = os.walk(path_dir).__next__()[2]
path = path_dir + '/' + filename
workbook = xlrd.open_workbook(path)

sheet_name = workbook.sheet_names()
sheet = workbook.sheet_by_name(sheet_name)

n_rows = sheet.nrows
n_cols = sheet.ncols

# 建立表名
headers = {}

for row_index in range(nrows):
    line = {}
    for col_index in range(ncols):
        try: 
        	if row_index == 0:
        		value = sheet.cell_value(0, col_index)
        		headers[row_index] = value
            else:		
            	value = sheet.cell_value(row_index, col_index)
        except Exception as e:
            print('filename:', filename, ' sheet_name:', sheet_name, ' row_index:', row_index,
                  ' col_index:', col_index)
        header = headers[col_index]
        line[header] = value
    data_all[filename][sheet_name].append(line)
