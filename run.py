import os
"""
找出成绩占全班10%且每门超80分，
有参加活动或者特长生找出来
"""


# 从excel中获取数据
def get_data():
    data_all = {}
    path_dir = './data'
    for filename in os.walk(path_dir).__next__()[2]:
        data_all[filename] = {}


# 数据处理
def deal_data():
    pass

# 代码开始的地方
if __name__ == '__main__':
    get_data()
