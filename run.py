import json
import re
import difflib
import xlrd
import xlwt
from functools import reduce
import time
from datetime import datetime, timedelta
from itertools import compress
import functools


def _map():
    def multiply(x):
        return (x*x)

    def add(x):
        return (x+x)

    funcs = [multiply, add]
    for i in range(5):
        # 输出函数写x(i)
        value = list(map(lambda x: x(i), funcs))
        print(value)


def parse_ymd(s):
    year_s, mon_s, day_s = s.split('-')
    return datetime(int(year_s), int(mon_s), int(day_s))


# 此函数只要一条语句不同，是否可以写为装饰器
def demo_during(is_pfs_307, start_map, whether='yes'):
    during_time = {}

    def yes_or_not():
        if whether == 'yes':
            return k in is_pfs_307
        else:
            return k not in is_pfs_307

    for k, start_v in start_map.items():
        # 是的情况下，且开始和结束值都存在
        if yes_or_not():
            during_time[k] = start_v

    return during_time


# 关于计算问题
def demo_count():
    a = 31
    print(str(a//30) + '.' + str(a % 30))


# 正则匹配
def demo_re():
    # 匹配如下
    con = r"(拉帕替尼|L|lapatinib)(卡培他滨|X)(^赫赛汀|^曲妥珠单抗|^H|\w*)"
    name1 = '拉帕替尼卡培他滨_'
    name2 = 'lapatinibXH'
    refuse = re.findall(con, name1)

    # 1 正则匹配带有'发展'字段
    all_name = ['疾病进展时间[符号]', '进展后选择', '数据库编号']
    for li in all_name:
        kk = re.findall(r'\w?进展.*', li)  # 错的
        print(kk)

    # 2 匹配 '药物通用名-商品名'
    others = ['药物通用名-商品名-2', '药物通用名-商品名', 'qvwbvhwiebvihwbv']
    for li in others:
        kk = re.findall(r'药物通用名-商品名[-]?[\d]?', li)
        print(kk)


def demo_filter():
    # 1 简单使用，推导和生成器
    values = [2, 3, -4, -5]
    clip_neg_1 = [n if n > 0 else 0 for n in values]
    clip_neg = [n for n in values if n > 0]

    # 2 稍复杂使用，filter()过滤非整数的值
    value = ['1', '2', '-3', '-', '4', 'N/A', '5']

    def is_int(val):
        try:
            int(val)
            return True
        except ValueError:
            return False

    ivals = list(filter(is_int, value))
    print(ivals)

    # 3 复杂使用，关联过滤
    addresses = [
        '5412 N CLARK',
        '5148 N CLARK',
        '5800 E 58TH',
        '2122 N CLARK',
        '5645 N RAVENSWOOD',
        '1060 W ADDISON',
        '4801 N BROADWAY',
        '1039 W GRANVILLE',
    ]
    counts = [0, 3, 10, 4, 1, 7, 6, 1]
    # 返回true或false，神奇的推导式
    more5 = [n > 5 for n in counts]
    # compress 关联过滤
    ads = list(compress(addresses, more5))
    print(ads)


# args's type is tuple, kwargs's type is dict
def comprehensions_lambda_(*args, **kwargs):
    # 传进来的参数看实参逗号的个数
    print(args, kwargs)
    '''
    lambda：只使用一次，一行函数
    lambda x:x **2  传入值:返回值
    推导式 for之前为返回值
    ====
    
    推导式省了新建dict，最前面是返回的结果
    dict, keys可以根据k获取v，items若只写一项k可用k[0],k[1]
    将同一字母对应的数值相加
    '''
    d = {'2': 21, '4': 41, '5': 51}
    items = sorted(d.items(), key=lambda c: c[1])
    print(items)

    mcase = {'a': 10, 'b': 34, 'A': 7, 'Z': 3}
    # 若key存在则忽视，不存在则将参数2赋值，不更改原来的字典
    ggg = mcase.get('a', "Never")

    # 写法1
    mcase_frequency_1 = {}
    for k in mcase.keys():
        mcase_frequency_1[k.lower()] = mcase.get(k.lower(), 0) + mcase.get(k.upper(), 0)

    # 写法2，推导式, 可以遍历dict的keys，使用get方法获取values,互换例子{v: k for k, v in some_dict.items()}
    mcase_frequency_2 = {k.lower(): mcase.get(k.lower(), 0) + mcase.get(k.upper(), 0) for k in mcase.keys()}

    # 写法三
    # mcase_frequency_2 = dict(map(lambda k: mcase.get(k.lower()) + mcase.get(k.upper()), mcase.keys()))
    '''
    set
    和dict很像，区别是返回值的类型不一样{k:v}, {k1, k2}
    '''
    squared_set = {x ** 2 for x in [1, 1, 2]}
    print(squared_set)

    '''
    list
    实现数组的平方
    '''
    items = [1, 2, 3, 4, 5]
    # 写法1，普通
    squared_1 = []
    for ii in items:
        squared_1.append(ii ** 2)
    # 写法2，map
    squared_2 = list(map(lambda x: x ** 2, items))
    # 写法3，推倒式, [out_exp for out_exp in input_list if out_exp == 2]
    squared_3 = [x**2 for x in items]
    # args 看成整体（{},）
    # 只对key起作用
    _map_ = list(map(str, {1: 323, 2: 33, 3: 55}))
    keyvals = [' %s="%s"' % item for item in kwargs.items()]
    dd = kwargs.pop('silent', False)
    # print('argument "{0}"'.format(list(kwargs.keys())[0]))
    # print('demo_comprehensions')
def same_radio():
    seq = difflib.SequenceMatcher(None, ' 汉字'.lower(), '汗子'.lower())
    ratio = seq.ratio()
    reject = re.findall(r"^.9", '[9')

# 代码开始的地方
if __name__ == '__main__':
    ba = int('12345')  # 默认为10
    ba = int('12345', base=8)
    shilu = int('12345', 16)

    # 现在大量使用二进制
    def int2(x, base=2):
        return int(x, base)
    int('12345')
    # 替代上面
    int2 = functools.partial(int, base=2)
    # excel 开始处理
    # start = excel.ExcelHandle()
    # d1 = parse_ymd('2017-08-01')
    # d2 = parse_ymd('2017-08-31')
    #  匹配度
    row = 0
    # global row

    def exited(a, b):
        b += 1
        print("a = %s,b=%s" % (a, b))
        # print("#define %s_%s \n" % (a, b))
    # 偏函数和默认参数区别是作用域
    exited_ = functools.partial(exited, row)
    exited_(1)
    exited_(2)
    exited_(3)

    def parse_ymd(s):
        year_s, mon_s, day_s = s.split('-')
        return datetime(int(year_s), int(mon_s), int(day_s))

    def get_next_day(base_day, n):
        return str((datetime.strptime(str(base_day), '%Y-%m-%d') + timedelta(days=n)).date())

    # d = datetime.strptime(str('2016/3/20'), '%Y-%m-%d')

    date_time = ''
    date_ = 'j'
    if date_time in ['1970-01-01', ''] or date_ in ('', 'pid'):
        print(date_time, 'df')
    ad = get_next_day('2013-12-12', 19)

    prefix = 'sdf ssdf'
    table = [0] * len(prefix)
    # 检测和第一个重复的其余
    for i in range(1, len(prefix)):
        idx = table[i - 1]
        while prefix[i] != prefix[idx]:
            if idx == 0:
                table[i] = 0
                break
            idx = table[idx - 1]
        else:
            table[i] = idx + 1
    print(table)
    # prefix的前一个和
    # 用户留存率,1天，3天，7天

    demo_filter()
    comprehensions_lambda_(23, 33)
    a = [(1, 2), (88, 1), (9, 10), (13, -3)]
    a.sort(key=lambda x: x[1])

    # Example
    # Creates '<item size="large" quantity="6">Albatross</item>'
    # demo_comprehensions(fool, len=dkk)
    prefix = []
    prefixappend = prefix.append
    for i in range(10):
        prefixappend(i)
    print(prefix)
    # not_match = demo_during({1, 2, 3}, {1: 11, 2: 22, 4: 33}, 'no')
    # print(not_match)

    from dateutil import rrule



