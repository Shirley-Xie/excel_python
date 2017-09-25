from datetime import datetime
from functools import reduce
li = ['', '2010-01-01', '1971-01-01', '1971-11-01', '  ']

# d = list(filter(lambda s: s and s.strip(), li))

def parse_ymd(s):
    year_s, mon_s, day_s = s.split('-')
    return datetime(int(year_s), int(mon_s), int(day_s))


def compared_early(x, y):
    if type(x) == str:
        x = parse_ymd(x)
    if type(y) == str:
        y = parse_ymd(y)
    if x <= y:
        return x
    else:
        return y


d = list(filter(lambda s: s and s.strip(), li))
all = reduce(compared_early, d)
print(all)
print('dd')

bss = parse_ymd('2018-02-12')
ass = parse_ymd('2017-02-12')

text = '2012-09-20'
y = datetime.strptime(text, '%Y-%m-%d')
z = datetime.now()
diff = z - y
print(diff)