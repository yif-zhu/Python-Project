import pandas as pd
import re,os
import calendar

def getDate(value):
    model1 = r"(\d{4}年\d{1,2}月(\d{1,2}日)?)"
    model2 = r"(\d{4}[-/]\d{1,2}[^a-z]([-/]\d{1,2}[^a-z])?)|((\d{1,2}[^a-z][-/])?\d{1,2}[^a-z][-/]\d{4})"
    model3 = r"^(\d{5,6})$"
    #value = '2016年3月31日'
    if re.search(model1, value) is not None:
        value = re.sub('\D$', '', value)
        value = re.sub(r'\D', r'-', value)
        value = pd.to_datetime(value)
    elif re.search(model3, value) is not None:
        if re.search(model2, value) is not None:
            value = pd.to_datetime(value)
        else:
            value = pd.to_datetime(value, format=('%Y%m'))
    elif value == '':
        return ''
    else:
        value = pd.to_datetime(value)
    year, month = str(value).split('-')[0], str(value).split('-')[1]  # 获取月末日期
    end = calendar.monthrange(int(year), int(month))[1]
    value = pd.to_datetime('%s-%s-%s' % (year, month, end))
    return value.strftime('%Y%m%d')

# print(getDate('2012年12月'))
# print(getDate('200807'))
# print(getDate('2016/11/30'))
print(getDate('2013-01-01 00:00:00'))
# print(getDate('2013-1'))



# temp = {'money': 1000,'fee': 900}
# print(temp['money'])
#
# num = "insert into PortfolioManagemnt.DvImport.StaticPoolData("
# num = num + 'FileNames,'
# num = num + 'LoanDate,'
# print(num.rstrip(',')+') Values (')


def getNumber(Value,Type):
    cvalue = str(Value).replace(' ', '').replace(',', '').replace('.00', '').replace('\t', '').replace('\n', '')
    if cvalue == 'NA' or cvalue == '-' or cvalue == '':
        return 0
    if Type == 'Int':
        return int(cvalue)
    elif Type == 'Float':
        return '%.2f' %float(cvalue)

# print(getNumber('197,991.40', 'Float'))
# print(getNumber('36,101', 'Int'))
# print(getNumber('-', 'Float'))
# print(getNumber('0.10', 'Float'))
#
# filePath = r'C:/Users/DELL/Desktop/关于未来两年“兴晴”系列个人消费类贷款资产支持证券注册申请报告.pdf'
# fileName = os.path.basename(filePath)
# print(fileName)
# columns = '静态池ID,静态池名'
# for c in columns.split(','):
#     print(c)



str = 'yif'
print(str[0].upper() + str[1:].lower())

from functools import reduce
def prod(L):
    reduce(lambda x, y: x*y, L)


def str2float(s):
    n = s.find(".") + 1
    l = len(s)
    s = s.replace(".", "")
    f = reduce(lambda x, y: x * 10 + y, map(int, s))
    f = f / (10 ** (l - n))
    return f



print('str2float(\'123.456\') =', str2float('123.456'))
if abs(str2float('123.456') - 123.456) < 0.00001:
    print('测试成功!')
else:
    print('测试失败!')
