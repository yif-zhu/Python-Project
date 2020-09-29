import re
from decimal import Decimal

t4 = 1.2

t2= 100.02

cvalue = '2025/5/3！'

reg = r"(\d{4}[-/]\d{1,2}([-/]\d{1,2})?)|((\d{1,2}[-/])?\d{1,2}[-/]\d{4})|(\d{4}年\d{1,2}月(\d{1,2}日)?)"
dtcheck = 1 if re.search(reg, cvalue) is not None else 0

print(dtcheck)
