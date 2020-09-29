import re
value = '......'
lenth = len(value.split('.'))-2
print(lenth)

floatFormat = '[^\d|\d.\d]'
value = re.sub(floatFormat, '', value)
print(value)
value = value.replace('.', '', lenth)
if value == '' or value == '.':
    value = 0
value = float(value)

print(value)