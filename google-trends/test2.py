import re

my_str = '2004-01-04 - 2004-01-10,51'

reg_week = re.compile('^(\d{4}-\d{2}-\d{2}) - (\d{4}-\d{2}-\d{2}),(\d+)')
result = reg_week.match(my_str)
print result.group(3)

my_str2 = '2006-10,62'

reg_month = re.compile('^(\d{4}-\d{2}),(\d+)')
result = reg_month.match(my_str2)
print result.group(2)

print 3 % 3
