from datetime import datetime

from waste_keeper.utils.core import value_to_pydate, value_to_sqlite_date
s = '30/3/26'
#print(datetime.strptime(s, '%d/%m/%y').date().isoformat())
print(value_to_sqlite_date(s))


# def year_servnumb_forrao(row):
#     ch,zal = divmod(row-10, 3)
#     year = 2022 + zal
#     return year, ch+1
#
# for row in range(10,19):
#     print(year_servnumb_forrao(row))

# print(year_servnumb_forrao(10))