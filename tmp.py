import datetime

d = datetime.datetime.now()
print(type(d), d)
if type(d) == datetime.datetime:
    d = d.date()
print(type(d), d)
