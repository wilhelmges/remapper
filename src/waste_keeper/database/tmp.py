import re

s ="(зміни в 5602; 318)                       5469"
result = re.sub(r"\([^)]*\)", "", s).strip()

print(result)

# with Session(engine) as session:
#     # repo = Repository(Test2, session)
#     # t = repo.create(Test2(title='gfdgdfg'))
#     # t = repo.get(1)
#     # print(t.title)
#
#     orepo=Repository(OrderWasteAccountBook, session)
#     o = orepo.create(OrderWasteAccountBook(ordernum=234, date='2000-11-11'))
#     print(o.id)
#
#
#     trepo = Repository(TitleWastebook, session)
#     t = trepo.create(TitleWastebook(title='gfdgdfg', orderid = o.id))