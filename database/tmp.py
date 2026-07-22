from models import Test2
from Repository import Repository
from sqlmodel import Session
from mysqlmodel import engine

with Session(engine) as session:
    repo = Repository(Test2, session)
    t = repo.create(Test2(title='gfdgdfg'))
    t = repo.get(1)
    print(t.title)