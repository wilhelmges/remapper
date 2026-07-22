from sqlmodel import Session; from mysqlmodel import engine
from database.models import *

if __name__ == "__main__":
    with Session(engine) as session:
        print(SQLModel.metadata.tables.keys())
        print('creating tables')
        SQLModel.metadata.create_all(engine)
        testn = Test23(title="test")
        session.add(testn)
        session.commit()
