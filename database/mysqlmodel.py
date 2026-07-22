from sqlmodel import create_engine, SQLModel
engine = create_engine("sqlite:///../wasted.db")

from sqlmodel import Session

def create_session():
    return Session(engine)

#for fastapi
def get_session():
    with Session(engine) as session:
        yield session
