from sqlalchemy.orm import declarative_base
from sqlalchemy.orm import sessionmaker
from sqlalchemy import create_engine
from sqlmodel import Session

from waste_keeper.config import db_path

engine = create_engine(f"sqlite:///{db_path.as_posix()}")
session = Session(engine)
SessionFactory = sessionmaker(
    bind=engine,
    class_=Session,
)
Base = declarative_base()