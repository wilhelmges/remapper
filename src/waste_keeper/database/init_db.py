from sqlmodel import SQLModel

from waste_keeper.database.db import engine
from waste_keeper.database.wasted import Wasted
from waste_keeper.database.models import OrderWasteAccountBook, TitleWastebook
from waste_keeper.database.property_loss_order import Property_loss_order

def remove_tables():
    SQLModel.metadata.drop_all(engine)

def clear_tables():
    SQLModel.metadata.drop_all(engine)
    SQLModel.metadata.create_all(engine)

if __name__=="__main__":
    SQLModel.metadata.drop_all(engine)
    SQLModel.metadata.create_all(engine)

    #print(SQLModel.metadata.tables.keys())

    # if __name__ == "__main__":
    #     with Session(engine) as session:

    #         print('creating tables')
    #         SQLModel.metadata.create_all(engine)
    #         testn = Test23(title="test")
    #         session.add(testn)
    #         session.commit()