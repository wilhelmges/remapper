from datetime import date

from openpyxl import load_workbook
from sqlmodel import create_engine, SQLModel, Field

from waste_keeper.domain.property_loss_order import Property_loss_order as domain_Property_loss_order


class Property_loss_order(SQLModel, table=True):
    __tablename__ = "property_loss_orders"

    id: int | None = Field(default=None, primary_key=True)
    ordernum: int
    orderdate: date

    canceled: bool
    cancelling_order_num: int|None = None
    cancelling_order_date: date|None = None

    effects_to_order_num: int|None = None

    description: str|None = None

    @classmethod
    def from_excel(cls, ws, num):
        plo: domain_Property_loss_order = domain_Property_loss_order(ws, num)
        if not plo.isorder:
            raise Exception(f'bad order rownum: {num}')
        return cls(
            ordernum=plo.ordernum,
            orderdate=plo.orderdate,
            canceled=plo.canceled,
            cancelling_order_num=plo.canceled_by_ordernum,
            cancelling_order_date=plo.canceled_by_orderdate,
            effects_to_order_num=plo.effects_to_order_num,
            description=plo.description,
        )

if __name__ == '__main__':
    from waste_keeper.config import tmp_property_loss_orders
    wb = load_workbook(tmp_property_loss_orders, data_only=True)
    ws = wb["Sheet1"]

    plo: Property_loss_order = Property_loss_order.from_excel(ws, 11217)
    print(plo.ordernum)
