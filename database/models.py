from decimal import Decimal
from sqlmodel import SQLModel, Field, Column, Numeric
from enum import Enum

class OrderStatus(str, Enum):
    CANCELED = "canceled"
    UPDATED = "updated"
    WROTEOFF = "wroteoff"
    GUILTED = "guilted"

class OrderWasteAccountBook(SQLModel, table=True):
    __tablename__ = "orders_wasteaccountbook"

    id: int | None = Field(default=None, primary_key=True)
    date: str
    ordernum: int | None = None
    totalsum: Decimal | None = None
    eventyear: int | None = None
    status: OrderStatus | None = None
    status_description: str | None = None

class TitleWastebook(SQLModel, table=True):
    __tablename__ = "titles_wastebook"

    id: int | None = Field(default=None, primary_key=True)
    title: str
    order_id: int | None = None
    total_for_items: Decimal | None = Field(
        default=None,
        sa_column=Column(Numeric)
    )
    service_str: int | None = None


class Test2(SQLModel, table=True):
    __tablename__ = "test2"
    id: int | None = Field(default=None, primary_key=True)
    title: str
