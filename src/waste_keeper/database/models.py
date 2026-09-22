from decimal import Decimal
from typing import Optional

from sqlmodel import SQLModel, Field, Column, Numeric, Relationship
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
    status: str | None = None
    status_description: str | None = None
    strike: bool
    superorder_id: int | None
    unit: str | None
    sheetrow: int

    titles: list["TitleWastebook"] = Relationship(back_populates="order")

class TitleWastebook(SQLModel, table=True):
    __tablename__ = "titles_wastebook"
    id: int | None = Field(default=None, primary_key=True)
    title: str
    order_id: int | None = Field(
        default=None,
        foreign_key="orders_wasteaccountbook.id",
        index=True,
    )
    amount: float
    sheetrow: int
    #service_str: int | None = None

    order: OrderWasteAccountBook | None = Relationship(back_populates="titles")


class Test2(SQLModel, table=True):
    __tablename__ = "test2"
    id: int | None = Field(default=None, primary_key=True)
    title: str
