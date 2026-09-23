from sqlmodel import SQLModel, Field


class Wasted(SQLModel, table=True):
    __tablename__ = "wasted"

    id: int = Field(primary_key=True)
    order_id: int
    order_date: str
    root_order_id: int | None = None
    money: float | None = None
    department: bytes
    row: float | None = None
    root_date: str | None = None
    sheet_name: str | None = None
    filled: int = 0
    impacted: int = 0