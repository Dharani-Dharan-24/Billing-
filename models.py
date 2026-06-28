from datetime import datetime
from typing import Optional

from sqlalchemy import DateTime, Float, ForeignKey, Integer, String
from sqlalchemy.orm import Mapped, mapped_column, relationship

from database import Base


class Customer(Base):
    __tablename__ = "customers"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True, index=True)
    name: Mapped[str] = mapped_column(String, unique=True, nullable=False, index=True)
    phone: Mapped[Optional[str]] = mapped_column(String, nullable=True)
    location: Mapped[Optional[str]] = mapped_column(String, nullable=True)
    customer_type: Mapped[str] = mapped_column(String, nullable=False)

    bills: Mapped[list["Bill"]] = relationship(back_populates="customer")


class Product(Base):
    __tablename__ = "products"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    name: Mapped[str] = mapped_column(String, unique=True, nullable=False, index=True)
    default_price: Mapped[float] = mapped_column(Float, nullable=False)


class Bill(Base):
    __tablename__ = "bills"

    bill_id: Mapped[str] = mapped_column(String, primary_key=True)
    customer_id: Mapped[int] = mapped_column(ForeignKey("customers.id"), nullable=False)
    subtotal: Mapped[float] = mapped_column(Float, nullable=False)
    packaging_fee: Mapped[float] = mapped_column(Float, nullable=False)
    grand_total: Mapped[float] = mapped_column(Float, nullable=False)
    advance_paid: Mapped[float] = mapped_column(Float, nullable=False)
    balance_due: Mapped[float] = mapped_column(Float, nullable=False)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow, nullable=False)

    customer: Mapped[Customer] = relationship(back_populates="bills")
    items: Mapped[list["BillItem"]] = relationship(
        "BillItem",
        backref="bill",
        cascade="all, delete-orphan",
    )


class BillItem(Base):
    __tablename__ = "bill_items"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    bill_id: Mapped[str] = mapped_column(ForeignKey("bills.bill_id"), nullable=False, index=True)
    product_name: Mapped[str] = mapped_column(String, nullable=False)
    quantity: Mapped[int] = mapped_column(Integer, nullable=False)
    price: Mapped[float] = mapped_column(Float, nullable=False)
    total: Mapped[float] = mapped_column(Float, nullable=False)
