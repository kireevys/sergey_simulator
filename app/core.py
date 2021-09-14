from abc import ABC, abstractmethod
from datetime import datetime
from enum import Enum
from pathlib import Path

from pydantic import BaseModel


class OrderStatus(Enum):
    NEW = "Новая"
    ACCEPTED = "Принята"
    PROCESS = "В работе"
    WAITING = "Ожидание закрытия"
    CLOSED = "Закрыта"


class Order(BaseModel):
    order_id: int
    warehouse_id: int
    description: str
    status: OrderStatus
    date: datetime


class Parser(ABC):
    @abstractmethod
    def parse(self, path: Path) -> Order:  # pragma: no cover
        ...


class Item(ABC):
    @abstractmethod
    def get_item(self):  # pragma: no cover
        ...


class Storage(ABC):
    @abstractmethod
    def add_order(self, order: Order):  # pragma: no cover
        ...

    @abstractmethod
    def add_attachment(self, order: Order, attachment: Path):
        ...


class Builder(ABC):
    @abstractmethod
    def build(self, order: Order) -> Item:  # pragma: no cover
        ...
