import json
import shelve
from datetime import datetime
from pathlib import Path

from app.core import Order, OrderStatus
from app.index import FileIndex
from tests.test_core.conftest import config


class MockWs:
    c = 0

    @property
    def max_row(self):
        yield self.c
        self.c += 1


def test_add_index(root):
    row = 2
    order = Order(
        order_id=1,
        warehouse_id=4734,
        description="ПРИКЛЕИТЬ СТИКЕР НА ВИТРИНУ",
        date=datetime.strptime("08.09.2021", "%d.%m.%Y"),
        status=OrderStatus.NEW,
    )
    result = FileIndex(config).add(order, row)

    expected_value = f"{root}/{order.date.year}/{order.date.year}_учет_заявок.xlsx:{order.date.month}:{row}"
    assert result == expected_value

    index_path = Path(root, "index")

    with shelve.open(str(index_path)) as db:
        assert db["1"] == expected_value


def test_get(root):
    order = Order(
        order_id=1,
        warehouse_id=4734,
        description="ПРИКЛЕИТЬ СТИКЕР НА ВИТРИНУ",
        date=datetime.strptime("08.09.2021", "%d.%m.%Y"),
        status=OrderStatus.NEW,
    )
    path_value = f"{root}/{order.date.year}/{order.date.year}_учет_заявок.xlsx:{order.date.month}:2"
    index_path = Path(root, "index")

    index_path.parent.mkdir(parents=True, exist_ok=True)

    with shelve.open(str(index_path)) as f:
        f["1"] = path_value

    index = FileIndex(config)

    result = index.get("1")

    assert result == path_value


def test_add_and_get(root):
    orders = [
        Order(
            order_id=1,
            warehouse_id=4734,
            description="ПРИКЛЕИТЬ СТИКЕР НА ВИТРИНУ",
            date=datetime.strptime("08.09.2021", "%d.%m.%Y"),
            status=OrderStatus.NEW,
        ),
        Order(
            order_id=2,
            warehouse_id=4734,
            description="ПРИКЛЕИТЬ СТИКЕР НА ВИТРИНУ",
            date=datetime.strptime("08.09.2021", "%d.%m.%Y"),
            status=OrderStatus.NEW,
        ),
    ]
    index = FileIndex(config)
    for row, order in enumerate(orders):
        index.add(order, row)

    for row, order in enumerate(orders):
        assert (
            index.get(str(order.order_id))
            == f"{root}/{order.date.year}/{order.date.year}_учет_заявок.xlsx:{order.date.month}:{row}"
        )
