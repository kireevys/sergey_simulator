from datetime import datetime
from pathlib import Path

from app.core import Order, OrderStatus
from app.parsers import EmailParser


class TestEmailParser:
    extract_datafile = Path("data/content.html")

    def test_parse_order(self, root):
        order = EmailParser().parse(Path("data/email.eml"))

        expected_order = Order(
            order_id=12747295,
            warehouse_id=4734,
            description="ПРИКЛЕИТЬ СТИКЕР НА ВИТРИНУ",
            date=datetime.strptime("08.09.2021", "%d.%m.%Y"),
            status=OrderStatus.NEW,
        )

        assert order == expected_order

    def test_extract_order_id(self):
        page = self.extract_datafile.read_text()

        order_id = EmailParser()._extract_order_id(page)

        assert order_id == 12747295

    def test_extract_order_date(self):
        page = self.extract_datafile.read_text()

        date = EmailParser()._extract_order_date(page)

        assert date == datetime(2021, 9, 8)

    def test_extract_description(self):
        page = self.extract_datafile.read_text()

        description = EmailParser()._extract_description(page)

        assert description == "ПРИКЛЕИТЬ СТИКЕР НА ВИТРИНУ"

    def test_extract_wirehouse_id(self):
        page = self.extract_datafile.read_text()

        wirehouse_id = EmailParser()._extract_warehouse_id(page)

        assert wirehouse_id == 4734
