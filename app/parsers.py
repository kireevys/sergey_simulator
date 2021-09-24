from datetime import datetime
from io import StringIO
from pathlib import Path
from typing import List

import eml_parser
from lxml import etree
from lxml.etree import Element

from app.core import Order, OrderStatus, Parser


class EmailParser(Parser):
    def parse(self, path: Path) -> Order:
        ep = eml_parser.EmlParser(include_raw_body=True)
        parsed_eml = ep.decode_email_bytes(path.read_bytes())

        content = parsed_eml["body"][0]["content"]

        return Order(
            order_id=self._extract_order_id(content),
            warehouse_id=self._extract_warehouse_id(content),
            date=self._extract_order_date(content),
            description=self._extract_description(content),
            status=OrderStatus.NEW,
        )

    def _build_xpath(self, anchor: str) -> str:
        return f'//p[text()="{anchor}"]/../../following-sibling::tr//p[string-length(text())>0]'

    def _get_by_xpath(self, html: str, xpath: str) -> list:
        htmlparser = etree.HTMLParser()
        tree = etree.parse(StringIO(html), htmlparser)
        els: List[Element] = tree.xpath(xpath)
        return els

    def _extract_order_id(self, html: str) -> int:
        xpath = self._build_xpath("Заказ на покупку")

        try:
            el: Element = self._get_by_xpath(html, xpath)[0]
        except IndexError:
            xpath = self._build_xpath("Purchase order")
            el: Element = self._get_by_xpath(html, xpath)[0]

        return int(el.text)

    def _extract_order_date(self, html: str) -> datetime:
        xpath = self._build_xpath("Дата")
        try:
            el: Element = self._get_by_xpath(html, xpath)[0]
        except IndexError:
            xpath = self._build_xpath("Date")
            el: Element = self._get_by_xpath(html, xpath)[0]

        return datetime.strptime(el.text, "%d/%m/%Y")

    def _extract_description(self, html: str) -> str:
        xpath = self._build_xpath("Описание")
        try:
            el: Element = self._get_by_xpath(html, xpath)[0]
        except IndexError:
            xpath = self._build_xpath("Description")
            el: Element = self._get_by_xpath(html, xpath)[0]

        return el.text.strip().upper()

    def _extract_warehouse_id(self, html: str) -> int:
        xpath = self._build_xpath("Место назначения")
        try:
            el: Element = self._get_by_xpath(html, xpath)[0]
        except IndexError:
            xpath = self._build_xpath("Destination")
            el: Element = self._get_by_xpath(html, xpath)[0]

        return int(el.text.split("/")[0].strip())
