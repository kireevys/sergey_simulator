import logging
from pathlib import Path

from app.config import Config
from app.core import Order, Storage
from app.parsers import EmailParser
from app.storages import ExcelStorage

logger = logging.getLogger()


def process(eml_path: Path, storage: Storage) -> Order:
    order = EmailParser().parse(eml_path)
    storage.add_order(order)
    storage.add_attachment(order, eml_path)

    return order


def run(config: str, email: str):
    config = Config(Path(config))
    storage = ExcelStorage(config)
    order = process(Path(email), storage)

    logger.info(f"{order.order_id} processed")


def bulk(config: str, emails: str):
    dir = Path(emails)
    for email in dir.rglob("*.eml"):
        run(config, str(email))
