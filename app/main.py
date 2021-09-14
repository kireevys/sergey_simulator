import logging
from pathlib import Path

from app.config import Config
from app.core import Order, Storage
from app.parsers import EmailParser
from app.storages import ExcelStorage

logger = logging.getLogger()


def process(eml_path: Path, storage: Storage) -> Order:
    order = EmailParser().parse(eml_path)
    if not storage.is_exists(order):
        storage.add_order(order)
        storage.add_attachment(order, eml_path)
    else:
        logger.warning(f"{order.order_id} already exists")

    return order


def run(config: str, email: str):
    config = Config(Path(config))
    storage = ExcelStorage(config)

    process(Path(email), storage)


def bulk(config: str, emails: str):
    dir = Path(emails)
    for email in dir.rglob("*.eml"):
        try:
            run(config, str(email))
        except Exception:
            logger.warning(f"{email} not processed")
