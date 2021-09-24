import asyncio
import logging
import time
from asyncio import get_event_loop
from pathlib import Path

from app.config import Config
from app.core import Order, Parser, Storage
from app.parsers import EmailParser
from app.storages import ExcelStorage

logger = logging.getLogger()


def handle_exception(loop, context):
    logger.error("handled")


async def get_order(eml_path: Path, parser: Parser) -> (Order, Path):
    try:
        order = parser.parse(eml_path)
    except Exception:
        logger.error(f"{eml_path} parse error")
        raise

    logger.info(f"Parsed: {order.order_id}")
    return order, eml_path


def add_to_storage(order: Order, eml_path: Path, storage: Storage) -> bool:
    if not storage.is_exists(order):
        storage.add_order(order)
        storage.add_attachment(order, eml_path)
        logger.info(f"Added: {order.order_id}")
        return True
    else:
        logger.warning(f"{order.order_id} already exists")
        return False


def process(eml_path: Path, storage: Storage) -> Order:
    order = EmailParser().parse(eml_path)
    add_to_storage(order, eml_path, storage)

    return order


def run(config: str, email: str):
    config = Config(Path(config))
    storage = ExcelStorage(config)

    process(Path(email), storage)


def bulk(config: str, emails: str):
    dir = Path(emails)
    config = Config(Path(config))
    storage = ExcelStorage(config)
    parser = EmailParser()

    loop = get_event_loop()
    loop.set_exception_handler(handle_exception)

    tasks = [loop.create_task(get_order(email, parser)) for email in dir.rglob("*.eml")]

    total = len(tasks)
    logger.info(f"Start by {total} emails")

    start = time.time()
    results = loop.run_until_complete(asyncio.gather(*tasks, return_exceptions=True))

    logger.info(f"Parsed on {round(time.time() - start, 2)} seconds")

    logger.info(f"Writing to storage")
    start = time.time()
    for n, i in enumerate(results):

        if isinstance(i, Exception):
            continue

        order, eml_path = i

        add_to_storage(order, eml_path, storage)
        logger.info(f"Progress: {n:>5}/{total}")

    logger.info(f"Wrote on {round(time.time() - start, 2)} seconds")
