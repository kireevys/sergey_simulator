import logging
import shelve
from pathlib import Path
from typing import Optional

from app.config import Config
from app.core import Order

logger = logging.getLogger()


class FileIndex:
    def __init__(self, config: Config):
        self.config = config
        self._index = {}

    def _get_cwd(self, order: Order) -> Path:
        return Path(self.config.storage["ROOT"], str(order.date.year))

    def _get_wb_path(self, order: Order) -> Path:
        return self._get_cwd(order) / f"{order.date.year}_учет_заявок.xlsx"

    @property
    def db(self) -> shelve.DbfilenameShelf:
        path = Path(self.config.storage["ROOT"], "index")
        path.parent.mkdir(parents=True, exist_ok=True)
        return shelve.open(str(path))

    def add(self, order: Order, row: int) -> str:
        value = f"{self._get_wb_path(order)}:{order.date.month}:{row}"
        self.db[str(order.order_id)] = value
        logger.info(f"{order.order_id} -> {value}")

        return value

    def get(self, order_id: str) -> Optional[str]:
        return self.db[str(order_id)]
