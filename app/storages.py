import logging
import shutil
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from app.config import Config, Template
from app.core import Order, Storage

logger = logging.getLogger()


class ExcelStorage(Storage):
    def __init__(self, config: Config):
        self.config = config

    def _get_cwd(self, order: Order) -> Path:
        return Path(self.config.storage["ROOT"], str(order.date.year))

    def _get_wb_path(self, order: Order) -> Path:
        path = self._get_cwd(order) / f"{order.date.year}_учет_заявок.xlsx"
        if not path.is_file():
            path.parent.mkdir(exist_ok=True, parents=True)

        return path

    def _get_sheet_name(self, order: Order):
        return str(order.date.month)

    def _create_by_template(self, order: Order):
        path = self._get_wb_path(order)
        template = Template(self.config).get_template_by_name("orders_book")
        shutil.copy(template, path)

    def _get_wb_by_order(self, order: Order) -> Workbook:
        path = self._get_wb_path(order)

        if not path.is_file():
            self._create_by_template(order)

        return load_workbook(filename=path)

    def _get_ws_by_order(self, order: Order, wb: Workbook) -> Worksheet:
        name = self._get_sheet_name(order)
        try:
            ws = wb.get_sheet_by_name(name)
        except KeyError:
            tmpl = wb.get_sheet_by_name("default")
            ws = wb.copy_worksheet(tmpl)
            ws.title = name

        return ws

    def is_exists(self, order: Order) -> bool:
        return Path(
            self._get_cwd(order), order.date.strftime("%m"), str(order.order_id)
        ).is_dir()

    def add_order(self, order: Order):
        wb = self._get_wb_by_order(order)
        ws = self._get_ws_by_order(order, wb)
        order_as_row = [
            order.order_id,
            order.date.strftime("%d.%m.%Y"),
            order.warehouse_id,
            None,
            None,
            order.description,
        ]
        ws.append(order_as_row)
        logger.info(f"{order.order_id} added")
        wb.save(self._get_wb_path(order))

    def add_attachment(self, order: Order, attachment: Path):
        path = self._get_cwd(order)
        cwd = Path(path, order.date.strftime("%m"), str(order.order_id))
        cwd.mkdir(parents=True, exist_ok=True)
        shutil.copy(attachment, cwd)
