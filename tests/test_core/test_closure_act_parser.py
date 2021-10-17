import datetime
from pathlib import Path

import pytest

from app.core import ClosureAct
from app.parsers import EmailCAPParser


@pytest.mark.parametrize(
    "eml,expected",
    [
        (
            Path(  # noqa
                "data/Resolución OTs - Tienda  Bershka 13851 KAZ-MEGA IKEA - Pedido  12175150 _ WOs resolution - St...175150 - Compras Indirectas Inditex (noreply.sficompras@inditex.com) - 2021-09-13 1932.eml"
            ),
            ClosureAct(
                order_id=12175150,
                date=datetime.datetime(2021, 9, 13),
            ),
        )
    ],
)
def test_closure_act_parser(eml, expected):
    parser = EmailCAPParser()

    result = parser.parse(eml)

    assert result == expected
