import shutil
from pathlib import Path

import pytest

from app.config import Config

config = Config(Path("../data/test_cfg.yml"))


@pytest.fixture(autouse=True)
def root():
    yield config.storage["ROOT"]
    shutil.rmtree(config.storage["ROOT"], ignore_errors=True)
