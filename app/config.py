from functools import cached_property
from logging import INFO
from logging.config import dictConfig
from pathlib import Path

import yaml

LOGGING = {
    "version": 1,
    "formatters": {
        "simple": {
            "format": "[%(asctime)s] | %(funcName)+15s:%(lineno)s | %(levelname)-8s | %(message)s"
        }
    },
    "handlers": {
        "stream": {
            "class": "logging.StreamHandler",
            "formatter": "simple",
        },
        "to_file": {
            "class": "logging.handlers.TimedRotatingFileHandler",
            "formatter": "simple",
            "filename": "app.log",
            "utc": True,
            "when": "midnight",
        },
    },
    "loggers": {
        "": {
            "handlers": ["stream", "to_file"],
            "level": INFO,
        }
    },
}

dictConfig(LOGGING)


class Config:
    def __init__(self, path: Path):
        if not path.is_file():
            raise EnvironmentError("Config %s not found", path.absolute())

        self.path = path

    @cached_property
    def config(self) -> dict:
        with open(self.path) as f:
            return yaml.load(f, Loader=yaml.FullLoader)

    @cached_property
    def template(self):
        return self.config["TEMPLATE"]

    @cached_property
    def storage(self):
        return self.config["STORAGE"]


class Template:
    def __init__(self, config: Config):
        self.config = config

    def get_template_by_name(self, name: str) -> Path:
        return self.config.template["ROOT"] / Path(f"{name}_template.xlsx")
