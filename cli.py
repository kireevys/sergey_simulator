#!python

import fire

from app.main import bulk, run, close_orders, fill_index  # noqa

if __name__ == "__main__":
    fire.Fire()
