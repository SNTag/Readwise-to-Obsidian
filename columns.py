"""
Derived column schema — computed from the raw lists in config.py.
Import ALL_COLS, XLSX_COLS, KEY_TO_HEADER, and HEADER_TO_KEY from here,
not from config.
"""

from config import READWISE_COLS, COL_NAMES, CURATION_COLS

_n         = len(READWISE_COLS)
_overrides = (list(COL_NAMES) + [""] * _n)[:_n]

XLSX_COLS     = [_overrides[i] or READWISE_COLS[i] for i in range(_n)]
ALL_COLS      = XLSX_COLS + CURATION_COLS
KEY_TO_HEADER = dict(zip(READWISE_COLS, XLSX_COLS))
HEADER_TO_KEY = {v: k for k, v in KEY_TO_HEADER.items()}
