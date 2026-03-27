"""
config.py -- Load application configuration from config.yaml.

Usage:
    from src.config import TICKERS, LOOKBACK_DAYS, STRESS_LABEL, STRESS_START_DATE, STRESS_END_DATE
"""

from pathlib import Path

import yaml

_CONFIG_PATH = Path(__file__).resolve().parent.parent / "config.yaml"

with open(_CONFIG_PATH) as _f:
    _cfg = yaml.safe_load(_f)

TICKERS: list[str] = _cfg["tickers"]
LOOKBACK_DAYS: int = _cfg["lookback_days"]
STRESS_LABEL: str = _cfg["stressed_period_label"]
STRESS_START_DATE: str = _cfg["stressed_period_start_date"]
STRESS_END_DATE: str = _cfg["stressed_period_end_date"]
