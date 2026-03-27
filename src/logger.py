"""
logger.py -- Centralized loguru configuration.

Import this module once (in app.py) to activate console + file sinks.
All other modules just do `from loguru import logger` directly.
"""

import sys
from pathlib import Path

from loguru import logger

# Project root = parent of src/
PROJECT_ROOT = Path(__file__).resolve().parent.parent
LOG_DIR = PROJECT_ROOT / "log"
LOG_FILE = LOG_DIR / "var_engine.log"

# Remove default stderr handler
logger.remove()

# Colored console output (INFO+)
logger.add(
    sys.stderr,
    level="INFO",
    colorize=True,
    format="<green>{time:HH:mm:ss}</green> | <level>{level: <8}</level> | <cyan>{name}</cyan>:<cyan>{function}</cyan> - <level>{message}</level>",
)

# Single rotating file (DEBUG+)
LOG_DIR.mkdir(parents=True, exist_ok=True)
logger.add(
    str(LOG_FILE),
    level="INFO",
    rotation="10 MB",
    retention="30 days",
    format="{time:YYYY-MM-DD HH:mm:ss} | {level: <8} | {name}:{function}:{line} - {message}",
)
