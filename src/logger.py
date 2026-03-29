"""
logger.py -- Centralized loguru configuration.

Import this module once (in app.py) to activate console + file sinks.
All other modules just do `from loguru import logger` directly.

Log levels by environment (set via ENV variable):
  prod  → WARNING (console) / INFO (file)
  test  → DEBUG   (console) / DEBUG (file)
  dev   → INFO    (console) / DEBUG (file)  [default]
"""

import sys
from pathlib import Path

from loguru import logger

from src.config import ENV

CONSOLE_LEVELS = {"prod": "WARNING", "test": "DEBUG", "dev": "INFO"}
FILE_LEVELS = {"prod": "INFO", "test": "DEBUG", "dev": "DEBUG"}

console_level = CONSOLE_LEVELS.get(ENV, "INFO")
file_level = FILE_LEVELS.get(ENV, "DEBUG")

# Project root = parent of src/
PROJECT_ROOT = Path(__file__).resolve().parent.parent
LOG_DIR = PROJECT_ROOT / "log"
LOG_FILE = LOG_DIR / "var_engine.log"

# Remove default stderr handler
logger.remove()

# Colored console output
logger.add(
    sys.stderr,
    level=console_level,
    colorize=True,
    format="<green>{time:HH:mm:ss}</green> | <level>{level: <8}</level> | <cyan>{name}</cyan>:<cyan>{function}</cyan> - <level>{message}</level>",
)

# Single rotating file
LOG_DIR.mkdir(parents=True, exist_ok=True)
logger.add(
    str(LOG_FILE),
    level=file_level,
    rotation="10 MB",
    retention="30 days",
    format="{time:YYYY-MM-DD HH:mm:ss} | {level: <8} | {name}:{function}:{line} - {message}",
)

logger.info(f"Logger initialized | ENV={ENV} | console={console_level} | file={file_level}")
