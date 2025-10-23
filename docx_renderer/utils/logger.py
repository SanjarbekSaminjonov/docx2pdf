"""Central logging configuration for the library."""
from __future__ import annotations

import logging
from typing import Optional

_DEFAULT_LEVEL = logging.INFO


def get_logger(name: Optional[str] = None) -> logging.Logger:
    """Return a module-level logger with default configuration applied."""
    logger = logging.getLogger(name)
    if not logging.getLogger().handlers:
        logging.basicConfig(level=_DEFAULT_LEVEL, format="%(asctime)s [%(levelname)s] %(name)s: %(message)s")
    return logger
