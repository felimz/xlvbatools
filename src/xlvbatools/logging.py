"""
Centralized Logging for xlvbatools
====================================
Configures console and rotating file handlers for all CLI and library operations.

Usage in CLI tools:
    from xlvbatools.logging import setup_logging
    log_file = setup_logging(verbose=args.verbose, tool_name="extract")

Usage in library code (no changes needed):
    import logging
    logger = logging.getLogger(__name__)
    logger.info("This goes to both console and file")
"""

import logging
from logging.handlers import RotatingFileHandler
import os
import sys


# Formats
CONSOLE_FORMAT = "%(levelname)s: %(message)s"
FILE_FORMAT = "%(asctime)s %(levelname)-8s [%(name)s] %(message)s"
FILE_DATE_FORMAT = "%Y-%m-%d %H:%M:%S"


def setup_logging(
    verbose: bool = False,
    tool_name: str = "xlvbatools",
    log_dir: str | None = None,
    log_name: str = "xlvbatools",
) -> str:
    """
    Configure logging for the current session.

    Parameters
    ----------
    verbose : bool
        If True, console shows DEBUG. Otherwise INFO.
    tool_name : str
        Name of the CLI tool or operation (logged at session start).
    log_dir : str, optional
        Directory for log files. Defaults to ``logs/`` in CWD.
    log_name : str
        Base name for the log file (default: ``xlvbatools``).

    Returns
    -------
    str
        Absolute path to the log file.
    """
    if log_dir is None:
        log_dir = os.path.join(os.getcwd(), "logs")
    os.makedirs(log_dir, exist_ok=True)

    log_filepath = os.path.join(log_dir, f"{log_name}.log")

    # Get root logger
    root_logger = logging.getLogger()
    root_logger.handlers.clear()
    root_logger.setLevel(logging.DEBUG)

    # Console handler (stderr)
    console_handler = logging.StreamHandler(sys.stderr)
    console_handler.setLevel(logging.DEBUG if verbose else logging.INFO)
    console_handler.setFormatter(logging.Formatter(CONSOLE_FORMAT))
    root_logger.addHandler(console_handler)

    # Rotating File handler (always DEBUG, max 1MB, 3 backups)
    file_handler = RotatingFileHandler(
        log_filepath, maxBytes=1024 * 1024, backupCount=3, encoding="utf-8"
    )
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(
        logging.Formatter(FILE_FORMAT, datefmt=FILE_DATE_FORMAT)
    )
    root_logger.addHandler(file_handler)

    # Log the session start
    logger = logging.getLogger(tool_name)
    logger.debug(f"Session started: {tool_name}")
    logger.debug(f"Log file: {log_filepath}")
    logger.debug(f"Python: {sys.version}")
    logger.debug(f"CWD: {os.getcwd()}")

    return log_filepath


def get_latest_log(log_dir: str | None = None) -> str | None:
    """Return the path to the main log file, or None."""
    if log_dir is None:
        log_dir = os.path.join(os.getcwd(), "logs")
    log_file = os.path.join(log_dir, "xlvbatools.log")
    return log_file if os.path.exists(log_file) else None
