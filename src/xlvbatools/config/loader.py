"""
Configuration Loader
=====================
Reads xlvbatools.toml from the project directory, merges with defaults,
and returns a validated XlvbaConfig.

Search strategy: walk up from CWD (or a given starting directory) until
xlvbatools.toml is found, or return defaults.
"""

import logging
import os
import sys

from xlvbatools.config.schema import (
    XlvbaConfig, SnapshotConfig, BackupConfig, LintConfig,
)

logger = logging.getLogger(__name__)

CONFIG_FILENAME = "xlvbatools.toml"


def load_config(start_dir: str | None = None) -> XlvbaConfig:
    """
    Load configuration from xlvbatools.toml, walking up from start_dir.

    If no config file is found, returns defaults.

    Parameters
    ----------
    start_dir : str, optional
        Starting directory (default: CWD).

    Returns
    -------
    XlvbaConfig
        Parsed and validated configuration.
    """
    config_path = find_config(start_dir)

    if config_path is None:
        logger.debug("No xlvbatools.toml found, using defaults")
        return XlvbaConfig()

    logger.info(f"Loading config from: {config_path}")
    return _parse_toml(config_path)


def find_config(start_dir: str | None = None) -> str | None:
    """
    Walk up from start_dir looking for xlvbatools.toml.

    Returns the absolute path to the config file, or None.
    """
    current = os.path.abspath(start_dir or os.getcwd())

    for _ in range(50):  # Safety limit
        candidate = os.path.join(current, CONFIG_FILENAME)
        if os.path.isfile(candidate):
            return candidate

        parent = os.path.dirname(current)
        if parent == current:
            break
        current = parent

    return None


def _parse_toml(config_path: str) -> XlvbaConfig:
    """Parse a TOML file into an XlvbaConfig."""
    # Python 3.11+ has tomllib built-in
    if sys.version_info >= (3, 11):
        import tomllib
    else:
        try:
            import tomllib
        except ImportError:
            try:
                import tomli as tomllib
            except ImportError:
                logger.warning(
                    "Neither tomllib (Python 3.11+) nor tomli found. "
                    "Install tomli for Python < 3.11: pip install tomli"
                )
                return XlvbaConfig()

    with open(config_path, "rb") as f:
        data = tomllib.load(f)

    section = data.get("xlvbatools", {})

    # Build nested configs
    snap_data = section.get("snapshots", {})
    backup_data = section.get("backups", {})
    lint_data = section.get("lint", {})

    cfg = XlvbaConfig(
        workbook=section.get("workbook", "workbook.xlsm"),
        vba_source=section.get("vba_source", "vba_source"),
        snapshots_dir=section.get("snapshots_dir", "snapshots"),
        log_dir=section.get("log_dir", "logs"),
        log_name=section.get("log_name", "xlvbatools"),
        snapshots=SnapshotConfig(
            rolling_limit=snap_data.get("rolling_limit", 10),
        ),
        backups=BackupConfig(
            limit=backup_data.get("limit", 5),
        ),
        lint=LintConfig(
            protected_sheets=lint_data.get("protected_sheets", []),
            disabled_rules=lint_data.get("disabled_rules", []),
        ),
    )

    errors = cfg.validate()
    if errors:
        for err in errors:
            logger.warning(f"Config validation: {err}")

    return cfg
