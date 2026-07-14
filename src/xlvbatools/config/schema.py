"""
Configuration Schema
=====================
Dataclass definitions for xlvbatools.toml configuration.
"""

import os
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class SnapshotConfig:
    """Snapshot system configuration."""
    rolling_limit: int = 10


@dataclass
class BackupConfig:
    """Injection backup configuration."""
    limit: int = 5


@dataclass
class LintConfig:
    """Static analysis configuration."""
    protected_sheets: list = field(default_factory=list)
    disabled_rules: list = field(default_factory=list)


@dataclass
class XlvbaConfig:
    """
    Complete project configuration, typically loaded from xlvbatools.toml.

    Attributes
    ----------
    workbook : str
        Path to the .xlsm workbook (relative or absolute).
    vba_source : str
        Path to the vba_source/ directory.
    snapshots_dir : str
        Path to the snapshots directory.
    log_dir : str
        Directory for log files.
    log_name : str
        Base name for the log file.
    snapshots : SnapshotConfig
        Snapshot system settings.
    backups : BackupConfig
        Injection backup settings.
    lint : LintConfig
        Static analysis settings.
    """
    workbook: str = "workbook.xlsm"
    vba_source: str = "vba_source"
    snapshots_dir: str = "snapshots"
    log_dir: str = "logs"
    log_name: str = "xlvbatools"
    snapshots: SnapshotConfig = field(default_factory=SnapshotConfig)
    backups: BackupConfig = field(default_factory=BackupConfig)
    lint: LintConfig = field(default_factory=LintConfig)
    config_dir: Optional[str] = None

    def _resolve_path(self, value: str) -> str:
        """Resolve a configured project path relative to the TOML directory."""
        if self.config_dir and not os.path.isabs(value):
            return os.path.abspath(os.path.join(self.config_dir, value))
        return os.path.abspath(value)

    @property
    def workbook_path(self) -> str:
        """Resolve workbook path to absolute."""
        return self._resolve_path(self.workbook)

    @property
    def vba_source_path(self) -> str:
        """Resolve vba_source path to absolute."""
        return self._resolve_path(self.vba_source)

    @property
    def snapshots_path(self) -> str:
        """Resolve snapshots directory to absolute."""
        return self._resolve_path(self.snapshots_dir)

    @property
    def log_dir_path(self) -> str:
        """Resolve log directory relative to the project configuration."""
        return self._resolve_path(self.log_dir)

    def validate(self) -> list[str]:
        """Validate configuration, returning list of error messages."""
        errors = []
        if not self.workbook.endswith((".xlsm", ".xlsb", ".xls")):
            errors.append(f"Workbook must be .xlsm/.xlsb/.xls, got: {self.workbook}")
        if self.snapshots.rolling_limit < 1:
            errors.append(f"Rolling limit must be >= 1, got: {self.snapshots.rolling_limit}")
        if self.backups.limit < 0:
            errors.append(f"Backup limit must be >= 0, got: {self.backups.limit}")
        return errors
