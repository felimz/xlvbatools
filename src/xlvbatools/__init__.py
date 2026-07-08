"""
xlvbatools -- General-Purpose Python Toolkit for Headless Excel VBA Automation
===============================================================================

Top-level public API re-exports for convenient access.

Core automation:
    from xlvbatools import ExcelSession

VBA source management:
    from xlvbatools.vba import extract_all, inject_all, diff_all

Static analysis:
    from xlvbatools.analysis import lint_workbook, lint_files

Snapshots:
    from xlvbatools.snapshot import SnapshotManager

Workbook inspection:
    from xlvbatools.workbook import dump_sheet_data, export_screenshots, modify_cell
"""

__version__ = "0.1.0"

# Lazy imports to avoid pulling in win32com on import for cross-platform use.
# The actual COM-dependent classes guard themselves with require_windows().


def __getattr__(name):
    """Lazy-load core classes on first access to avoid import-time COM dependency."""
    if name == "ExcelSession":
        from xlvbatools.core.session import ExcelSession
        return ExcelSession
    if name == "DialogWatchdog":
        from xlvbatools.core.watchdog import DialogWatchdog
        return DialogWatchdog
    if name == "DialogEvent":
        from xlvbatools.core.watchdog import DialogEvent
        return DialogEvent
    if name == "SnapshotManager":
        from xlvbatools.snapshot.manager import SnapshotManager
        return SnapshotManager
    raise AttributeError(f"module 'xlvbatools' has no attribute {name!r}")


__all__ = [
    "ExcelSession",
    "DialogWatchdog",
    "DialogEvent",
    "SnapshotManager",
    "__version__",
]
