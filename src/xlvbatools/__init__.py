"""General-purpose toolkit for reliable, headless Excel VBA automation."""

from importlib import import_module

from xlvbatools.version import get_version_info


__version__ = get_version_info().version

_LAZY_IMPORTS = {
    "Artifact": ("xlvbatools.results", "Artifact"),
    "CleanupReport": ("xlvbatools.results", "CleanupReport"),
    "ConfigurationError": ("xlvbatools.errors", "ConfigurationError"),
    "Diagnostics": ("xlvbatools.results", "Diagnostics"),
    "DialogEvent": ("xlvbatools.core.watchdog", "DialogEvent"),
    "DialogWatchdog": ("xlvbatools.core.watchdog", "DialogWatchdog"),
    "ErrorInfo": ("xlvbatools.results", "ErrorInfo"),
    "ExcelSession": ("xlvbatools.core.session", "ExcelSession"),
    "HeadlessCleanupError": ("xlvbatools.errors", "HeadlessCleanupError"),
    "InspectionOutput": ("xlvbatools.results", "InspectionOutput"),
    "OperationFailedError": ("xlvbatools.errors", "OperationFailedError"),
    "OperationResult": ("xlvbatools.results", "OperationResult"),
    "SnapshotManager": ("xlvbatools.snapshot.manager", "SnapshotManager"),
    "TrustCenterError": ("xlvbatools.errors", "TrustCenterError"),
    "VersionInfo": ("xlvbatools.version", "VersionInfo"),
    "XlvbaConfig": ("xlvbatools.config.schema", "XlvbaConfig"),
    "XlvbaError": ("xlvbatools.errors", "XlvbaError"),
    "XlvbaProject": ("xlvbatools.project", "XlvbaProject"),
    "inspect_workbook": ("xlvbatools.workbook.dumper", "inspect_workbook"),
    "lint_files": ("xlvbatools.analysis.preflight", "lint_files"),
    "lint_workbook": ("xlvbatools.analysis.preflight", "lint_workbook"),
    "load_config": ("xlvbatools.config.loader", "load_config"),
    "run_macro": ("xlvbatools.macro.runner", "run_macro"),
}


def __getattr__(name):
    """Lazy-load public APIs so importing xlvbatools never initializes COM."""
    target = _LAZY_IMPORTS.get(name)
    if target is None:
        raise AttributeError(f"module 'xlvbatools' has no attribute {name!r}")
    module_name, attribute = target
    value = getattr(import_module(module_name), attribute)
    globals()[name] = value
    return value


__all__ = sorted([*_LAZY_IMPORTS, "__version__", "get_version_info"])
