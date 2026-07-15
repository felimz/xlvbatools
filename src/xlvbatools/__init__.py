"""Reliable, isolated Excel/VBA automation for Python projects."""

from xlvbatools._version import __version__
from xlvbatools.analysis.issue import VBAIssue
from xlvbatools.errors import (
    ConfigurationError,
    HeadlessCleanupError,
    OperationFailedError,
    TrustCenterError,
    XlvbaError,
)
from xlvbatools.execution import (
    Executor,
    IsolatedExecutor,
    Operation,
    OperationRequest,
)
from xlvbatools.project import Project, ProjectSettings
from xlvbatools.results import (
    Artifact,
    CleanupReport,
    Diagnostics,
    ErrorInfo,
    InspectionOutput,
    OperationResult,
    RESULT_SCHEMA_VERSION,
)
from xlvbatools.version import VersionInfo, get_version_info


__all__ = [
    "Artifact",
    "CleanupReport",
    "ConfigurationError",
    "Diagnostics",
    "ErrorInfo",
    "Executor",
    "HeadlessCleanupError",
    "InspectionOutput",
    "IsolatedExecutor",
    "Operation",
    "OperationFailedError",
    "OperationRequest",
    "OperationResult",
    "Project",
    "ProjectSettings",
    "RESULT_SCHEMA_VERSION",
    "TrustCenterError",
    "VBAIssue",
    "VersionInfo",
    "XlvbaError",
    "__version__",
    "get_version_info",
]
