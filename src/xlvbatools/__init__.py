"""Reliable, isolated Excel/VBA automation for Python projects."""

from xlvbatools._version import __version__
from xlvbatools.analysis.issue import VBAIssue
from xlvbatools.errors import (
    ConfigurationError,
    HeadlessCleanupError,
    OperationFailedError,
    SnapshotError,
    SnapshotNotFoundError,
    TrustCenterError,
    XlvbaError,
)
from xlvbatools.execution import (
    Executor,
    IsolatedExecutor,
    Operation,
    OperationRequest,
)
from xlvbatools.outputs import (
    ComponentDiff,
    ExtractionOutput,
    InjectionChange,
    InjectionOutput,
    MacroOutput,
    ModificationOutput,
    VBAComponent,
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
from xlvbatools.snapshots import SnapshotGitInfo, SnapshotRecord, SnapshotService
from xlvbatools.version import VersionInfo, get_version_info


__all__ = [
    "Artifact",
    "CleanupReport",
    "ComponentDiff",
    "ConfigurationError",
    "Diagnostics",
    "ErrorInfo",
    "Executor",
    "ExtractionOutput",
    "HeadlessCleanupError",
    "InspectionOutput",
    "InjectionChange",
    "InjectionOutput",
    "IsolatedExecutor",
    "MacroOutput",
    "ModificationOutput",
    "Operation",
    "OperationFailedError",
    "OperationRequest",
    "OperationResult",
    "Project",
    "ProjectSettings",
    "RESULT_SCHEMA_VERSION",
    "SnapshotError",
    "SnapshotGitInfo",
    "SnapshotNotFoundError",
    "SnapshotRecord",
    "SnapshotService",
    "TrustCenterError",
    "VBAComponent",
    "VBAIssue",
    "VersionInfo",
    "XlvbaError",
    "__version__",
    "get_version_info",
]
