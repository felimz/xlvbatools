"""Reliable, isolated Excel/VBA automation for Python projects."""

from xlvbatools._version import __version__
from xlvbatools.analysis.issue import VBAIssue
from xlvbatools.analysis.filtering import LINT_BASELINE_SCHEMA_VERSION
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
    AttemptDiagnostic,
    CleanupReport,
    Diagnostics,
    ErrorInfo,
    InspectionOutput,
    OperationResult,
    RESULT_SCHEMA_VERSION,
    WorkerExitReport,
)
from xlvbatools.snapshots import SnapshotGitInfo, SnapshotRecord, SnapshotService
from xlvbatools.version import VersionInfo, get_version_info
from xlvbatools.workflow import (
    InspectStep,
    MacroStep,
    ModifyStep,
    ModifyStepOutput,
    RangeWriteResult,
    WORKFLOW_SCHEMA_VERSION,
    WorkflowOutput,
    WorkflowStep,
    WorkflowStepResult,
)


__all__ = [
    "Artifact",
    "AttemptDiagnostic",
    "CleanupReport",
    "ComponentDiff",
    "ConfigurationError",
    "Diagnostics",
    "ErrorInfo",
    "Executor",
    "ExtractionOutput",
    "HeadlessCleanupError",
    "InspectionOutput",
    "InspectStep",
    "InjectionChange",
    "InjectionOutput",
    "IsolatedExecutor",
    "LINT_BASELINE_SCHEMA_VERSION",
    "MacroOutput",
    "MacroStep",
    "ModificationOutput",
    "ModifyStep",
    "ModifyStepOutput",
    "Operation",
    "OperationFailedError",
    "OperationRequest",
    "OperationResult",
    "Project",
    "ProjectSettings",
    "RESULT_SCHEMA_VERSION",
    "RangeWriteResult",
    "SnapshotError",
    "SnapshotGitInfo",
    "SnapshotNotFoundError",
    "SnapshotRecord",
    "SnapshotService",
    "TrustCenterError",
    "VBAComponent",
    "VBAIssue",
    "VersionInfo",
    "WorkerExitReport",
    "WORKFLOW_SCHEMA_VERSION",
    "WorkflowOutput",
    "WorkflowStep",
    "WorkflowStepResult",
    "XlvbaError",
    "__version__",
    "get_version_info",
]
