"""Typed, versioned result contracts for public xlvbatools operations."""

from __future__ import annotations

from dataclasses import dataclass, field, fields, is_dataclass
from pathlib import Path
from typing import Any, Generic, Mapping, Optional, TypeVar

from xlvbatools.errors import HeadlessCleanupError, OperationFailedError


SCHEMA_VERSION = "1.0"
T = TypeVar("T")


def _plain(value: Any) -> Any:
    """Convert public result values into JSON-compatible Python containers."""
    if is_dataclass(value) and not isinstance(value, type):
        return {item.name: _plain(getattr(value, item.name)) for item in fields(value)}
    if hasattr(value, "to_dict") and callable(value.to_dict):
        return _plain(value.to_dict())
    if isinstance(value, Mapping):
        return {str(key): _plain(item) for key, item in value.items()}
    if isinstance(value, (list, tuple, set)):
        return [_plain(item) for item in value]
    if isinstance(value, Path):
        return str(value)
    return value


@dataclass(frozen=True)
class ErrorInfo:
    """Stable machine-readable description of an operation failure."""

    message: str
    code: str = "operation_failed"
    error_type: Optional[str] = None
    details: Mapping[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class Artifact:
    """A file or other durable output produced by an operation."""

    kind: str
    path: str
    media_type: Optional[str] = None
    label: Optional[str] = None
    metadata: Mapping[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class CleanupReport:
    """Lifecycle report for an Excel process owned by one operation."""

    pid: Optional[int] = None
    quit_requested: bool = False
    exited_gracefully: bool = False
    force_terminated: bool = False
    still_running: bool = False
    worker_terminated: bool = False
    workbook_close_error: Optional[str] = None
    workbook_save_error: Optional[str] = None
    details: Mapping[str, Any] = field(default_factory=dict)

    @classmethod
    def from_mapping(cls, value: Mapping[str, Any]) -> "CleanupReport":
        known = {
            "pid", "quit_requested", "exited_gracefully", "force_terminated",
            "still_running", "worker_terminated", "workbook_close_error",
            "workbook_save_error",
        }
        return cls(
            pid=value.get("pid"),
            quit_requested=bool(value.get("quit_requested", False)),
            exited_gracefully=bool(value.get("exited_gracefully", False)),
            force_terminated=bool(value.get("force_terminated", False)),
            still_running=bool(value.get("still_running", False)),
            worker_terminated=bool(value.get("worker_terminated", False)),
            workbook_close_error=value.get("workbook_close_error"),
            workbook_save_error=value.get("workbook_save_error"),
            details={key: item for key, item in value.items() if key not in known},
        )

    @property
    def is_clean(self) -> bool:
        """Whether Excel exited naturally without dialogs being force-resolved."""
        return bool(
            self.pid is not None
            and self.quit_requested
            and self.exited_gracefully
            and not self.force_terminated
            and not self.worker_terminated
            and not self.still_running
            and not self.workbook_close_error
            and not self.workbook_save_error
        )


@dataclass(frozen=True)
class Diagnostics:
    """Cross-operation diagnostics retained for logging and automation."""

    dialog_events: tuple[Mapping[str, Any], ...] = ()
    cleanup: Optional[CleanupReport] = None
    com_error: Optional[Mapping[str, Any]] = None
    worker_pid: Optional[int] = None
    excel_pid: Optional[int] = None

    @classmethod
    def from_legacy(cls, value: Mapping[str, Any]) -> "Diagnostics":
        cleanup_value = value.get("cleanup") or None
        cleanup = (
            CleanupReport.from_mapping(cleanup_value)
            if isinstance(cleanup_value, Mapping) else None
        )
        return cls(
            dialog_events=tuple(value.get("dialog_events") or ()),
            cleanup=cleanup,
            com_error=value.get("com_error"),
            worker_pid=value.get("worker_pid"),
            excel_pid=value.get("excel_pid") or (cleanup.pid if cleanup else None),
        )


@dataclass(frozen=True)
class InspectionOutput:
    """Structured workbook data plus screenshot paths from one inspection."""

    workbook_data: Optional[Mapping[str, Any]]
    screenshots: Mapping[str, str] = field(default_factory=dict)


@dataclass(frozen=True)
class OperationResult(Generic[T]):
    """Stable outer envelope shared by all wrapper-facing operations."""

    operation: str
    success: bool
    phase: str
    data: Optional[T] = None
    error: Optional[ErrorInfo] = None
    warnings: tuple[str, ...] = ()
    artifacts: tuple[Artifact, ...] = ()
    diagnostics: Diagnostics = field(default_factory=Diagnostics)
    metadata: Mapping[str, Any] = field(default_factory=dict)
    schema_version: str = SCHEMA_VERSION

    @classmethod
    def from_legacy(
        cls,
        operation: str,
        value: Mapping[str, Any],
        *,
        data: Optional[T] = None,
        artifacts: tuple[Artifact, ...] = (),
        metadata: Optional[Mapping[str, Any]] = None,
    ) -> "OperationResult[T]":
        success = bool(value.get("success"))
        error = None
        if not success:
            message = str(
                value.get("primary_error")
                or value.get("error")
                or f"{operation} failed"
            )
            error = ErrorInfo(
                message=message,
                code="timeout" if value.get("timed_out") else "operation_failed",
                error_type=value.get("error_type"),
            )
        return cls(
            operation=operation,
            success=success,
            phase=str(value.get("phase") or ("complete" if success else "unknown")),
            data=data,
            error=error,
            warnings=tuple(value.get("warnings") or ()),
            artifacts=artifacts,
            diagnostics=Diagnostics.from_legacy(value),
            metadata=dict(metadata or {}),
        )

    @classmethod
    def failed(
        cls,
        operation: str,
        error: Exception,
        *,
        phase: str = "setup",
        code: str = "operation_failed",
    ) -> "OperationResult[T]":
        return cls(
            operation=operation,
            success=False,
            phase=phase,
            error=ErrorInfo(
                message=str(error), code=code, error_type=type(error).__name__,
            ),
        )

    def require_success(self) -> Optional[T]:
        """Return data or raise a stable public exception on failure."""
        if not self.success:
            raise OperationFailedError(self)
        return self.data

    def require_clean_shutdown(self) -> CleanupReport:
        """Return cleanup details or raise when owned Excel did not exit cleanly."""
        cleanup = self.diagnostics.cleanup
        if cleanup is None:
            raise HeadlessCleanupError(self, "Operation did not report Excel cleanup")
        if not cleanup.is_clean:
            raise HeadlessCleanupError(
                self,
                f"Owned Excel PID {cleanup.pid!r} did not exit cleanly",
            )
        return cleanup

    def to_dict(self) -> dict[str, Any]:
        """Return the versioned JSON-compatible public representation."""
        return _plain(self)
