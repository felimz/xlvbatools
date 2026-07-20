"""Typed, versioned result contracts for the xlvbatools public API."""

from __future__ import annotations

from dataclasses import dataclass, field, fields, is_dataclass, replace
from pathlib import Path
from typing import Any, Callable, Generic, Mapping, Optional, TypeVar, cast

from xlvbatools.errors import HeadlessCleanupError, OperationFailedError


RESULT_SCHEMA_VERSION = "1.3"
T = TypeVar("T")
U = TypeVar("U")


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
            "workbook_save_error", "details",
        }
        details = dict(value.get("details") or {})
        details.update({
            key: item for key, item in value.items() if key not in known
        })
        return cls(
            pid=value.get("pid"),
            quit_requested=bool(value.get("quit_requested", False)),
            exited_gracefully=bool(value.get("exited_gracefully", False)),
            force_terminated=bool(value.get("force_terminated", False)),
            still_running=bool(value.get("still_running", False)),
            worker_terminated=bool(value.get("worker_terminated", False)),
            workbook_close_error=value.get("workbook_close_error"),
            workbook_save_error=value.get("workbook_save_error"),
            details=details,
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
class WorkerExitReport:
    """Lifecycle evidence for the isolated worker process itself."""

    pid: Optional[int] = None
    exit_code: Optional[int] = None
    exited: bool = False
    reaped: bool = False
    force_terminated: bool = False
    still_running: bool = False

    @classmethod
    def from_mapping(cls, value: Mapping[str, Any]) -> "WorkerExitReport":
        return cls(
            pid=value.get("pid"),
            exit_code=value.get("exit_code"),
            exited=bool(value.get("exited", False)),
            reaped=bool(value.get("reaped", False)),
            force_terminated=bool(value.get("force_terminated", False)),
            still_running=bool(value.get("still_running", False)),
        )

    @property
    def is_clean(self) -> bool:
        """Whether the worker exited and was reaped without being killed."""
        return bool(
            self.pid is not None
            and self.exited
            and self.reaped
            and not self.force_terminated
            and not self.still_running
        )


@dataclass(frozen=True)
class AttemptDiagnostic:
    """Evidence retained for one executor attempt and its retry decision."""

    attempt: int
    phase: str
    request_id: Optional[str] = None
    error_code: Optional[str] = None
    error_message: Optional[str] = None
    error_details: Mapping[str, Any] = field(default_factory=dict)
    worker: Optional[WorkerExitReport] = None
    excel_pid: Optional[int] = None
    cleanup: Optional[CleanupReport] = None
    dialog_count: int = 0
    elapsed_seconds: Optional[float] = None
    retryable: bool = False
    retry_reason: Optional[str] = None

    @classmethod
    def from_mapping(cls, value: Mapping[str, Any]) -> "AttemptDiagnostic":
        worker_value = value.get("worker")
        cleanup_value = value.get("cleanup")
        return cls(
            attempt=int(value.get("attempt") or 1),
            phase=str(value.get("phase") or "unknown"),
            request_id=value.get("request_id"),
            error_code=value.get("error_code"),
            error_message=value.get("error_message"),
            error_details=dict(value.get("error_details") or {}),
            worker=(
                WorkerExitReport.from_mapping(worker_value)
                if isinstance(worker_value, Mapping) else None
            ),
            excel_pid=value.get("excel_pid"),
            cleanup=(
                CleanupReport.from_mapping(cleanup_value)
                if isinstance(cleanup_value, Mapping) else None
            ),
            dialog_count=int(value.get("dialog_count") or 0),
            elapsed_seconds=value.get("elapsed_seconds"),
            retryable=bool(value.get("retryable", False)),
            retry_reason=value.get("retry_reason"),
        )


@dataclass(frozen=True)
class Diagnostics:
    """Cross-operation diagnostics retained for logging and automation."""

    dialog_events: tuple[Mapping[str, Any], ...] = ()
    cleanup: Optional[CleanupReport] = None
    com_error: Optional[Mapping[str, Any]] = None
    worker_pid: Optional[int] = None
    excel_pid: Optional[int] = None
    worker_exit: Optional[WorkerExitReport] = None
    attempts: tuple[AttemptDiagnostic, ...] = ()
    progress: Mapping[str, Any] = field(default_factory=dict)

    @classmethod
    def _from_worker(cls, value: Mapping[str, Any]) -> "Diagnostics":
        cleanup_value = value.get("cleanup") or None
        cleanup = (
            CleanupReport.from_mapping(cleanup_value)
            if isinstance(cleanup_value, Mapping) else None
        )
        worker_exit_value = value.get("worker_exit") or None
        worker_exit = (
            WorkerExitReport.from_mapping(worker_exit_value)
            if isinstance(worker_exit_value, Mapping) else None
        )
        raw_attempts = value.get("attempts") or ()
        return cls(
            dialog_events=tuple(value.get("dialog_events") or ()),
            cleanup=cleanup,
            com_error=value.get("com_error"),
            worker_pid=value.get("worker_pid"),
            excel_pid=value.get("excel_pid") or (cleanup.pid if cleanup else None),
            worker_exit=worker_exit,
            attempts=tuple(
                AttemptDiagnostic.from_mapping(item)
                for item in raw_attempts
                if isinstance(item, Mapping)
            ),
            progress=dict(value.get("progress") or {}),
        )


@dataclass(frozen=True)
class InspectionOutput:
    """Structured workbook data plus screenshot paths from one inspection."""

    workbook_data: Optional[Mapping[str, Any]]
    screenshots: Mapping[str, str] = field(default_factory=dict)
    screenshot_diagnostics: Mapping[str, Any] = field(default_factory=dict)


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
    request_id: Optional[str] = None
    elapsed_seconds: Optional[float] = None
    attempt_count: int = 1
    schema_version: str = RESULT_SCHEMA_VERSION

    @classmethod
    def _from_worker(
        cls,
        value: Mapping[str, Any],
        *,
        data: Optional[T] = None,
        artifacts: tuple[Artifact, ...] = (),
        metadata: Optional[Mapping[str, Any]] = None,
    ) -> "OperationResult[T]":
        operation = str(value.get("operation") or "unknown")
        success = bool(value.get("success"))
        error = None
        if not success:
            raw_error = value.get("error")
            structured_error = raw_error if isinstance(raw_error, Mapping) else {}
            message = str(
                value.get("primary_error")
                or structured_error.get("message")
                or raw_error
                or f"{operation} failed"
            )
            details = dict(structured_error.get("details") or {})
            for key in ("traceback", "worker_output", "timed_out", "progress"):
                if value.get(key) is not None:
                    details[key] = value[key]
            error = ErrorInfo(
                message=message,
                code=str(
                    structured_error.get("code")
                    or ("timeout" if value.get("timed_out") else "operation_failed")
                ),
                error_type=value.get("error_type") or structured_error.get("error_type"),
                details=details,
            )
        return cls(
            operation=operation,
            success=success,
            phase=str(value.get("phase") or ("complete" if success else "unknown")),
            data=data,
            error=error,
            warnings=tuple(value.get("warnings") or ()),
            artifacts=artifacts,
            diagnostics=Diagnostics._from_worker(value),
            metadata=dict(metadata or {}),
            request_id=value.get("request_id"),
            elapsed_seconds=value.get("elapsed_seconds"),
            attempt_count=int(value.get("attempt_count") or 1),
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

    def map_data(
        self, transform: Callable[[T], U],
    ) -> "OperationResult[U]":
        """Transform available operation data while preserving its envelope."""
        if self.data is None:
            return cast("OperationResult[U]", self)
        return cast(
            "OperationResult[U]",
            replace(self, data=cast(Any, transform(self.data))),
        )

    def to_dict(self) -> dict[str, Any]:
        """Return the versioned JSON-compatible public representation."""
        return _plain(self)
