"""Typed execution boundary for all isolated Excel operations."""

from __future__ import annotations

from dataclasses import dataclass, field, replace
from enum import Enum
import subprocess
import time
from types import MappingProxyType
from typing import Any, Mapping, Protocol

from xlvbatools.core.protocol import WORKER_PROTOCOL_VERSION
from xlvbatools.results import AttemptDiagnostic, Diagnostics, OperationResult


class Operation(str, Enum):
    """Operations supported by the isolated Excel worker."""

    INSPECT = "inspect"
    RUN = "run_macro"
    LIST_COMPONENTS = "list_components"
    EXTRACT = "extract"
    INJECT = "inject"
    DIFF = "diff"
    LINT_WORKBOOK = "lint_workbook"
    MODIFY = "modify"
    WORKFLOW = "workflow"


def _freeze(value: Any) -> Any:
    """Return a recursively immutable copy of JSON-compatible request data."""
    if isinstance(value, Mapping):
        return MappingProxyType({str(key): _freeze(item) for key, item in value.items()})
    if isinstance(value, (list, tuple)):
        return tuple(_freeze(item) for item in value)
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    raise TypeError(
        "operation arguments must contain only JSON-compatible mappings, "
        f"sequences, and scalar values; got {type(value).__name__}"
    )


def _thaw(value: Any) -> Any:
    """Convert frozen request data into ordinary JSON-compatible containers."""
    if isinstance(value, Mapping):
        return {str(key): _thaw(item) for key, item in value.items()}
    if isinstance(value, tuple):
        return [_thaw(item) for item in value]
    return value


@dataclass(frozen=True)
class OperationRequest:
    """Validated request passed from the public API to an executor."""

    operation: Operation
    arguments: Mapping[str, Any] = field(default_factory=dict)
    timeout: float = 60.0
    retry_transient: bool = False

    def __post_init__(self) -> None:
        if self.timeout <= 0:
            raise ValueError("timeout must be greater than zero")
        if self.retry_transient and self.operation is not Operation.MODIFY:
            raise ValueError("transient retries are supported only for modification")
        object.__setattr__(self, "arguments", _freeze(self.arguments))

    def _worker_arguments(self) -> dict[str, Any]:
        """Return a detached mutable payload for the private worker transport."""
        return _thaw(self.arguments)


class Executor(Protocol):
    """Dependency-injection boundary used by ``Project`` and tests."""

    def execute(self, request: OperationRequest) -> OperationResult[Any]: ...


class IsolatedExecutor:
    """Execute an operation with one bounded, ownership-safe retry at most."""

    def execute(self, request: OperationRequest) -> OperationResult[Any]:
        started = time.monotonic()
        deadline = started + request.timeout
        attempts: list[AttemptDiagnostic] = []

        for attempt_number in (1, 2):
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                break
            result = self._execute_once(request, timeout=remaining)
            retry_reason = None
            if attempt_number == 1 and not result.success:
                retry_reason = self._startup_retry_reason(result)
                if retry_reason is None and request.retry_transient:
                    retry_reason = self._transient_retry_reason(result)

            will_retry = bool(
                retry_reason is not None and deadline - time.monotonic() > 0
            )
            attempts.append(self._attempt_diagnostic(
                result,
                attempt_number=attempt_number,
                retry_reason=retry_reason if will_retry else None,
            ))
            if not will_retry:
                return self._with_attempts(result, attempts, started=started)

        # A positive timeout always permits the first attempt. This fallback is
        # defensive against an unusual monotonic clock jump before dispatch.
        result = OperationResult.failed(
            request.operation.value,
            TimeoutError(f"{request.operation.value} exhausted its timeout budget"),
            phase="transport",
            code="timeout",
        )
        return self._with_attempts(result, attempts, started=started)

    @staticmethod
    def _execute_once(
        request: OperationRequest,
        *,
        timeout: float,
    ) -> OperationResult[Any]:
        from xlvbatools.core.worker import WorkerCreationError, execute_worker_request

        try:
            response = execute_worker_request(
                request.operation.value,
                request._worker_arguments(),
                timeout=timeout,
            )
        except WorkerCreationError as error:
            return OperationResult.failed(
                request.operation.value,
                error,
                phase="transport",
                code="worker_start_failed",
            )
        except (OSError, subprocess.SubprocessError) as error:
            return OperationResult.failed(
                request.operation.value,
                error,
                phase="transport",
                code="worker_transport_failed",
            )
        if response.get("protocol_version") != WORKER_PROTOCOL_VERSION:
            return OperationResult.failed(
                request.operation.value,
                RuntimeError(
                    f"Worker protocol mismatch: expected {WORKER_PROTOCOL_VERSION!r}, got "
                    f"{response.get('protocol_version')!r}"
                ),
                phase="transport",
                code="protocol_mismatch",
            )
        if response.get("operation") != request.operation.value:
            return OperationResult.failed(
                request.operation.value,
                RuntimeError(
                    f"Worker returned operation {response.get('operation')!r} "
                    f"for {request.operation.value!r}"
                ),
                phase="transport",
                code="protocol_mismatch",
            )
        if not response.get("request_id") or not isinstance(response.get("success"), bool):
            return OperationResult.failed(
                request.operation.value,
                RuntimeError("Worker returned an incomplete result envelope"),
                phase="transport",
                code="invalid_worker_result",
            )
        return OperationResult._from_worker(response, data=response.get("data"))

    @staticmethod
    def _startup_retry_reason(result: OperationResult[Any]) -> str | None:
        """Return a reason only for a proven pre-ownership startup failure."""
        if result.success or result.phase not in {"transport", "worker_start"}:
            return None
        if result.error is None or result.error.code != "worker_start_failed":
            return None
        diagnostics = result.diagnostics
        if diagnostics.excel_pid is not None or diagnostics.dialog_events:
            return None
        if result.error.details.get("timed_out"):
            return None

        if result.phase == "transport":
            # Popen itself failed, so no worker or Excel process ever existed.
            if diagnostics.worker_pid is None and diagnostics.worker_exit is None:
                return "worker_creation_failed"
            return None

        worker = diagnostics.worker_exit
        if worker is None or not worker.is_clean:
            return None
        cleanup = diagnostics.cleanup
        if cleanup is not None and (
            cleanup.pid is not None
            or cleanup.force_terminated
            or cleanup.worker_terminated
            or cleanup.still_running
            or cleanup.workbook_close_error
            or cleanup.workbook_save_error
        ):
            return None
        return "worker_exited_before_session_start"

    @staticmethod
    def _transient_retry_reason(result: OperationResult[Any]) -> str | None:
        """Recognize the existing idempotent modification COM failure set."""
        cleanup = result.diagnostics.cleanup
        worker = result.diagnostics.worker_exit
        if (
            cleanup is None
            or cleanup.pid is None
            or cleanup.still_running
            or worker is None
            or not worker.is_clean
        ):
            return None
        if result.diagnostics.dialog_events:
            return None
        error = result.error
        if error is None:
            return None
        message = " ".join((
            error.message,
            str(error.details.get("worker_output") or ""),
        )).lower()
        markers = (
            "0x800706ba", "-2147023174", "rpc server is unavailable",
            "0x80010001", "-2147418111", "call was rejected by callee",
            "0x80010108", "-2147417848", "object invoked has disconnected",
        )
        if any(marker in message for marker in markers):
            return "transient_modify_com_failure"
        return None

    @staticmethod
    def _attempt_diagnostic(
        result: OperationResult[Any],
        *,
        attempt_number: int,
        retry_reason: str | None,
    ) -> AttemptDiagnostic:
        return AttemptDiagnostic(
            attempt=attempt_number,
            phase=result.phase,
            request_id=result.request_id,
            error_code=result.error.code if result.error else None,
            error_message=result.error.message if result.error else None,
            error_details=dict(result.error.details) if result.error else {},
            worker=result.diagnostics.worker_exit,
            excel_pid=result.diagnostics.excel_pid,
            cleanup=result.diagnostics.cleanup,
            dialog_count=len(result.diagnostics.dialog_events),
            elapsed_seconds=result.elapsed_seconds,
            retryable=retry_reason is not None,
            retry_reason=retry_reason,
        )

    @staticmethod
    def _with_attempts(
        result: OperationResult[Any],
        attempts: list[AttemptDiagnostic],
        *,
        started: float,
    ) -> OperationResult[Any]:
        diagnostics: Diagnostics = replace(
            result.diagnostics,
            attempts=tuple(attempts),
        )
        return replace(
            result,
            diagnostics=diagnostics,
            attempt_count=max(1, len(attempts)),
            elapsed_seconds=time.monotonic() - started,
        )
