"""Typed execution boundary for all isolated Excel operations."""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
import subprocess
from types import MappingProxyType
from typing import Any, Mapping, Protocol

from xlvbatools.core.protocol import WORKER_PROTOCOL_VERSION
from xlvbatools.results import OperationResult


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
        object.__setattr__(self, "arguments", MappingProxyType(dict(self.arguments)))


class Executor(Protocol):
    """Dependency-injection boundary used by ``Project`` and tests."""

    def execute(self, request: OperationRequest) -> OperationResult[Any]: ...


class IsolatedExecutor:
    """Execute one operation per child process and return a typed result."""

    def execute(self, request: OperationRequest) -> OperationResult[Any]:
        from xlvbatools.core.worker import execute_worker_request

        try:
            response = execute_worker_request(
                request.operation.value,
                request.arguments,
                timeout=request.timeout,
                retry_transient=request.retry_transient,
            )
        except (OSError, subprocess.SubprocessError) as error:
            return OperationResult.failed(
                request.operation.value,
                error,
                phase="transport",
                code="worker_start_failed",
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
