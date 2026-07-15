"""Machine-first serialization and optional terminal presentation helpers."""

from __future__ import annotations

import json
import sys
from collections.abc import Iterable
from typing import Any, NoReturn

from xlvbatools import Artifact, ErrorInfo, OperationResult


def is_machine(args: Any) -> bool:
    """Whether the stable JSON presentation is active."""
    return getattr(args, "output_format", "json") == "json"


def is_table(args: Any) -> bool:
    """Whether aligned terminal-table presentation was requested."""
    return getattr(args, "output_format", "json") == "table"


def print_result_json(result: OperationResult[Any]) -> None:
    """Write exactly one public result envelope to stdout."""
    print(json.dumps(result.to_dict(), indent=2, default=str))


def local_result(
    operation: str,
    data: Any = None,
    *,
    success: bool = True,
    error: Exception | str | None = None,
    code: str = "command_failed",
    metadata: dict[str, Any] | None = None,
    artifacts: tuple[Artifact, ...] = (),
) -> OperationResult[Any]:
    """Wrap non-worker CLI behavior in the public result contract."""
    error_info = None
    if error is not None:
        error_info = ErrorInfo(
            message=str(error),
            code=code,
            error_type=type(error).__name__ if isinstance(error, Exception) else None,
        )
    return OperationResult(
        operation=operation,
        success=success,
        phase="complete" if success else "failed",
        data=data,
        error=error_info,
        metadata=metadata or {},
        artifacts=artifacts,
    )


def print_table(headers: tuple[str, ...], rows: Iterable[Iterable[Any]]) -> None:
    """Render a dependency-free aligned table."""
    normalized = [
        tuple("" if value is None else str(value) for value in row)
        for row in rows
    ]
    widths = [len(header) for header in headers]
    for row in normalized:
        for index, value in enumerate(row):
            widths[index] = max(widths[index], len(value))
    print("  ".join(header.ljust(widths[index]) for index, header in enumerate(headers)))
    print("  ".join("-" * width for width in widths))
    for row in normalized:
        print("  ".join(value.ljust(widths[index]) for index, value in enumerate(row)))


def fail_result(
    args: Any,
    result: OperationResult[Any],
    fallback: str,
    *,
    exit_code: int = 1,
) -> NoReturn:
    """Emit a machine or presentation failure and terminate predictably."""
    if is_machine(args):
        print_result_json(result)
        raise SystemExit(exit_code)
    message = result.error.message if result.error is not None else fallback
    print(message, file=sys.stderr)
    raise SystemExit(exit_code)
