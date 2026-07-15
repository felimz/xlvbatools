"""Public exception hierarchy for wrapper-facing xlvbatools APIs."""

from __future__ import annotations

from typing import Any


class XlvbaError(RuntimeError):
    """Base class for errors intentionally exposed by xlvbatools."""


class ConfigurationError(XlvbaError):
    """Raised when project configuration is invalid or incomplete."""


class OperationFailedError(XlvbaError):
    """Raised by ``OperationResult.require_success`` for a failed operation."""

    def __init__(self, result: Any):
        self.result = result
        error = getattr(result, "error", None)
        message = getattr(error, "message", None) or (
            f"{getattr(result, 'operation', 'operation')} failed during "
            f"{getattr(result, 'phase', 'unknown')}"
        )
        super().__init__(message)


class HeadlessCleanupError(XlvbaError):
    """Raised when an operation did not produce a clean owned-Excel shutdown."""

    def __init__(self, result: Any, detail: str):
        self.result = result
        super().__init__(detail)


class TrustCenterError(XlvbaError):
    """Raised when Excel denies programmatic access to the VBA project."""


class SnapshotError(XlvbaError):
    """Raised when snapshot persistence or restoration fails."""


class SnapshotNotFoundError(SnapshotError):
    """Raised when a requested snapshot cannot be resolved."""
