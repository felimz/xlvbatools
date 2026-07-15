"""Typed public snapshot service for workbook and VBA source checkpoints."""

from __future__ import annotations

from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterator, Mapping

from xlvbatools.errors import SnapshotError, SnapshotNotFoundError
from xlvbatools.snapshot.manager import _SnapshotStore


@contextmanager
def _translate_store_errors(action: str) -> Iterator[None]:
    """Keep filesystem and corrupt-metadata failures behind the public hierarchy."""
    try:
        yield
    except SnapshotError:
        raise
    except (OSError, ValueError, KeyError, TypeError) as error:
        raise SnapshotError(f"Could not {action} snapshot data: {error}") from error


@dataclass(frozen=True)
class SnapshotGitInfo:
    """Git provenance captured when a snapshot is created."""

    branch: str
    commit: str
    dirty: bool

    @classmethod
    def _from_mapping(cls, value: Mapping[str, Any]) -> "SnapshotGitInfo":
        return cls(
            branch=str(value.get("branch") or "unknown"),
            commit=str(value.get("commit") or "unknown"),
            dirty=bool(value.get("dirty", False)),
        )


@dataclass(frozen=True)
class SnapshotRecord:
    """Immutable metadata describing one workbook/source checkpoint."""

    snapshot_id: str
    timestamp: str
    description: str
    workbook_file: str
    workbook_hash: str
    workbook_size_bytes: int
    vba_source_dir: str | None
    vba_hash: str
    milestone: bool
    git: SnapshotGitInfo | None = None

    @classmethod
    def _from_mapping(cls, value: Mapping[str, Any]) -> "SnapshotRecord":
        git_value = value.get("git")
        return cls(
            snapshot_id=str(value["snapshot_id"]),
            timestamp=str(value.get("timestamp") or ""),
            description=str(value.get("description") or ""),
            workbook_file=str(value.get("workbook_file") or ""),
            workbook_hash=str(value.get("workbook_hash") or ""),
            workbook_size_bytes=int(value.get("workbook_size_bytes") or 0),
            vba_source_dir=(
                str(value["vba_source_dir"])
                if value.get("vba_source_dir") is not None else None
            ),
            vba_hash=str(value.get("vba_hash") or ""),
            milestone=bool(value.get("milestone", False)),
            git=SnapshotGitInfo._from_mapping(git_value) if isinstance(git_value, Mapping) else None,
        )

    def to_dict(self) -> dict[str, object]:
        """Return JSON-serializable metadata without exposing mutable state."""
        return {
            "snapshot_id": self.snapshot_id,
            "timestamp": self.timestamp,
            "description": self.description,
            "workbook_file": self.workbook_file,
            "workbook_hash": self.workbook_hash,
            "workbook_size_bytes": self.workbook_size_bytes,
            "vba_source_dir": self.vba_source_dir,
            "vba_hash": self.vba_hash,
            "milestone": self.milestone,
            "git": (
                {
                    "branch": self.git.branch,
                    "commit": self.git.commit,
                    "dirty": self.git.dirty,
                }
                if self.git is not None
                else None
            ),
        }


class SnapshotService:
    """Project-bound, typed API for creating and restoring snapshots."""

    def __init__(
        self,
        workbook: str | Path,
        source: str | Path,
        snapshots: str | Path,
        *,
        rolling_limit: int = 10,
    ) -> None:
        self._store = _SnapshotStore(
            workbook_path=str(Path(workbook).resolve()),
            vba_source_dir=str(Path(source).resolve()),
            snapshots_dir=str(Path(snapshots).resolve()),
            rolling_limit=rolling_limit,
        )

    @staticmethod
    def _identifier(value: str | SnapshotRecord) -> str:
        return value.snapshot_id if isinstance(value, SnapshotRecord) else value

    def create(
        self,
        description: str = "",
        milestone: bool = False,
    ) -> SnapshotRecord:
        with _translate_store_errors("create"):
            snapshot_id = self._store.create(
                description=description,
                milestone=milestone,
            )
            record = self.info(snapshot_id)
            if record is None:
                raise SnapshotError(
                    f"Snapshot metadata was not persisted: {snapshot_id}"
                )
            return record

    def list(self) -> tuple[SnapshotRecord, ...]:
        """Return all snapshots in chronological order."""
        with _translate_store_errors("list"):
            return tuple(
                SnapshotRecord._from_mapping(item) for item in self._store.list()
            )

    def info(self, identifier: str | SnapshotRecord) -> SnapshotRecord | None:
        """Return one matching snapshot, or ``None`` when it does not exist."""
        with _translate_store_errors("read"):
            value = self._store.info(self._identifier(identifier))
            return SnapshotRecord._from_mapping(value) if value is not None else None

    def restore(
        self,
        identifier: str | SnapshotRecord,
        *,
        safety_snapshot: bool = True,
    ) -> SnapshotRecord:
        """Restore one snapshot or raise a typed public error."""
        with _translate_store_errors("restore"):
            resolved = self.info(identifier)
            if resolved is None:
                raise SnapshotNotFoundError(
                    f"Snapshot not found: {self._identifier(identifier)}"
                )
            if not self._store.restore(
                resolved.snapshot_id,
                safety_snapshot=safety_snapshot,
            ):
                raise SnapshotError(f"Snapshot restore failed: {resolved.snapshot_id}")
            return resolved

    def diff(self, identifier: str | SnapshotRecord) -> str:
        """Return the source diff between a snapshot and the current project."""
        with _translate_store_errors("diff"):
            resolved = self.info(identifier)
            if resolved is None:
                raise SnapshotNotFoundError(
                    f"Snapshot not found: {self._identifier(identifier)}"
                )
            return self._store.diff(resolved.snapshot_id)

    def prune(self, keep: int | None = None) -> int:
        """Remove old rolling snapshots and return the number removed."""
        if keep is not None and keep < 1:
            raise ValueError("keep must be at least one")
        with _translate_store_errors("prune"):
            return self._store.prune(keep=keep or self._store.rolling_limit)
