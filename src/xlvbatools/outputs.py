"""Operation-specific immutable output models for the public Project API."""

from __future__ import annotations

from dataclasses import dataclass, field
from types import MappingProxyType
from typing import Any, Mapping, TypeAlias


JsonScalar: TypeAlias = str | int | float | bool | None
FrozenJsonValue: TypeAlias = (
    JsonScalar | tuple["FrozenJsonValue", ...] | Mapping[str, "FrozenJsonValue"]
)


def _freeze_json(value: Any) -> FrozenJsonValue:
    if isinstance(value, Mapping):
        return MappingProxyType(
            {str(key): _freeze_json(item) for key, item in value.items()}
        )
    if isinstance(value, (list, tuple)):
        return tuple(_freeze_json(item) for item in value)
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    return str(value)


@dataclass(frozen=True)
class VBAComponent:
    """One VBA component as listed or extracted from a workbook."""

    name: str
    type_code: int
    type_name: str
    line_count: int = 0
    file: str | None = None
    sha256: str | None = None

    @classmethod
    def _from_mapping(cls, value: Mapping[str, Any]) -> "VBAComponent":
        return cls(
            name=str(value.get("name") or ""),
            type_code=int(value.get("type_code") or 0),
            type_name=str(value.get("type_name") or "unknown"),
            line_count=int(value.get("line_count") or 0),
            file=str(value["file"]) if value.get("file") is not None else None,
            sha256=(
                str(value["sha256"]) if value.get("sha256") is not None else None
            ),
        )


@dataclass(frozen=True)
class ExtractionOutput:
    """Normalized result of whole-project or single-component extraction."""

    workbook: str
    output_dir: str
    extracted_at: str
    components: tuple[VBAComponent, ...] = ()


@dataclass(frozen=True)
class InjectionChange:
    """Outcome for one component considered during injection."""

    name: str
    status: str
    file: str | None = None
    action: str | None = None
    reason: str | None = None
    error: str | None = None

    @classmethod
    def _from_mapping(cls, value: Mapping[str, Any]) -> "InjectionChange":
        def optional(key: str) -> str | None:
            item = value.get(key)
            return str(item) if item is not None else None

        return cls(
            name=str(value.get("name") or ""),
            status=str(value.get("status") or "unknown"),
            file=optional("file"),
            action=optional("action"),
            reason=optional("reason"),
            error=optional("error"),
        )


@dataclass(frozen=True)
class InjectionOutput:
    """Normalized injection report."""

    changes: tuple[InjectionChange, ...] = ()
    dry_run: bool = False
    backup_requested: bool = True


@dataclass(frozen=True)
class ComponentDiff:
    """Comparison of one workbook component against its source file."""

    name: str
    status: str
    lines_added: int = 0
    lines_removed: int = 0
    unified_diff: str | None = None

    @classmethod
    def _from_mapping(cls, value: Mapping[str, Any]) -> "ComponentDiff":
        return cls(
            name=str(value.get("name") or ""),
            status=str(value.get("status") or "unknown"),
            lines_added=int(value.get("lines_added") or 0),
            lines_removed=int(value.get("lines_removed") or 0),
            unified_diff=(
                str(value["unified_diff"])
                if value.get("unified_diff") is not None
                else None
            ),
        )


@dataclass(frozen=True)
class ModificationOutput:
    """Description of a requested workbook modification."""

    applied: bool
    sheet: str | None = None
    cell: str | None = None
    name: str | None = None
    action: str = "modify"


@dataclass(frozen=True)
class MacroOutput:
    """Stable macro-run identity plus extensible JSON-compatible details."""

    macro: str
    run_id: str | None = None
    excel_pid: int | None = None
    details: Mapping[str, FrozenJsonValue] = field(default_factory=dict)

    @classmethod
    def _from_mapping(
        cls,
        macro: str,
        value: Mapping[str, Any],
    ) -> "MacroOutput":
        known = {"macro", "run_id", "excel_pid"}
        details = {
            str(key): _freeze_json(item)
            for key, item in value.items()
            if key not in known
        }
        return cls(
            macro=str(value.get("macro") or macro),
            run_id=str(value["run_id"]) if value.get("run_id") is not None else None,
            excel_pid=(
                int(value["excel_pid"])
                if value.get("excel_pid") is not None
                else None
            ),
            details=MappingProxyType(details),
        )
