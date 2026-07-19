"""Typed public contracts for one-session Excel workflows."""

from __future__ import annotations

from dataclasses import dataclass, field
import math
import re
from types import MappingProxyType
from typing import Any, Iterable, Mapping, Sequence, TypeAlias

from xlvbatools.outputs import FrozenJsonValue, MacroOutput
from xlvbatools.results import Artifact, ErrorInfo, InspectionOutput


WORKFLOW_SCHEMA_VERSION = "1.0"
_STEP_ID_RE = re.compile(r"^[A-Za-z][A-Za-z0-9_.-]{0,63}$")
_RANGE_RE = re.compile(
    r"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)"
    r"(?::\$?([A-Za-z]{1,3})\$?([1-9][0-9]*))?$"
)
_MAX_EXCEL_ROW = 1_048_576
_MAX_EXCEL_COLUMN = 16_384


def _validate_step_id(value: str) -> str:
    if not isinstance(value, str):
        raise TypeError("workflow step id must be a string")
    identifier = value.strip()
    if not _STEP_ID_RE.fullmatch(identifier):
        raise ValueError(
            "workflow step id must start with a letter and contain at most 64 "
            "letters, digits, dots, underscores, or hyphens"
        )
    return identifier


def _excel_column_number(column: str) -> int:
    value = 0
    for character in column.upper():
        value = value * 26 + ord(character) - ord("A") + 1
    return value


def _validate_a1_range(value: str) -> str:
    if not isinstance(value, str):
        raise TypeError("A1 range address must be a string")
    address = value.strip()
    match = _RANGE_RE.fullmatch(address)
    if match is None:
        raise ValueError(f"invalid A1 range address: {value!r}")
    first_column, first_row, last_column, last_row = match.groups()
    cells = [(first_column, first_row)]
    if last_column is not None and last_row is not None:
        cells.append((last_column, last_row))
    if any(
        _excel_column_number(column) > _MAX_EXCEL_COLUMN
        or int(row) > _MAX_EXCEL_ROW
        for column, row in cells
    ):
        raise ValueError(f"A1 range address is outside Excel limits: {value!r}")
    return address


def _validate_json_scalar(value: Any, *, label: str) -> None:
    if isinstance(value, float) and not math.isfinite(value):
        raise ValueError(f"{label} must not contain NaN or infinite values")


def _freeze_range_value(value: Any) -> FrozenJsonValue:
    if value is None or isinstance(value, (str, int, float, bool)):
        _validate_json_scalar(value, label="range values")
        return value
    if not isinstance(value, (list, tuple)) or not value:
        raise TypeError("range values must be a JSON scalar or a non-empty 2D sequence")
    rows: list[tuple[FrozenJsonValue, ...]] = []
    width: int | None = None
    for row in value:
        if not isinstance(row, (list, tuple)) or not row:
            raise TypeError("range values must be a non-empty rectangular 2D sequence")
        frozen_row: list[FrozenJsonValue] = []
        for item in row:
            if item is not None and not isinstance(item, (str, int, float, bool)):
                raise TypeError("range cells must contain only JSON scalar values")
            _validate_json_scalar(item, label="range values")
            frozen_row.append(item)
        if width is None:
            width = len(frozen_row)
        elif len(frozen_row) != width:
            raise ValueError("range values must be rectangular")
        rows.append(tuple(frozen_row))
    return tuple(rows)


def _freeze_workflow_json(value: Any) -> FrozenJsonValue:
    if value is None or isinstance(value, (str, int, float, bool)):
        _validate_json_scalar(value, label="workflow values")
        return value
    if isinstance(value, Mapping):
        frozen: dict[str, FrozenJsonValue] = {}
        for key, item in value.items():
            if not isinstance(key, str):
                raise TypeError("workflow JSON object keys must be strings")
            frozen[key] = _freeze_workflow_json(item)
        return MappingProxyType(frozen)
    if isinstance(value, (list, tuple)):
        return tuple(_freeze_workflow_json(item) for item in value)
    raise TypeError(
        "workflow values must contain only JSON mappings, sequences, and scalars"
    )


@dataclass(frozen=True)
class MacroStep:
    """Run one VBA macro in the workflow's existing workbook session."""

    id: str
    macro: str
    named_ranges: Mapping[str, FrozenJsonValue] = field(default_factory=dict)
    strict_named_ranges: bool = True

    def __post_init__(self) -> None:
        object.__setattr__(self, "id", _validate_step_id(self.id))
        if not isinstance(self.macro, str):
            raise TypeError("macro must be a string")
        macro = self.macro.strip()
        if not macro:
            raise ValueError("macro must be non-empty")
        object.__setattr__(self, "macro", macro)
        if not isinstance(self.named_ranges, Mapping):
            raise TypeError("named_ranges must be a mapping")
        if not isinstance(self.strict_named_ranges, bool):
            raise TypeError("strict_named_ranges must be boolean")
        names: dict[str, FrozenJsonValue] = {}
        seen: set[str] = set()
        for raw_name, raw_value in self.named_ranges.items():
            if not isinstance(raw_name, str):
                raise TypeError("named-range names must be strings")
            name = raw_name.strip()
            if not name:
                raise ValueError("named-range names must be non-empty")
            normalized = name.casefold()
            if normalized in seen:
                raise ValueError(f"duplicate named-range assignment: {name!r}")
            names[name] = _freeze_workflow_json(raw_value)
            seen.add(normalized)
        object.__setattr__(self, "named_ranges", MappingProxyType(names))


@dataclass(frozen=True)
class ModifyStep:
    """Write one or more ranges on a sheet without opening another workbook."""

    id: str
    sheet: str
    values: Mapping[str, FrozenJsonValue]
    calculate: bool = False

    def __post_init__(self) -> None:
        object.__setattr__(self, "id", _validate_step_id(self.id))
        if not isinstance(self.sheet, str):
            raise TypeError("sheet must be a string")
        sheet = self.sheet.strip()
        if not sheet:
            raise ValueError("sheet must be non-empty")
        if not isinstance(self.values, Mapping):
            raise TypeError("values must be a mapping of A1 ranges to values")
        if not isinstance(self.calculate, bool):
            raise TypeError("calculate must be boolean")
        if not self.values:
            raise ValueError("ModifyStep requires at least one range assignment")
        assignments: dict[str, FrozenJsonValue] = {}
        seen: set[str] = set()
        for raw_address, raw_value in self.values.items():
            address = _validate_a1_range(raw_address)
            normalized = address.replace("$", "").casefold()
            if normalized in seen:
                raise ValueError(f"duplicate range assignment: {address!r}")
            assignments[address] = _freeze_range_value(raw_value)
            seen.add(normalized)
        object.__setattr__(self, "sheet", sheet)
        object.__setattr__(self, "values", MappingProxyType(assignments))


@dataclass(frozen=True)
class InspectStep:
    """Inspect current workbook state without creating another Excel session."""

    id: str
    sheets: tuple[str, ...]
    output_dir: str = "screenshots"
    cell_range: str | None = None
    include_data: bool = True
    include_screenshots: bool = False
    output_json: str | None = None
    output_markdown: str | None = None
    continue_on_render_error: bool = False
    include_hidden_sheets: bool = False

    def __post_init__(self) -> None:
        object.__setattr__(self, "id", _validate_step_id(self.id))
        if isinstance(self.sheets, str) or not isinstance(self.sheets, Sequence):
            raise TypeError("sheets must be a sequence of worksheet names")
        for name in (
            "include_data", "include_screenshots", "continue_on_render_error",
            "include_hidden_sheets",
        ):
            if not isinstance(getattr(self, name), bool):
                raise TypeError(f"{name} must be boolean")
        if not self.sheets:
            raise ValueError("InspectStep requires at least one non-empty sheet")
        if any(not isinstance(sheet, str) for sheet in self.sheets):
            raise TypeError("worksheet names must be strings")
        sheets = tuple(sheet.strip() for sheet in self.sheets)
        if any(not sheet for sheet in sheets):
            raise ValueError("InspectStep requires at least one non-empty sheet")
        if len({sheet.casefold() for sheet in sheets}) != len(sheets):
            raise ValueError("InspectStep sheet names must be unique")
        if not self.include_data and not self.include_screenshots:
            raise ValueError("InspectStep must include data, screenshots, or both")
        if not isinstance(self.output_dir, str):
            raise TypeError("output_dir must be a string")
        if not self.output_dir.strip():
            raise ValueError("output_dir must be non-empty")
        cell_range = (
            _validate_a1_range(self.cell_range)
            if self.cell_range is not None else None
        )
        for name in ("output_json", "output_markdown"):
            output_path = getattr(self, name)
            if output_path is not None:
                if not isinstance(output_path, str):
                    raise TypeError(f"{name} must be a string or None")
                if not output_path.strip():
                    raise ValueError(f"{name} must be non-empty when provided")
                object.__setattr__(self, name, output_path.strip())
        object.__setattr__(self, "sheets", sheets)
        object.__setattr__(self, "output_dir", self.output_dir.strip())
        object.__setattr__(self, "cell_range", cell_range)


WorkflowStep: TypeAlias = MacroStep | ModifyStep | InspectStep


@dataclass(frozen=True)
class RangeWriteResult:
    """One range write completed by a workflow modification step."""

    sheet: str
    range: str
    rows: int
    columns: int

    @classmethod
    def _from_mapping(cls, value: Mapping[str, Any]) -> "RangeWriteResult":
        return cls(
            sheet=str(value.get("sheet") or ""),
            range=str(value.get("range") or ""),
            rows=int(value.get("rows") or 0),
            columns=int(value.get("columns") or 0),
        )


@dataclass(frozen=True)
class ModifyStepOutput:
    """Typed output for a workflow's multi-range modification step."""

    applied: bool
    writes: tuple[RangeWriteResult, ...] = ()
    calculated: bool = False

    @classmethod
    def _from_mapping(cls, value: Mapping[str, Any]) -> "ModifyStepOutput":
        return cls(
            applied=bool(value.get("applied", False)),
            writes=tuple(
                RangeWriteResult._from_mapping(item)
                for item in value.get("writes") or ()
                if isinstance(item, Mapping)
            ),
            calculated=bool(value.get("calculated", False)),
        )


WorkflowStepData: TypeAlias = MacroOutput | ModifyStepOutput | InspectionOutput


@dataclass(frozen=True)
class WorkflowStepResult:
    """Outcome and evidence for one ordered workflow step."""

    id: str
    kind: str
    status: str
    phase: str
    elapsed_seconds: float | None = None
    data: WorkflowStepData | None = None
    error: ErrorInfo | None = None
    dialog_events: tuple[Mapping[str, Any], ...] = ()
    artifacts: tuple[Artifact, ...] = ()

    @classmethod
    def _from_mapping(cls, value: Mapping[str, Any]) -> "WorkflowStepResult":
        kind = str(value.get("kind") or "unknown")
        raw_data = value.get("data")
        data: WorkflowStepData | None = None
        artifacts: list[Artifact] = []
        if isinstance(raw_data, Mapping):
            if kind == "macro":
                data = MacroOutput._from_mapping(str(raw_data.get("macro") or ""), raw_data)
            elif kind == "modify":
                data = ModifyStepOutput._from_mapping(raw_data)
            elif kind == "inspect":
                screenshots = dict(raw_data.get("screenshots") or {})
                data = InspectionOutput(
                    workbook_data=raw_data.get("workbook_data"),
                    screenshots=screenshots,
                )
                artifacts.extend(
                    Artifact(
                        kind="screenshot",
                        path=path,
                        media_type="image/png",
                        label=sheet,
                        metadata={"sheet": sheet, "step_id": value.get("id")},
                    )
                    for sheet, path in screenshots.items()
                    if isinstance(path, str)
                    and path not in {"Not found", "Empty", "Hidden (skipped)"}
                    and not path.startswith("Error:")
                )
        raw_error = value.get("error")
        error = None
        if isinstance(raw_error, Mapping):
            error = ErrorInfo(
                message=str(raw_error.get("message") or "Workflow step failed"),
                code=str(raw_error.get("code") or "workflow_step_failed"),
                error_type=(
                    str(raw_error["error_type"])
                    if raw_error.get("error_type") is not None else None
                ),
                details=dict(raw_error.get("details") or {}),
            )
        return cls(
            id=str(value.get("id") or ""),
            kind=kind,
            status=str(value.get("status") or "unknown"),
            phase=str(value.get("phase") or "unknown"),
            elapsed_seconds=(
                float(value["elapsed_seconds"])
                if value.get("elapsed_seconds") is not None else None
            ),
            data=data,
            error=error,
            dialog_events=tuple(value.get("dialog_events") or ()),
            artifacts=tuple(artifacts),
        )


@dataclass(frozen=True)
class WorkflowOutput:
    """Typed ordered output from one isolated shared-session workflow."""

    steps: tuple[WorkflowStepResult, ...]
    failed_step_id: str | None = None
    save_requested: bool = False
    saved: bool = False
    workflow_schema_version: str = WORKFLOW_SCHEMA_VERSION

    @classmethod
    def _from_mapping(cls, value: Mapping[str, Any]) -> "WorkflowOutput":
        return cls(
            steps=tuple(
                WorkflowStepResult._from_mapping(item)
                for item in value.get("steps") or ()
                if isinstance(item, Mapping)
            ),
            failed_step_id=(
                str(value["failed_step_id"])
                if value.get("failed_step_id") is not None else None
            ),
            save_requested=bool(value.get("save_requested", False)),
            saved=bool(value.get("saved", False)),
            workflow_schema_version=str(
                value.get("workflow_schema_version") or WORKFLOW_SCHEMA_VERSION
            ),
        )

    @property
    def by_id(self) -> Mapping[str, WorkflowStepResult]:
        """Return an immutable ID lookup while preserving serialized step order."""
        return MappingProxyType({step.id: step for step in self.steps})

    def step(self, identifier: str) -> WorkflowStepResult:
        """Return one step by ID or raise ``KeyError``."""
        return self.by_id[identifier]


def _validate_workflow_steps(steps: Iterable[WorkflowStep]) -> tuple[WorkflowStep, ...]:
    validated = tuple(steps)
    if not validated:
        raise ValueError("workflow requires at least one step")
    if len(validated) > 100:
        raise ValueError("workflow supports at most 100 steps")
    supported = (MacroStep, ModifyStep, InspectStep)
    if any(not isinstance(step, supported) for step in validated):
        raise TypeError("workflow steps must be MacroStep, ModifyStep, or InspectStep")
    ids = [step.id.casefold() for step in validated]
    if len(set(ids)) != len(ids):
        raise ValueError("workflow step ids must be unique (case-insensitive)")
    return validated


def _steps_to_worker(steps: Sequence[WorkflowStep]) -> list[dict[str, Any]]:
    payload: list[dict[str, Any]] = []
    for step in steps:
        if isinstance(step, MacroStep):
            payload.append({
                "id": step.id,
                "kind": "macro",
                "macro": step.macro,
                "named_ranges": dict(step.named_ranges),
                "strict_named_ranges": step.strict_named_ranges,
            })
        elif isinstance(step, ModifyStep):
            payload.append({
                "id": step.id,
                "kind": "modify",
                "sheet": step.sheet,
                "values": dict(step.values),
                "calculate": step.calculate,
            })
        else:
            payload.append({
                "id": step.id,
                "kind": "inspect",
                "sheets": list(step.sheets),
                "output_dir": step.output_dir,
                "cell_range": step.cell_range,
                "include_data": step.include_data,
                "include_screenshots": step.include_screenshots,
                "output_json": step.output_json,
                "output_markdown": step.output_markdown,
                "continue_on_render_error": step.continue_on_render_error,
                "include_hidden_sheets": step.include_hidden_sheets,
            })
    return payload


def _steps_from_payload(value: Any) -> tuple[WorkflowStep, ...]:
    if not isinstance(value, list):
        raise ValueError("workflow file must contain a 'steps' JSON array")
    steps: list[WorkflowStep] = []
    for index, item in enumerate(value):
        if not isinstance(item, Mapping):
            raise ValueError(f"workflow step {index + 1} must be a JSON object")
        raw_kind = item.get("kind")
        if not isinstance(raw_kind, str):
            raise TypeError(f"workflow step {index + 1} kind must be a string")
        kind = raw_kind.casefold()
        step_id = _require_string_field(item, "id", index=index)
        if kind == "macro":
            _reject_unknown_fields(
                item,
                {"id", "kind", "macro", "named_ranges", "strict_named_ranges"},
                index=index,
            )
            raw_names = item["named_ranges"] if "named_ranges" in item else {}
            if not isinstance(raw_names, Mapping):
                raise TypeError(f"workflow step {index + 1} named_ranges must be an object")
            steps.append(MacroStep(
                id=step_id,
                macro=_require_string_field(item, "macro", index=index),
                named_ranges=dict(raw_names),
                strict_named_ranges=item.get("strict_named_ranges", True),
            ))
        elif kind == "modify":
            _reject_unknown_fields(
                item,
                {"id", "kind", "sheet", "values", "calculate"},
                index=index,
            )
            raw_values = item["values"] if "values" in item else {}
            if not isinstance(raw_values, Mapping):
                raise TypeError(f"workflow step {index + 1} values must be an object")
            steps.append(ModifyStep(
                id=step_id,
                sheet=_require_string_field(item, "sheet", index=index),
                values=dict(raw_values),
                calculate=item.get("calculate", False),
            ))
        elif kind == "inspect":
            _reject_unknown_fields(
                item,
                {
                    "id", "kind", "sheets", "output_dir", "cell_range",
                    "include_data", "include_screenshots", "output_json",
                    "output_markdown", "continue_on_render_error",
                    "include_hidden_sheets",
                },
                index=index,
            )
            raw_sheets = item["sheets"] if "sheets" in item else ()
            if isinstance(raw_sheets, str) or not isinstance(raw_sheets, (list, tuple)):
                raise TypeError(f"workflow step {index + 1} sheets must be an array")
            steps.append(InspectStep(
                id=step_id,
                sheets=tuple(raw_sheets),
                output_dir=(
                    item["output_dir"] if "output_dir" in item else "screenshots"
                ),
                cell_range=item.get("cell_range"),
                include_data=item.get("include_data", True),
                include_screenshots=item.get("include_screenshots", False),
                output_json=item.get("output_json"),
                output_markdown=item.get("output_markdown"),
                continue_on_render_error=item.get("continue_on_render_error", False),
                include_hidden_sheets=item.get("include_hidden_sheets", False),
            ))
        else:
            raise ValueError(
                f"workflow step {index + 1} has unsupported kind {kind!r}; "
                "choose macro, modify, or inspect"
            )
    return _validate_workflow_steps(steps)


def _require_string_field(
    value: Mapping[str, Any],
    field_name: str,
    *,
    index: int,
) -> str:
    field_value = value.get(field_name)
    if not isinstance(field_value, str):
        raise TypeError(
            f"workflow step {index + 1} {field_name} must be a string"
        )
    return field_value


def _reject_unknown_fields(
    value: Mapping[str, Any],
    allowed: set[str],
    *,
    index: int,
) -> None:
    unexpected = set(value) - allowed
    if unexpected:
        raise ValueError(
            f"workflow step {index + 1} has unknown fields: "
            + ", ".join(sorted(unexpected))
        )
