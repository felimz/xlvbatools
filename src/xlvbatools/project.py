"""Primary public API for project-scoped Excel/VBA automation."""

from __future__ import annotations

from dataclasses import dataclass, replace
import os
from pathlib import Path
from typing import Any, Iterable, Mapping
import uuid

from xlvbatools.analysis.issue import VBAIssue
from xlvbatools.config.loader import load_config
from xlvbatools.config.schema import XlvbaConfig
from xlvbatools.errors import ConfigurationError
from xlvbatools.execution import Executor, IsolatedExecutor, Operation, OperationRequest
from xlvbatools.outputs import (
    ComponentDiff,
    ExtractionOutput,
    InjectionChange,
    InjectionOutput,
    MacroOutput,
    ModificationOutput,
    VBAComponent,
)
from xlvbatools.results import Artifact, ErrorInfo, InspectionOutput, OperationResult
from xlvbatools.snapshots import SnapshotService
from xlvbatools.workflow import (
    InspectStep,
    WorkflowOutput,
    WorkflowStep,
    WORKFLOW_SCHEMA_VERSION,
    _steps_to_worker,
    _validate_workflow_steps,
)


@dataclass(frozen=True)
class ProjectSettings:
    """Immutable paths and policies required by the public project API."""

    workbook: Path
    source: Path
    snapshots: Path
    disabled_lint_rules: tuple[str, ...] = ()
    backup_limit: int = 5
    snapshot_limit: int = 10

    def __post_init__(self) -> None:
        object.__setattr__(self, "workbook", self.workbook.resolve())
        object.__setattr__(self, "source", self.source.resolve())
        object.__setattr__(self, "snapshots", self.snapshots.resolve())
        if self.workbook.suffix.casefold() not in {".xlsm", ".xlsb", ".xls"}:
            raise ConfigurationError(
                f"Workbook must be .xlsm, .xlsb, or .xls: {self.workbook}"
            )
        if self.backup_limit < 0:
            raise ConfigurationError("backup_limit must be non-negative")
        if self.snapshot_limit < 1:
            raise ConfigurationError("snapshot_limit must be at least one")

    @classmethod
    def _from_config(cls, config: XlvbaConfig) -> "ProjectSettings":
        errors = config.validate()
        if errors:
            raise ConfigurationError("; ".join(errors))
        return cls(
            workbook=Path(config.workbook_path),
            source=Path(config.vba_source_path),
            snapshots=Path(config.snapshots_path),
            disabled_lint_rules=tuple(config.lint.disabled_rules),
            backup_limit=config.backups.limit,
            snapshot_limit=config.snapshots.rolling_limit,
        )


class Project:
    """A configured workbook and its VBA source tree.

    ``Project`` is the sole high-level entry point for Python consumers and
    project-specific wrappers.  Every Excel-backed method crosses the same
    typed executor boundary; COM objects and worker transport dictionaries are
    never exposed to callers.
    """

    def __init__(
        self,
        settings: ProjectSettings,
        *,
        executor: Executor | None = None,
    ) -> None:
        self.settings = settings
        self.executor: Executor = executor or IsolatedExecutor()

    @classmethod
    def from_config(
        cls,
        start_dir: str | os.PathLike[str] | None = None,
        *,
        executor: Executor | None = None,
    ) -> "Project":
        """Load the nearest ``xlvbatools.toml`` configuration."""
        config = load_config(os.fspath(start_dir) if start_dir is not None else None)
        return cls(ProjectSettings._from_config(config), executor=executor)

    @classmethod
    def open(
        cls,
        workbook: str | os.PathLike[str],
        *,
        source: str | os.PathLike[str] | None = None,
        executor: Executor | None = None,
    ) -> "Project":
        """Create a project from explicit paths without a TOML file."""
        workbook_path = Path(workbook).resolve()
        source_path = Path(source).resolve() if source else workbook_path.parent / "vba_source"
        return cls(
            ProjectSettings(
                workbook=workbook_path,
                source=source_path,
                snapshots=workbook_path.parent / "snapshots",
            ),
            executor=executor,
        )

    @property
    def workbook(self) -> Path:
        return self.settings.workbook

    @property
    def source(self) -> Path:
        return self.settings.source

    def execute(
        self,
        operation: Operation,
        arguments: Mapping[str, Any],
        *,
        timeout: float,
        retry_transient: bool = False,
    ) -> OperationResult[Any]:
        """Execute an advanced worker operation through the typed boundary."""
        return self.executor.execute(
            OperationRequest(
                operation=operation,
                arguments=arguments,
                timeout=timeout,
                retry_transient=retry_transient,
            )
        )

    def list_components(
        self, *, timeout: float = 60.0,
    ) -> OperationResult[tuple[VBAComponent, ...]]:
        result = self.execute(
            Operation.LIST_COMPONENTS,
            {"workbook_path": str(self.workbook)},
            timeout=timeout,
        )
        components = tuple(
            VBAComponent._from_mapping(item)
            for item in (result.data or ())
            if isinstance(item, Mapping)
        )
        return replace(result, data=components)

    def inspect(
        self,
        sheets: Iterable[str],
        *,
        output_dir: str | os.PathLike[str] = "screenshots",
        cell_range: str | None = None,
        include_data: bool = True,
        include_screenshots: bool = True,
        output_json: str | os.PathLike[str] | None = None,
        output_markdown: str | os.PathLike[str] | None = None,
        continue_on_render_error: bool = False,
        include_hidden_sheets: bool = False,
        timeout: float = 60.0,
    ) -> OperationResult[InspectionOutput]:
        requested_sheets = tuple(str(sheet) for sheet in sheets)
        if not requested_sheets:
            raise ValueError("at least one sheet is required")
        result = self.execute(
            Operation.INSPECT,
            {
                "workbook_path": str(self.workbook),
                "sheets": list(requested_sheets),
                "output_dir": os.path.abspath(os.fspath(output_dir)),
                "custom_range": cell_range,
                "include_data": include_data,
                "include_screenshots": include_screenshots,
                "output_json": str(Path(output_json).resolve()) if output_json else None,
                "output_md": str(Path(output_markdown).resolve()) if output_markdown else None,
                "continue_on_render_error": continue_on_render_error,
                "include_hidden_sheets": include_hidden_sheets,
            },
            timeout=timeout,
        )
        if result.data is None:
            return result
        payload = dict(result.data)
        screenshots = dict(payload.get("screenshots") or {})
        artifacts = tuple(
            Artifact(
                kind="screenshot",
                path=path,
                media_type="image/png",
                label=sheet,
                metadata={"sheet": sheet},
            )
            for sheet, path in screenshots.items()
            if isinstance(path, str)
            and path not in {"Not found", "Empty", "Hidden (skipped)"}
            and not path.startswith("Error:")
        )
        return replace(
            result,
            data=InspectionOutput(
                workbook_data=payload.get("workbook_data"),
                screenshots=screenshots,
            ),
            artifacts=artifacts,
            metadata={"sheets": requested_sheets, "cell_range": cell_range},
        )

    def run(
        self,
        macro: str,
        *,
        named_ranges: Mapping[str, Any] | None = None,
        timeout: float = 120.0,
        visible: bool = False,
        save: bool = True,
        strict_named_ranges: bool = True,
    ) -> OperationResult[MacroOutput]:
        if not macro.strip():
            raise ValueError("macro must be non-empty")
        result = self.execute(
            Operation.RUN,
            {
                "workbook_path": str(self.workbook),
                "macro_name": macro,
                "named_ranges": dict(named_ranges) if named_ranges else None,
                "timeout": timeout,
                "visible": visible,
                "save_on_exit": save,
                "strict_named_ranges": strict_named_ranges,
                "run_id": str(uuid.uuid4()),
            },
            timeout=timeout,
        )
        payload = result.data if isinstance(result.data, Mapping) else {}
        return replace(
            result,
            data=MacroOutput._from_mapping(macro, payload),
            metadata={"macro": macro},
        )

    def workflow(
        self,
        steps: Iterable[WorkflowStep],
        *,
        timeout: float = 240.0,
        visible: bool = False,
        save: bool = False,
    ) -> OperationResult[WorkflowOutput]:
        """Run ordered steps in one isolated worker and one Excel session."""
        if not isinstance(visible, bool):
            raise TypeError("visible must be boolean")
        if not isinstance(save, bool):
            raise TypeError("save must be boolean")
        validated = _validate_workflow_steps(steps)
        worker_steps = _steps_to_worker(validated)
        for step, step_payload in zip(validated, worker_steps):
            if isinstance(step, InspectStep):
                step_payload["output_dir"] = os.path.abspath(step.output_dir)
                step_payload["output_json"] = (
                    str(Path(step.output_json).resolve()) if step.output_json else None
                )
                step_payload["output_markdown"] = (
                    str(Path(step.output_markdown).resolve())
                    if step.output_markdown else None
                )
        result = self.execute(
            Operation.WORKFLOW,
            {
                "workbook_path": str(self.workbook),
                "steps": worker_steps,
                "visible": visible,
                "save_on_success": save,
                "workflow_schema_version": WORKFLOW_SCHEMA_VERSION,
            },
            timeout=timeout,
        )
        result_payload = (
            result.data
            if isinstance(result.data, Mapping)
            else {
                "workflow_schema_version": WORKFLOW_SCHEMA_VERSION,
                "steps": [],
                "failed_step_id": result.diagnostics.progress.get("step_id"),
                "save_requested": save,
                "saved": False,
            }
        )
        output = WorkflowOutput._from_mapping(result_payload)
        artifacts = tuple(
            artifact
            for step_result in output.steps
            for artifact in step_result.artifacts
        )
        return replace(
            result,
            data=output,
            artifacts=artifacts,
            metadata={
                "step_ids": tuple(step.id for step in validated),
                "save_requested": save,
                "visible": visible,
            },
        )

    def lint_source(
        self,
        source: str | os.PathLike[str] | None = None,
    ) -> OperationResult[tuple[VBAIssue, ...]]:
        from xlvbatools.analysis.preflight import lint_files

        issues = tuple(
            lint_files(
                os.fspath(source) if source is not None else str(self.source),
                disabled_rules=list(self.settings.disabled_lint_rules),
            )
        )
        errors = tuple(issue for issue in issues if issue.severity == "ERROR")
        return OperationResult(
            operation="lint_source",
            success=not errors,
            phase="complete",
            data=issues,
            error=(
                ErrorInfo(
                    message=f"Static analysis found {len(errors)} error(s)",
                    code="lint_failed",
                )
                if errors else None
            ),
            metadata={"error_count": len(errors), "issue_count": len(issues)},
        )

    def lint_workbook(
        self,
        *,
        compile_test: bool = True,
        timeout: float = 120.0,
    ) -> OperationResult[tuple[VBAIssue, ...]]:
        result = self.execute(
            Operation.LINT_WORKBOOK,
            {
                "workbook_path": str(self.workbook),
                "disabled_rules": list(self.settings.disabled_lint_rules),
                "compile_test": compile_test,
            },
            timeout=timeout,
        )
        issues = tuple(
            item if isinstance(item, VBAIssue) else VBAIssue(**item)
            for item in (result.data or ())
        )
        errors = tuple(issue for issue in issues if issue.severity == "ERROR")
        return replace(
            result,
            data=issues,
            metadata={
                "error_count": len(errors),
                "issue_count": len(issues),
                "compile_test": compile_test,
            },
        )

    def extract(
        self,
        *,
        output: str | os.PathLike[str] | None = None,
        component: str | None = None,
        timeout: float = 120.0,
    ) -> OperationResult[ExtractionOutput]:
        target = Path(output).resolve() if output else self.source
        result = self.execute(
            Operation.EXTRACT,
            {
                "workbook_path": str(self.workbook),
                "output_dir": str(target),
                "component": component,
            },
            timeout=timeout,
        )
        payload = result.data if isinstance(result.data, Mapping) else {}
        components: tuple[VBAComponent, ...]
        if component:
            components = (
                (VBAComponent._from_mapping(payload),) if payload else ()
            )
            workbook = self.workbook.name
            extracted_at = ""
        else:
            raw_components = payload.get("components") or ()
            components = tuple(
                VBAComponent._from_mapping(item)
                for item in raw_components
                if isinstance(item, Mapping)
            )
            workbook = str(payload.get("workbook") or self.workbook.name)
            extracted_at = str(payload.get("extracted_at") or "")
        return replace(
            result,
            data=ExtractionOutput(
                workbook=workbook,
                output_dir=str(target),
                extracted_at=extracted_at,
                components=components,
            ),
        )

    def inject(
        self,
        *,
        source: str | os.PathLike[str] | None = None,
        component: str | None = None,
        backup: bool = True,
        dry_run: bool = False,
        timeout: float = 120.0,
    ) -> OperationResult[InjectionOutput]:
        if component and dry_run:
            raise ValueError("dry_run cannot be combined with component injection")
        source_path = Path(source).resolve() if source else self.source
        result = self.execute(
            Operation.INJECT,
            {
                "workbook_path": str(self.workbook),
                "source_dir": str(source_path),
                "component": component,
                "backup": backup,
                "dry_run": dry_run,
                "backup_limit": self.settings.backup_limit,
            },
            timeout=timeout,
        )
        changes: tuple[InjectionChange, ...]
        if component:
            changes = (
                InjectionChange(
                    name=component,
                    status="injected" if bool(result.data) else "error",
                ),
            )
        else:
            changes = tuple(
                InjectionChange._from_mapping(item)
                for item in (result.data or ())
                if isinstance(item, Mapping)
            )
        return replace(
            result,
            data=InjectionOutput(
                changes=changes,
                dry_run=dry_run,
                backup_requested=backup,
            ),
        )

    def diff(
        self,
        *,
        source: str | os.PathLike[str] | None = None,
        component: str | None = None,
        timeout: float = 120.0,
    ) -> OperationResult[tuple[ComponentDiff, ...]]:
        source_path = Path(source).resolve() if source else self.source
        result = self.execute(
            Operation.DIFF,
            {
                "workbook_path": str(self.workbook),
                "source_dir": str(source_path),
                "component": component,
            },
            timeout=timeout,
        )
        if component:
            raw_diffs = (result.data,) if isinstance(result.data, Mapping) else ()
        else:
            raw_diffs = result.data or ()
        diffs = tuple(
            ComponentDiff._from_mapping(item)
            for item in raw_diffs
            if isinstance(item, Mapping)
        )
        return replace(result, data=diffs)

    def modify(
        self,
        *,
        sheet: str = "Sheet1",
        cell: str | None = None,
        value: Any = None,
        formula: str | None = None,
        name: str | None = None,
        refers_to: str | None = None,
        delete_name: bool = False,
        timeout: float = 120.0,
    ) -> OperationResult[ModificationOutput]:
        result = self.execute(
            Operation.MODIFY,
            {
                "workbook_path": str(self.workbook),
                "sheet": sheet,
                "cell": cell,
                "value": value,
                "formula": formula,
                "name": name,
                "refers_to": refers_to,
                "delete_name": delete_name,
            },
            timeout=timeout,
            retry_transient=True,
        )
        if delete_name:
            action = "delete_name"
        elif name and refers_to:
            action = "create_name"
        elif formula is not None:
            action = "set_formula"
        else:
            action = "set_value"
        return replace(
            result,
            data=ModificationOutput(
                applied=bool(result.data),
                sheet=sheet if cell else None,
                cell=cell,
                name=name,
                action=action,
            ),
        )

    def snapshots(self) -> SnapshotService:
        """Return the snapshot service bound to this project."""
        return SnapshotService(
            self.workbook,
            self.source,
            self.settings.snapshots,
            rolling_limit=self.settings.snapshot_limit,
        )
