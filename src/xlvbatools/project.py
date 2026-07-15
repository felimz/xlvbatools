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
from xlvbatools.results import Artifact, ErrorInfo, InspectionOutput, OperationResult


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

    def list_components(self, *, timeout: float = 60.0) -> OperationResult[Any]:
        return self.execute(
            Operation.LIST_COMPONENTS,
            {"workbook_path": str(self.workbook)},
            timeout=timeout,
        )

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
            return result  # type: ignore[return-value]
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
    ) -> OperationResult[dict[str, Any]]:
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
        return replace(result, metadata={"macro": macro})

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
    ) -> OperationResult[Any]:
        target = Path(output).resolve() if output else self.source
        return self.execute(
            Operation.EXTRACT,
            {
                "workbook_path": str(self.workbook),
                "output_dir": str(target),
                "component": component,
            },
            timeout=timeout,
        )

    def inject(
        self,
        *,
        source: str | os.PathLike[str] | None = None,
        component: str | None = None,
        backup: bool = True,
        dry_run: bool = False,
        timeout: float = 120.0,
    ) -> OperationResult[Any]:
        if component and dry_run:
            raise ValueError("dry_run cannot be combined with component injection")
        source_path = Path(source).resolve() if source else self.source
        return self.execute(
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

    def diff(
        self,
        *,
        source: str | os.PathLike[str] | None = None,
        component: str | None = None,
        timeout: float = 120.0,
    ) -> OperationResult[Any]:
        source_path = Path(source).resolve() if source else self.source
        return self.execute(
            Operation.DIFF,
            {
                "workbook_path": str(self.workbook),
                "source_dir": str(source_path),
                "component": component,
            },
            timeout=timeout,
        )

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
    ) -> OperationResult[Any]:
        return self.execute(
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

    def snapshots(self):
        """Return the snapshot service bound to this project."""
        from xlvbatools.snapshot.manager import SnapshotManager

        return SnapshotManager(
            str(self.workbook),
            str(self.source),
            str(self.settings.snapshots),
            rolling_limit=self.settings.snapshot_limit,
        )
