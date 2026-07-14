"""Config-bound high-level facade intended for project-specific wrappers."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any, Iterable, Optional

from xlvbatools.config.loader import load_config
from xlvbatools.config.schema import XlvbaConfig
from xlvbatools.errors import ConfigurationError
from xlvbatools.results import (
    Artifact,
    ErrorInfo,
    InspectionOutput,
    OperationResult,
)


class XlvbaProject:
    """A resolved project configuration with stable operation entry points.

    Existing function APIs remain supported. This facade gives consumer
    projects one import and one versioned result contract while the execution
    internals continue to evolve behind it.
    """

    def __init__(self, config: XlvbaConfig, *, validate: bool = True):
        if validate:
            errors = config.validate()
            if errors:
                raise ConfigurationError("; ".join(errors))
        self.config = config

    @classmethod
    def from_config(
        cls, start_dir: str | os.PathLike[str] | None = None,
    ) -> "XlvbaProject":
        """Load the nearest ``xlvbatools.toml`` and resolve paths from it."""
        return cls(load_config(os.fspath(start_dir) if start_dir is not None else None))

    @classmethod
    def for_workbook(
        cls,
        workbook_path: str | os.PathLike[str],
        *,
        vba_source: str | os.PathLike[str] | None = None,
    ) -> "XlvbaProject":
        """Create a project directly from explicit paths, without a TOML file."""
        workbook = Path(workbook_path).resolve()
        source = Path(vba_source).resolve() if vba_source else workbook.parent / "vba_source"
        return cls(XlvbaConfig(workbook=str(workbook), vba_source=str(source)))

    @property
    def workbook_path(self) -> str:
        return self.config.workbook_path

    @property
    def vba_source_path(self) -> str:
        return self.config.vba_source_path

    def inspect(
        self,
        sheets: Iterable[str],
        *,
        workbook_path: str | os.PathLike[str] | None = None,
        output_dir: str | os.PathLike[str] = "screenshots",
        cell_range: Optional[str] = None,
        include_data: bool = True,
        include_screenshots: bool = True,
        output_json: str | os.PathLike[str] | None = None,
        output_markdown: str | os.PathLike[str] | None = None,
        continue_on_render_error: bool = False,
        include_hidden_sheets: bool = False,
        timeout: float = 60.0,
    ) -> OperationResult[InspectionOutput]:
        """Inspect workbook data and screenshots in the isolated worker."""
        requested_sheets = [str(sheet) for sheet in sheets]
        try:
            from xlvbatools.workbook.dumper import inspect_workbook

            raw = inspect_workbook(
                os.fspath(workbook_path) if workbook_path else self.workbook_path,
                requested_sheets,
                output_dir=os.fspath(output_dir),
                custom_range=cell_range,
                include_data=include_data,
                include_screenshots=include_screenshots,
                output_json=os.fspath(output_json) if output_json else None,
                output_md=os.fspath(output_markdown) if output_markdown else None,
                continue_on_render_error=continue_on_render_error,
                include_hidden_sheets=include_hidden_sheets,
                timeout_seconds=timeout,
            )
            screenshots = dict(raw.get("screenshots") or {})
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
                and path not in ("Not found", "Empty", "Hidden (skipped)")
                and not path.startswith("Error:")
            )
            output = InspectionOutput(
                workbook_data=raw.get("data"), screenshots=screenshots,
            )
            return OperationResult.from_legacy(
                "inspect",
                raw,
                data=output,
                artifacts=artifacts,
                metadata={"sheets": requested_sheets, "cell_range": cell_range},
            )
        except Exception as error:
            return OperationResult.failed("inspect", error)

    def run_macro(
        self,
        macro_name: str,
        *,
        workbook_path: str | os.PathLike[str] | None = None,
        named_ranges: Optional[dict] = None,
        timeout: float = 120.0,
        visible: bool = False,
        save_on_exit: bool = True,
        strict_named_ranges: bool = True,
    ) -> OperationResult[dict[str, Any]]:
        """Run a macro through the parent-enforced isolated worker."""
        try:
            from xlvbatools.macro.runner import run_macro

            raw = run_macro(
                os.fspath(workbook_path) if workbook_path else self.workbook_path,
                macro_name,
                named_ranges=named_ranges,
                timeout=timeout,
                visible=visible,
                save_on_exit=save_on_exit,
                strict_named_ranges=strict_named_ranges,
            )
            return OperationResult.from_legacy(
                "run_macro", raw, data=dict(raw), metadata={"macro": macro_name},
            )
        except Exception as error:
            return OperationResult.failed("run_macro", error)

    def lint(
        self,
        source: str | os.PathLike[str] | None = None,
        *,
        workbook: bool = False,
        compile_test: bool = True,
    ) -> OperationResult[tuple[Any, ...]]:
        """Lint project source offline, or lint/compile the workbook via COM."""
        try:
            if workbook:
                from xlvbatools.analysis.preflight import lint_workbook

                issues = lint_workbook(
                    self.workbook_path,
                    disabled_rules=self.config.lint.disabled_rules,
                    compile_test=compile_test,
                )
            else:
                from xlvbatools.analysis.preflight import lint_files

                issues = lint_files(
                    os.fspath(source) if source else self.vba_source_path,
                    disabled_rules=self.config.lint.disabled_rules,
                )
            errors = [issue for issue in issues if issue.severity == "ERROR"]
            return OperationResult(
                operation="lint",
                success=not errors,
                phase="complete",
                data=tuple(issues),
                error=(
                    ErrorInfo(
                        message=f"Static analysis found {len(errors)} error(s)",
                        code="lint_failed",
                    )
                    if errors else None
                ),
                metadata={"error_count": len(errors), "issue_count": len(issues)},
            )
        except Exception as error:
            return OperationResult.failed("lint", error)

    def extract(
        self,
        *,
        output_dir: str | os.PathLike[str] | None = None,
        component: Optional[str] = None,
    ) -> OperationResult[Any]:
        """Extract one or all VBA components using resolved project paths."""
        target = os.fspath(output_dir) if output_dir else self.vba_source_path
        try:
            if component:
                from xlvbatools.vba.extractor import extract_component

                data = extract_component(self.workbook_path, component, target)
                if data is None:
                    return OperationResult(
                        operation="extract", success=False, phase="extract",
                        error=ErrorInfo(
                            f"Component not found: {component}", code="not_found",
                        ),
                    )
            else:
                from xlvbatools.vba.extractor import extract_all

                data = extract_all(self.workbook_path, target)
            return OperationResult(
                operation="extract", success=True, phase="complete", data=data,
            )
        except Exception as error:
            return OperationResult.failed("extract", error)

    def inject(
        self,
        *,
        source_dir: str | os.PathLike[str] | None = None,
        component: Optional[str] = None,
        backup: bool = True,
        dry_run: bool = False,
    ) -> OperationResult[Any]:
        """Inject one or all VBA components using resolved project paths."""
        source = os.fspath(source_dir) if source_dir else self.vba_source_path
        try:
            if component:
                if dry_run:
                    raise ValueError("dry_run is supported only for project injection")
                from xlvbatools.vba.injector import inject_component

                success = inject_component(
                    self.workbook_path, source, component, backup=backup,
                )
                data: Any = {"component": component, "injected": success}
            else:
                from xlvbatools.vba.injector import inject_all

                data = inject_all(
                    self.workbook_path,
                    source,
                    backup=backup,
                    dry_run=dry_run,
                    backup_limit=self.config.backups.limit,
                )
                success = all(item.get("status") != "error" for item in data)
            return OperationResult(
                operation="inject",
                success=success,
                phase="complete" if success else "inject",
                data=data,
                error=(
                    None if success else ErrorInfo(
                        "One or more VBA components failed to inject",
                        code="injection_failed",
                    )
                ),
            )
        except Exception as error:
            return OperationResult.failed("inject", error)

    def diff(
        self,
        *,
        source_dir: str | os.PathLike[str] | None = None,
        component: Optional[str] = None,
    ) -> OperationResult[Any]:
        """Compare live workbook VBA with resolved project source."""
        source = os.fspath(source_dir) if source_dir else self.vba_source_path
        try:
            if component:
                from xlvbatools.vba.differ import diff_component

                data = diff_component(self.workbook_path, source, component)
                if data is None:
                    return OperationResult(
                        operation="diff", success=False, phase="diff",
                        error=ErrorInfo(
                            f"Component not found: {component}", code="not_found",
                        ),
                    )
            else:
                from xlvbatools.vba.differ import diff_all

                data = diff_all(self.workbook_path, source)
            return OperationResult(
                operation="diff", success=True, phase="complete", data=data,
            )
        except Exception as error:
            return OperationResult.failed("diff", error)

    def modify(
        self,
        *,
        sheet: str = "Sheet1",
        cell: Optional[str] = None,
        value: Any = None,
        formula: Optional[str] = None,
        name: Optional[str] = None,
        refers_to: Optional[str] = None,
        delete_name: bool = False,
    ) -> OperationResult[dict[str, Any]]:
        """Modify a cell or named range through the compatibility implementation."""
        try:
            from xlvbatools.workbook.modifier import modify_cell

            success = modify_cell(
                self.workbook_path,
                sheet=sheet,
                cell=cell,
                value=value,
                formula=formula,
                name=name,
                refers_to=refers_to,
                delete_name=delete_name,
            )
            return OperationResult(
                operation="modify",
                success=success,
                phase="complete" if success else "modify",
                data={"sheet": sheet, "cell": cell, "name": name},
                error=(
                    None if success else ErrorInfo(
                        "Workbook modification failed", code="modification_failed",
                    )
                ),
            )
        except Exception as error:
            return OperationResult.failed("modify", error)

    def snapshot_manager(self):
        """Return a SnapshotManager bound to the resolved project paths."""
        from xlvbatools.snapshot.manager import SnapshotManager

        return SnapshotManager(
            self.workbook_path,
            self.vba_source_path,
            self.config.snapshots_path,
            rolling_limit=self.config.snapshots.rolling_limit,
        )
