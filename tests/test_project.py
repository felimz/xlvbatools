"""Tests for the v1 Project API and typed executor boundary."""

from __future__ import annotations

from dataclasses import replace
import json

import pytest

from xlvbatools import (
    AttemptDiagnostic,
    CleanupReport,
    Diagnostics,
    Operation,
    OperationResult,
    Project,
)
from xlvbatools.results import ErrorInfo


class RecordingExecutor:
    def __init__(self, responses=None):
        self.responses = list(responses or [])
        self.requests = []

    def execute(self, request):
        self.requests.append(request)
        if self.responses:
            return self.responses.pop(0)
        return _successful(request.operation.value, None)


def _successful(operation, data):
    return OperationResult(
        operation=operation,
        success=True,
        phase="complete",
        data=data,
        request_id="request-1",
        elapsed_seconds=0.1,
        diagnostics=Diagnostics(
            worker_pid=90,
            excel_pid=91,
            cleanup=CleanupReport(
                pid=91,
                quit_requested=True,
                exited_gracefully=True,
            ),
        ),
    )


@pytest.mark.unit
def test_project_inspection_returns_typed_contract(tmp_path):
    screenshot = str(tmp_path / "Input.png")
    executor = RecordingExecutor([
        _successful(
            Operation.INSPECT.value,
            {
                "workbook_data": {"sheets": {"Input": {"cells": {}}}},
                "screenshots": {"Input": screenshot},
            },
        )
    ])
    project = Project.open(tmp_path / "book.xlsm", executor=executor)

    result = project.inspect(["Input"], cell_range="B2:C3", output_dir=tmp_path)

    assert result.success is True
    assert result.data.screenshots == {"Input": screenshot}
    assert result.artifacts[0].metadata["sheet"] == "Input"
    assert result.require_clean_shutdown().pid == 91
    json.dumps(result.to_dict())
    request = executor.requests[0]
    assert request.operation is Operation.INSPECT
    assert request.arguments["custom_range"] == "B2:C3"
    assert request.arguments["include_hidden_sheets"] is False
    assert request.arguments["include_rich_text"] is False


@pytest.mark.unit
def test_project_run_uses_typed_request_and_result(tmp_path):
    executor = RecordingExecutor([
        _successful(
            Operation.RUN.value,
            {"macro": "Calculate", "run_id": "run-1", "result": 42},
        )
    ])
    project = Project.open(tmp_path / "book.xlsm", executor=executor)

    result = project.run("Calculate", timeout=5, save=False)

    assert result.data.details["result"] == 42
    assert result.metadata == {"macro": "Calculate"}
    request = executor.requests[0]
    assert request.operation is Operation.RUN
    assert request.timeout == 5
    assert request.arguments["save_on_exit"] is False


@pytest.mark.unit
def test_project_run_preserves_executor_attempt_diagnostics(tmp_path):
    response = _successful(
        Operation.RUN.value,
        {"macro": "Calculate", "run_id": "run-1", "result": 42},
    )
    response = replace(
        response,
        attempt_count=2,
        diagnostics=Diagnostics(
            cleanup=response.diagnostics.cleanup,
            worker_pid=90,
            excel_pid=91,
            attempts=(
                AttemptDiagnostic(
                    attempt=1,
                    phase="worker_start",
                    error_code="worker_start_failed",
                    retryable=True,
                    retry_reason="worker_creation_failed",
                ),
                AttemptDiagnostic(attempt=2, phase="complete"),
            ),
        ),
    )
    project = Project.open(
        tmp_path / "book.xlsm",
        executor=RecordingExecutor([response]),
    )

    result = project.run("Calculate")

    assert result.attempt_count == 2
    assert result.diagnostics.attempts[0].retry_reason == "worker_creation_failed"
    assert result.to_dict()["diagnostics"]["attempts"][1]["phase"] == "complete"


@pytest.mark.unit
def test_project_lint_source_uses_resolved_source_path(tmp_path):
    source = tmp_path / "vba_source"
    source.mkdir()
    (source / "modTest.bas").write_text(
        "Option Explicit\nPublic Sub Test()\nEnd Sub\n", encoding="utf-8",
    )
    project = Project.open(tmp_path / "book.xlsm", source=source)

    result = project.lint_source()

    assert result.phase == "complete"
    assert result.metadata["issue_count"] >= 0
    json.dumps(result.to_dict())


@pytest.mark.unit
def test_project_lint_source_writes_baseline_and_returns_only_new_findings(tmp_path):
    source = tmp_path / "vba_source"
    source.mkdir()
    (source / "modTest.bas").write_text(
        "Option Explicit\nPublic Sub Test()\n    x = 1\nEnd Sub\n",
        encoding="utf-8",
    )
    baseline = tmp_path / "lint-baseline.json"
    project = Project.open(tmp_path / "book.xlsm", source=source)

    initial = project.lint_source(write_baseline=baseline)
    filtered = project.lint_source(baseline=baseline, new_only=True)

    assert initial.metadata["baseline_written"] == str(baseline.resolve())
    assert initial.artifacts[0].kind == "lint_baseline"
    assert filtered.success is True
    assert filtered.data == ()
    assert filtered.metadata["known_issue_count"] == initial.metadata["raw_issue_count"]


@pytest.mark.unit
def test_project_diff_propagates_vba_comparison_mode(tmp_path):
    executor = RecordingExecutor([_successful(Operation.DIFF.value, [])])
    project = Project.open(tmp_path / "book.xlsm", executor=executor)

    project.diff(comparison="text")

    assert executor.requests[0].arguments["comparison"] == "text"


@pytest.mark.unit
def test_workbook_lint_baseline_can_clear_only_analyzer_failure(tmp_path):
    from xlvbatools.analysis.filtering import write_lint_baseline
    from xlvbatools.analysis.issue import VBAIssue

    issue = VBAIssue(
        "IP001", "ERROR", "modules/modMain.bas", 20,
        "Variable 'FileCount' is undeclared", "Run",
    )
    baseline = tmp_path / "lint-baseline.json"
    write_lint_baseline(baseline, [issue])
    failed_lint = replace(
        _successful(Operation.LINT_WORKBOOK.value, [issue.to_dict()]),
        success=False,
        phase="lint_workbook",
        error=ErrorInfo("Static analysis found workbook errors", code="lint_failed"),
    )
    executor = RecordingExecutor([failed_lint])
    project = Project.open(tmp_path / "book.xlsm", executor=executor)

    result = project.lint_workbook(baseline=baseline, new_only=True)

    assert result.success is True
    assert result.phase == "complete"
    assert result.error is None
    assert result.data == ()
    assert result.diagnostics.cleanup.pid == 91


@pytest.mark.unit
def test_workbook_lint_filter_does_not_hide_operational_failure(tmp_path):
    failed = replace(
        _successful(Operation.LINT_WORKBOOK.value, []),
        success=False,
        phase="cleanup",
        error=ErrorInfo("Owned Excel remained running", code="cleanup_failed"),
    )
    project = Project.open(
        tmp_path / "book.xlsm", executor=RecordingExecutor([failed]),
    )

    result = project.lint_workbook(severities=["WARNING"])

    assert result.success is False
    assert result.error.code == "cleanup_failed"


@pytest.mark.unit
@pytest.mark.parametrize(
    ("method", "kwargs", "operation", "worker_data"),
    [
        ("extract", {}, Operation.EXTRACT, {"components": []}),
        ("inject", {"backup": False}, Operation.INJECT, []),
        ("diff", {}, Operation.DIFF, []),
        ("modify", {"cell": "A1", "value": 4}, Operation.MODIFY, True),
    ],
)
def test_excel_methods_share_one_executor(
    tmp_path, method, kwargs, operation, worker_data,
):
    executor = RecordingExecutor([_successful(operation.value, worker_data)])
    project = Project.open(tmp_path / "book.xlsm", executor=executor)

    result = getattr(project, method)(**kwargs)

    assert result.success is True
    assert result.diagnostics.worker_pid == 90
    assert result.require_clean_shutdown().pid == 91
    assert executor.requests[0].operation is operation
    assert executor.requests[0].retry_transient is (operation is Operation.MODIFY)


@pytest.mark.unit
def test_excel_methods_normalize_operation_specific_outputs(tmp_path):
    executor = RecordingExecutor([
        _successful(
            Operation.LIST_COMPONENTS.value,
            [{"name": "modMain", "type_code": 1, "type_name": "standard_module", "line_count": 4}],
        ),
        _successful(
            Operation.EXTRACT.value,
            {
                "workbook": "book.xlsm",
                "extracted_at": "2026-07-15T12:00:00",
                "components": [
                    {
                        "name": "modMain",
                        "type_code": 1,
                        "type_name": "standard_module",
                        "file": "modules/modMain.bas",
                    }
                ],
            },
        ),
        _successful(
            Operation.INJECT.value,
            [{"name": "modMain", "status": "injected"}],
        ),
        _successful(
            Operation.DIFF.value,
            [{"name": "modMain", "status": "identical"}],
        ),
        _successful(Operation.MODIFY.value, True),
    ])
    project = Project.open(tmp_path / "book.xlsm", executor=executor)

    assert project.list_components().data[0].name == "modMain"
    assert project.extract().data.components[0].file == "modules/modMain.bas"
    assert project.inject().data.changes[0].status == "injected"
    assert project.diff().data[0].status == "identical"
    assert project.modify(cell="A1", value=4).data.action == "set_value"


@pytest.mark.unit
def test_project_from_nested_directory_resolves_config_paths(tmp_path):
    (tmp_path / "xlvbatools.toml").write_text(
        "[xlvbatools]\n"
        'workbook = "workbook/book.xlsm"\n'
        'vba_source = "workbook/vba_source"\n'
        'snapshots_dir = "workbook/snapshots"\n'
        'log_dir = "logs"\n',
        encoding="utf-8",
    )
    nested = tmp_path / "tools" / "nested"
    nested.mkdir(parents=True)

    project = Project.from_config(nested, executor=RecordingExecutor())

    assert project.workbook == (tmp_path / "workbook" / "book.xlsm").resolve()
    assert project.source == (tmp_path / "workbook" / "vba_source").resolve()
    assert project.settings.snapshots == (tmp_path / "workbook" / "snapshots").resolve()


@pytest.mark.unit
def test_public_api_is_small_and_does_not_import_win32com():
    import sys
    import xlvbatools

    for name in tuple(sys.modules):
        if name.startswith("win32com"):
            del sys.modules[name]

    from xlvbatools import OperationRequest, OperationResult, Project, VBAIssue

    assert Project is not None
    assert OperationRequest is not None
    assert OperationResult is not None
    assert VBAIssue is not None
    assert "XlvbaProject" not in xlvbatools.__all__
    assert "ExcelSession" not in xlvbatools.__all__
    assert "lint_files" not in xlvbatools.__all__
    assert "win32com" not in sys.modules


@pytest.mark.excel
@pytest.mark.smoke
def test_project_inspection_reports_clean_owned_process(minimal_workbook):
    result = Project.open(minimal_workbook).inspect(
        ["Sheet1"], include_data=True, include_screenshots=False, timeout=60,
    )

    assert result.success is True, result.to_dict()
    assert result.schema_version == "1.3"
    assert result.data.workbook_data["sheets"]["Sheet1"]
    assert result.diagnostics.cleanup.is_clean, result.to_dict()
    assert result.require_clean_shutdown().still_running is False


@pytest.mark.excel
@pytest.mark.smoke
def test_project_macro_reports_clean_owned_process(runtime_error_workbook):
    result = Project.open(runtime_error_workbook).run(
        "CompleteNormally", timeout=60, save=False,
    )

    assert result.success is True, json.dumps(result.to_dict(), indent=2)
    assert result.data.macro == "CompleteNormally"
    assert result.require_clean_shutdown().still_running is False


@pytest.mark.excel
def test_project_vba_round_trip_uses_clean_sequential_workers(
    runtime_error_workbook, tmp_path,
):
    source = tmp_path / "isolated_source"
    project = Project.open(runtime_error_workbook, source=source)

    extracted = project.extract(timeout=90)
    assert extracted.success is True, extracted.to_dict()
    assert extracted.data.components
    assert extracted.require_clean_shutdown().still_running is False

    compared = project.diff(timeout=90)
    assert compared.success is True, compared.to_dict()
    assert all(item.status == "identical" for item in compared.data)

    injected = project.inject(backup=False, timeout=90)
    assert injected.success is True, injected.to_dict()
    assert all(item.status == "injected" for item in injected.data.changes)

    linted = project.lint_workbook(compile_test=False, timeout=90)
    assert linted.require_clean_shutdown().still_running is False

    executed = project.run("CompleteNormally", timeout=90, save=False)
    assert executed.success is True, executed.to_dict()
    assert executed.diagnostics.cleanup.is_clean, executed.to_dict()
    assert executed.require_clean_shutdown().still_running is False


@pytest.mark.excel
def test_live_diff_classifies_vba_case_and_spacing_as_equivalent(
    runtime_error_workbook, tmp_path,
):
    source = tmp_path / "case_source"
    project = Project.open(runtime_error_workbook, source=source)
    extracted = project.extract(timeout=90)
    assert extracted.success is True, extracted.to_dict()
    component = next(
        item for item in extracted.data.components
        if item.name == "modReliabilityTest"
    )
    component_path = source / component.file
    original = component_path.read_text(encoding="utf-8")
    changed = original.replace(
        "Public Sub CompleteNormally()",
        "PUBLIC  SUB  completenormally ( )",
    )
    assert changed != original
    component_path.write_text(changed, encoding="utf-8")

    semantic = project.diff(comparison="vba", timeout=90)
    semantic_component = next(
        item for item in semantic.data if item.name == "modReliabilityTest"
    )
    assert semantic.success is True, semantic.to_dict()
    assert semantic_component.status == "equivalent"
    assert semantic_component.equivalence == "vba_token_equivalent"
    assert semantic.require_clean_shutdown().still_running is False

    raw = project.diff(comparison="text", timeout=90)
    raw_component = next(
        item for item in raw.data if item.name == "modReliabilityTest"
    )
    assert raw.success is True, raw.to_dict()
    assert raw_component.status == "modified"
    assert raw_component.lines_added > 0
    assert raw_component.lines_removed > 0
    assert raw.require_clean_shutdown().still_running is False


@pytest.mark.excel
def test_live_lint_conclusively_compiles_valid_minimal_workbook(minimal_workbook):
    result = Project.open(minimal_workbook).lint_workbook(
        compile_test=True,
        timeout=90,
    )

    assert result.success is True, result.to_dict()
    assert not [issue for issue in result.data if issue.rule_id == "CT001"]
    assert result.require_clean_shutdown().still_running is False


@pytest.mark.excel
def test_live_lint_reports_compile_error_location_and_closes_cleanly(
    compile_error_workbook,
):
    result = Project.open(compile_error_workbook).lint_workbook(
        compile_test=True,
        timeout=90,
    )

    compile_issues = [issue for issue in result.data if issue.rule_id == "CT001"]
    assert result.success is False, result.to_dict()
    assert result.error.code == "lint_failed"
    assert any(
        issue.module == "modCompileFailure" and issue.line_num == 3
        for issue in compile_issues
    ), result.to_dict()
    assert any(
        event.get("type") == "compile_error"
        for event in result.diagnostics.dialog_events
    ), result.to_dict()
    cleanup = result.require_clean_shutdown()
    assert cleanup.exited_gracefully is True


@pytest.mark.excel
def test_source_operations_never_execute_workbook_startup_code(
    startup_event_workbook, tmp_path,
):
    """Extract/diff/inject/lint open untrusted workbooks with code disabled."""
    workbook, marker = startup_event_workbook
    source = tmp_path / "startup_safe_source"
    project = Project.open(workbook, source=source)

    extracted = project.extract(timeout=90)
    assert extracted.success is True, extracted.to_dict()
    assert not marker.exists()

    compared = project.diff(timeout=90)
    assert compared.success is True, compared.to_dict()
    assert not marker.exists()

    injected = project.inject(backup=False, timeout=90)
    assert injected.success is True, injected.to_dict()
    assert not marker.exists()

    linted = project.lint_workbook(compile_test=True, timeout=90)
    assert linted.require_clean_shutdown().still_running is False
    assert not marker.exists()


@pytest.mark.excel
def test_live_lint_rejects_duplicate_declaration_and_closes_cleanly(
    duplicate_declaration_workbook, tmp_path,
):
    baseline = tmp_path / "duplicate-baseline.json"
    result = Project.open(duplicate_declaration_workbook).lint_workbook(
        compile_test=True,
        write_baseline=baseline,
        timeout=90,
    )

    rule_ids = {issue.rule_id for issue in result.data}
    assert result.success is False, result.to_dict()
    assert "DV001" in rule_ids
    assert "CT001" in rule_ids
    assert result.error.code == "lint_failed"
    assert result.metadata["baseline_written"] == str(baseline.resolve())
    cleanup = result.require_clean_shutdown()
    assert cleanup.exited_gracefully is True

    new_only = Project.open(duplicate_declaration_workbook).lint_workbook(
        compile_test=True,
        baseline=baseline,
        new_only=True,
        timeout=90,
    )
    assert new_only.success is True, new_only.to_dict()
    assert new_only.data == ()
    assert new_only.metadata["known_issue_count"] == len(result.data)
    new_cleanup = new_only.require_clean_shutdown()
    assert new_cleanup.exited_gracefully is True


@pytest.mark.excel
def test_project_modify_then_inspect_across_isolated_workers(minimal_workbook):
    project = Project.open(minimal_workbook)
    modified = project.modify(sheet="Sheet1", cell="B2", value=73, timeout=90)
    assert modified.success is True, modified.to_dict()
    assert modified.require_clean_shutdown().still_running is False

    inspected = project.inspect(
        ["Sheet1"], cell_range="B2", include_screenshots=False, timeout=90,
    )
    assert inspected.success is True, inspected.to_dict()
    cell = inspected.data.workbook_data["sheets"]["Sheet1"]["cells"]["B2"]
    assert cell["value"] == 73
