"""Tests for the v1 Project API and typed executor boundary."""

from __future__ import annotations

import json

import pytest

from xlvbatools import (
    CleanupReport,
    Diagnostics,
    Operation,
    OperationResult,
    Project,
)


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


@pytest.mark.com
@pytest.mark.integration
def test_project_inspection_reports_clean_owned_process(minimal_workbook):
    result = Project.open(minimal_workbook).inspect(
        ["Sheet1"], include_data=True, include_screenshots=False, timeout=60,
    )

    assert result.success is True, result.to_dict()
    assert result.schema_version == "1.0"
    assert result.data.workbook_data["sheets"]["Sheet1"]
    assert result.diagnostics.cleanup.is_clean, result.to_dict()
    assert result.require_clean_shutdown().still_running is False


@pytest.mark.com
@pytest.mark.integration
def test_project_macro_reports_clean_owned_process(runtime_error_workbook):
    result = Project.open(runtime_error_workbook).run(
        "CompleteNormally", timeout=60, save=False,
    )

    assert result.success is True, result.to_dict()
    assert result.data.macro == "CompleteNormally"
    assert result.require_clean_shutdown().still_running is False


@pytest.mark.com
@pytest.mark.e2e
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
    assert executed.require_clean_shutdown().still_running is False


@pytest.mark.com
@pytest.mark.e2e
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
